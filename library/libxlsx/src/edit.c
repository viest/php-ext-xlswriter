/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

#include <ctype.h>
#include <stdarg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "libxlsx/edit.h"
#include "libxlsx/source_package.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/workbook.h"
#include "libxlsx/xmlwriter.h"

#include "xlsx_private.h"

typedef enum {
    LXLSX_EDIT_CHANGE_NUMBER,
    LXLSX_EDIT_CHANGE_STRING,
    LXLSX_EDIT_CHANGE_BOOLEAN,
    LXLSX_EDIT_CHANGE_FORMULA
} lxlsx_edit_change_type;

typedef struct {
    char *target;
    lxlsx_row_t row;
    lxlsx_col_t col;
    lxlsx_edit_change_type type;
    double number;
    char *string;
    int boolean;
    char *formula;
    char *cached_result;
} lxlsx_edit_change;

struct lxlsx_edit_session {
    lxlsx_source_package *package;
    lxlsx_reader_workbook *workbook;
    lxlsx_edit_change *changes;
    size_t change_count;
    size_t change_cap;
};

typedef struct {
    char *data;
    size_t len;
    size_t cap;
} lxlsx_edit_buf;

typedef struct {
    char *target;
    size_t entry_index;
    unsigned char *xml;
    size_t xml_len;
    const lxlsx_edit_change **changes;
    size_t change_count;
    size_t change_cap;
} lxlsx_dirty_sheet;

#define LXLSX_EDIT_CELL_CLOSE       "</c>"
#define LXLSX_EDIT_ROW_CLOSE        "</row>"

typedef struct {
    const char *start;
    const char *end;
    const char *name;
    size_t name_len;
    int is_end;
    int is_self_closing;
} lxlsx_edit_xml_tag;

typedef struct {
    const lxlsx_edit_change *change;
    size_t order;
    char *ref;
    char *row_ref;
} lxlsx_edit_prepared_change;

static int buf_reserve(lxlsx_edit_buf *buf, size_t need)
{
    char *next;
    size_t cap;
    if (need <= buf->cap)
        return 0;
    cap = buf->cap ? buf->cap : 128;
    while (cap < need)
        cap *= 2;
    next = (char *)realloc(buf->data, cap);
    if (!next)
        return -1;
    buf->data = next;
    buf->cap = cap;
    return 0;
}

static int buf_append(lxlsx_edit_buf *buf, const char *data, size_t len)
{
    if (buf_reserve(buf, buf->len + len + 1) != 0)
        return -1;
    memcpy(buf->data + buf->len, data, len);
    buf->len += len;
    buf->data[buf->len] = 0;
    return 0;
}

static int buf_append_s(lxlsx_edit_buf *buf, const char *str)
{
    return buf_append(buf, str, strlen(str));
}

static int buf_append_matching_end_tag(lxlsx_edit_buf *buf,
                                       const char *tag_start,
                                       const char *tag_end)
{
    const char *name_start = tag_start + 1;
    const char *name_end = name_start;

    if (name_start >= tag_end || *name_start == '/' ||
        *name_start == '!' || *name_start == '?')
        return -1;

    while (name_end < tag_end && !isspace((unsigned char)*name_end) &&
           *name_end != '>' && *name_end != '/')
        name_end++;

    if (name_start == name_end)
        return -1;

    if (buf_append_s(buf, "</") != 0 ||
        buf_append(buf, name_start, (size_t)(name_end - name_start)) != 0 ||
        buf_append_s(buf, ">") != 0)
        return -1;

    return 0;
}

static int buf_appendf(lxlsx_edit_buf *buf, const char *fmt, ...)
{
    va_list ap;
    va_list ap2;
    int n;

    va_start(ap, fmt);
    va_copy(ap2, ap);
    n = vsnprintf(NULL, 0, fmt, ap);
    va_end(ap);
    if (n < 0) {
        va_end(ap2);
        return -1;
    }
    if (buf_reserve(buf, buf->len + (size_t)n + 1) != 0) {
        va_end(ap2);
        return -1;
    }
    vsnprintf(buf->data + buf->len, (size_t)n + 1, fmt, ap2);
    va_end(ap2);
    buf->len += (size_t)n;
    return 0;
}

static char *dup_range(const char *start, size_t len)
{
    char *out = (char *)malloc(len + 1);
    if (!out)
        return NULL;
    memcpy(out, start, len);
    out[len] = 0;
    return out;
}

static int edit_buf_write_callback(void *userdata, const char *data, size_t len)
{
    return buf_append((lxlsx_edit_buf *)userdata, data, len);
}

static int append_xml_escaped(lxlsx_edit_buf *buf, const char *str)
{
    if (!str)
        return 0;
    return lxlsx_xml_escape_data_write(str, edit_buf_write_callback, buf);
}

static int needs_xml_space_preserve(const char *str)
{
    size_t len;
    if (!str || !*str)
        return 0;
    len = strlen(str);
    return isspace((unsigned char)str[0]) ||
           isspace((unsigned char)str[len - 1]);
}

static char *make_cell_ref(lxlsx_row_t row, lxlsx_col_t col)
{
    char col_name[16];
    char ref[32];
    unsigned int n = (unsigned int)col + 1;
    int pos = 0;
    char tmp[16];

    do {
        n--;
        tmp[pos++] = (char)('A' + (n % 26));
        n /= 26;
    } while (n > 0 && pos < (int)sizeof(tmp));

    n = 0;
    while (pos > 0)
        col_name[n++] = tmp[--pos];
    col_name[n] = 0;

    snprintf(ref, sizeof(ref), "%s%u", col_name, (unsigned int)row + 1);
    return strdup(ref);
}

static char *make_row_ref(lxlsx_row_t row)
{
    char ref[32];
    snprintf(ref, sizeof(ref), "%u", (unsigned int)row + 1);
    return strdup(ref);
}

static int parse_row_ref(const char *ref, lxlsx_row_t *out_row)
{
    unsigned long row = 0;
    const char *p = ref;

    if (!ref || !*ref)
        return 0;

    while (*p) {
        if (*p < '0' || *p > '9')
            return 0;
        row = row * 10 + (unsigned long)(*p - '0');
        if (row > LXLSX_ROW_MAX)
            return 0;
        p++;
    }

    if (row == 0)
        return 0;

    if (out_row)
        *out_row = (lxlsx_row_t)(row - 1);
    return 1;
}

static int parse_cell_ref(const char *ref,
                          lxlsx_row_t *out_row,
                          lxlsx_col_t *out_col)
{
    unsigned long row = 0;
    unsigned long col = 0;
    const char *p = ref;

    if (!ref || !*ref)
        return 0;

    if (*p == '$')
        p++;

    while ((*p >= 'A' && *p <= 'Z') || (*p >= 'a' && *p <= 'z')) {
        col = col * 26 + (unsigned long)(toupper((unsigned char)*p) - 'A' + 1);
        if (col > LXLSX_COL_MAX)
            return 0;
        p++;
    }

    if (col == 0)
        return 0;

    if (*p == '$')
        p++;

    while (*p >= '0' && *p <= '9') {
        row = row * 10 + (unsigned long)(*p - '0');
        if (row > LXLSX_ROW_MAX)
            return 0;
        p++;
    }

    if (row == 0 || *p != 0)
        return 0;

    if (out_row)
        *out_row = (lxlsx_row_t)(row - 1);
    if (out_col)
        *out_col = (lxlsx_col_t)(col - 1);
    return 1;
}

static int xml_local_name_eq(const char *start, size_t len, const char *name)
{
    const char *local = start;
    size_t local_len = len;
    size_t i;

    for (i = 0; i < len; i++) {
        if (start[i] == ':') {
            local = start + i + 1;
            local_len = len - i - 1;
        }
    }

    return strlen(name) == local_len && strncmp(local, name, local_len) == 0;
}

static char *extract_attr(const char *tag_start, const char *tag_end,
                          const char *name)
{
    const char *p = tag_start + 1;

    if (p < tag_end && *p == '/')
        p++;
    while (p < tag_end && !isspace((unsigned char)*p) &&
           *p != '>' && *p != '/')
        p++;

    while (p < tag_end) {
        const char *attr_start;
        const char *attr_end;
        const char *value_start;
        char quote;

        while (p < tag_end && isspace((unsigned char)*p))
            p++;
        if (p >= tag_end || *p == '/' || *p == '>')
            break;

        attr_start = p;
        while (p < tag_end && *p != '=' && !isspace((unsigned char)*p) &&
               *p != '>' && *p != '/')
            p++;
        attr_end = p;
        while (p < tag_end && isspace((unsigned char)*p))
            p++;
        if (p >= tag_end || *p != '=')
            break;
        p++;
        while (p < tag_end && isspace((unsigned char)*p))
            p++;
        if (p >= tag_end || (*p != '"' && *p != '\''))
            break;
        quote = *p++;
        value_start = p;
        while (p < tag_end && *p != quote)
            p++;
        if (p >= tag_end)
            break;
        if (xml_local_name_eq(attr_start, (size_t)(attr_end - attr_start), name))
            return dup_range(value_start, (size_t)(p - value_start));
        p++;
    }

    return NULL;
}

static int is_self_closing(const char *tag_start, const char *tag_end)
{
    const char *p = tag_end;
    (void)tag_start;
    while (p > tag_start) {
        p--;
        if (isspace((unsigned char)*p))
            continue;
        return *p == '/';
    }
    return 0;
}

static const char *find_bytes(const char *start, const char *limit,
                              const char *needle, size_t needle_len)
{
    const char *p;

    if (needle_len == 0)
        return start;

    for (p = start; p <= limit && (size_t)(limit - p) >= needle_len; p++) {
        if (memcmp(p, needle, needle_len) == 0)
            return p;
    }

    return NULL;
}

static const char *find_tag_end(const char *start, const char *limit)
{
    const char *p;
    char quote = 0;

    for (p = start; p < limit; p++) {
        if (quote) {
            if (*p == quote)
                quote = 0;
            continue;
        }
        if (*p == '"' || *p == '\'') {
            quote = *p;
            continue;
        }
        if (*p == '>')
            return p;
    }

    return NULL;
}

static const char *skip_special_tag(const char *start, const char *limit)
{
    const char *end;

    if ((size_t)(limit - start) >= 4 &&
        memcmp(start, "<!--", 4) == 0) {
        end = find_bytes(start + 4, limit, "-->", 3);
        return end ? end + 3 : NULL;
    }

    if ((size_t)(limit - start) >= 9 &&
        memcmp(start, "<![CDATA[", 9) == 0) {
        end = find_bytes(start + 9, limit, "]]>", 3);
        return end ? end + 3 : NULL;
    }

    if ((size_t)(limit - start) >= 2 &&
        memcmp(start, "<?", 2) == 0) {
        end = find_bytes(start + 2, limit, "?>", 2);
        return end ? end + 2 : NULL;
    }

    end = find_tag_end(start, limit);
    return end ? end + 1 : NULL;
}

static int parse_xml_tag(const char *tag_start, const char *tag_end,
                         lxlsx_edit_xml_tag *tag)
{
    const char *p = tag_start + 1;
    const char *name_start;
    const char *name_end;
    const char *local_name;

    memset(tag, 0, sizeof(*tag));
    if (p >= tag_end)
        return 0;
    if (*p == '!' || *p == '?')
        return 0;
    if (*p == '/') {
        tag->is_end = 1;
        p++;
    }
    if (p >= tag_end)
        return 0;

    name_start = p;
    while (p < tag_end && !isspace((unsigned char)*p) &&
           *p != '>' && *p != '/')
        p++;
    name_end = p;
    if (name_start == name_end)
        return 0;

    local_name = name_start;
    while (name_end > name_start) {
        name_end--;
        if (*name_end == ':') {
            local_name = name_end + 1;
            break;
        }
    }

    tag->start = tag_start;
    tag->end = tag_end;
    tag->name = local_name;
    tag->name_len = (size_t)(p - local_name);
    tag->is_self_closing = !tag->is_end && is_self_closing(tag_start, tag_end);
    return tag->name_len > 0;
}

static const char *xml_next_tag(const char *start, const char *limit,
                                lxlsx_edit_xml_tag *tag)
{
    const char *p = start;

    while (p < limit) {
        const char *lt;
        const char *end;

        while (p < limit && *p != '<')
            p++;
        if (p >= limit)
            return NULL;
        lt = p;

        if ((size_t)(limit - lt) >= 2 &&
            (lt[1] == '!' || lt[1] == '?')) {
            p = skip_special_tag(lt, limit);
            if (!p)
                return NULL;
            continue;
        }

        end = find_tag_end(lt, limit);
        if (!end)
            return NULL;
        if (parse_xml_tag(lt, end, tag))
            return lt;
        p = end + 1;
    }

    return NULL;
}

static int tag_name_is(const lxlsx_edit_xml_tag *tag, const char *name)
{
    return strlen(name) == tag->name_len &&
           strncmp(tag->name, name, tag->name_len) == 0;
}

static int find_matching_end_tag(const char *start, const char *limit,
                                 const char *name,
                                 lxlsx_edit_xml_tag *out)
{
    lxlsx_edit_xml_tag tag;
    const char *p = start;
    size_t depth = 0;

    while (xml_next_tag(p, limit, &tag)) {
        if (tag_name_is(&tag, name)) {
            if (tag.is_end) {
                if (depth == 0) {
                    if (out)
                        *out = tag;
                    return 1;
                }
                depth--;
            } else if (!tag.is_self_closing) {
                depth++;
            }
        }
        p = tag.end + 1;
    }

    return 0;
}

static int find_sheet_data(const char *xml, size_t len,
                           const char **sheet_start,
                           const char **sheet_end,
                           const char **tag_end,
                           const char **close_start)
{
    const char *limit = xml + len;
    const char *p = xml;
    lxlsx_edit_xml_tag tag;

    while (xml_next_tag(p, limit, &tag)) {
        lxlsx_edit_xml_tag close_tag;

        p = tag.end + 1;
        if (tag.is_end || !tag_name_is(&tag, "sheetData"))
            continue;

        *sheet_start = tag.start;
        *tag_end = tag.end;
        if (close_start)
            *close_start = NULL;
        if (tag.is_self_closing) {
            *sheet_end = tag.end + 1;
            return 1;
        }
        if (!find_matching_end_tag(tag.end + 1, limit, "sheetData", &close_tag))
            return 0;
        *sheet_end = close_tag.end + 1;
        if (close_start)
            *close_start = close_tag.start;
        return 1;
    }

    return 0;
}

static int looks_numeric(const char *str)
{
    char *end = NULL;
    if (!str || !*str)
        return 0;
    (void)strtod(str, &end);
    return end && *end == 0;
}

static char *build_cell_xml(const lxlsx_edit_change *change,
                            const char *ref,
                            const char *style)
{
    lxlsx_edit_buf buf = {0};
    char number[64];

    if (buf_appendf(&buf, "<c r=\"%s\"", ref) != 0)
        goto fail;
    if (style && *style) {
        if (buf_appendf(&buf, " s=\"%s\"", style) != 0)
            goto fail;
    }

    if (change->type == LXLSX_EDIT_CHANGE_STRING) {
        if (buf_append_s(&buf, " t=\"inlineStr\"") != 0)
            goto fail;
    } else if (change->type == LXLSX_EDIT_CHANGE_BOOLEAN) {
        if (buf_append_s(&buf, " t=\"b\"") != 0)
            goto fail;
    }

    if (change->type == LXLSX_EDIT_CHANGE_FORMULA &&
        change->cached_result && !looks_numeric(change->cached_result)) {
        if (buf_append_s(&buf, " t=\"str\"") != 0)
            goto fail;
    }

    if (buf_append_s(&buf, ">") != 0)
        goto fail;

    if (change->type == LXLSX_EDIT_CHANGE_NUMBER) {
        snprintf(number, sizeof(number), "%.17g", change->number);
        if (buf_append_s(&buf, "<v>") != 0 ||
            buf_append_s(&buf, number) != 0 ||
            buf_append_s(&buf, "</v>") != 0)
            goto fail;
    } else if (change->type == LXLSX_EDIT_CHANGE_STRING) {
        const char *string = change->string ? change->string : "";
        if (buf_append_s(&buf, "<is><t") != 0)
            goto fail;
        if (needs_xml_space_preserve(string) &&
            buf_append_s(&buf, " xml:space=\"preserve\"") != 0) {
            goto fail;
        }
        if (buf_append_s(&buf, ">") != 0 ||
            append_xml_escaped(&buf, string) != 0 ||
            buf_append_s(&buf, "</t></is>") != 0) {
            goto fail;
        }
    } else if (change->type == LXLSX_EDIT_CHANGE_BOOLEAN) {
        if (buf_appendf(&buf, "<v>%d</v>", change->boolean ? 1 : 0) != 0)
            goto fail;
    } else {
        const char *formula = change->formula;
        const char *cached = change->cached_result ? change->cached_result : "0";
        if (formula && formula[0] == '=')
            formula++;
        if (buf_append_s(&buf, "<f>") != 0 ||
            append_xml_escaped(&buf, formula ? formula : "") != 0 ||
            buf_append_s(&buf, "</f><v>") != 0 ||
            append_xml_escaped(&buf, cached) != 0 ||
            buf_append_s(&buf, "</v>") != 0)
            goto fail;
    }

    if (buf_append_s(&buf, LXLSX_EDIT_CELL_CLOSE) != 0)
        goto fail;
    return buf.data;

fail:
    free(buf.data);
    return NULL;
}

static int prepared_change_cmp(const void *a, const void *b)
{
    const lxlsx_edit_prepared_change *left =
        (const lxlsx_edit_prepared_change *)a;
    const lxlsx_edit_prepared_change *right =
        (const lxlsx_edit_prepared_change *)b;

    if (left->change->row != right->change->row)
        return left->change->row < right->change->row ? -1 : 1;
    if (left->change->col != right->change->col)
        return left->change->col < right->change->col ? -1 : 1;
    if (left->order != right->order)
        return left->order < right->order ? -1 : 1;
    return 0;
}

static void free_prepared_changes(lxlsx_edit_prepared_change *changes,
                                  size_t count)
{
    size_t i;
    if (!changes)
        return;
    for (i = 0; i < count; i++) {
        free(changes[i].ref);
        free(changes[i].row_ref);
    }
    free(changes);
}

static lxlsx_error prepare_sheet_changes(const lxlsx_edit_change **changes,
                                         size_t change_count,
                                         lxlsx_edit_prepared_change **out,
                                         size_t *out_count)
{
    lxlsx_edit_prepared_change *prepared = NULL;
    lxlsx_edit_prepared_change *kept = NULL;
    size_t i;
    size_t kept_count = 0;

    *out = NULL;
    *out_count = 0;

    prepared = (lxlsx_edit_prepared_change *)calloc(change_count,
                                                    sizeof(*prepared));
    if (!prepared)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    for (i = 0; i < change_count; i++) {
        prepared[i].change = changes[i];
        prepared[i].order = i;
        prepared[i].ref = make_cell_ref(changes[i]->row, changes[i]->col);
        prepared[i].row_ref = make_row_ref(changes[i]->row);
        if (!prepared[i].ref || !prepared[i].row_ref) {
            free_prepared_changes(prepared, change_count);
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    if (change_count > 1)
        qsort(prepared, change_count, sizeof(*prepared), prepared_change_cmp);

    kept = (lxlsx_edit_prepared_change *)calloc(change_count, sizeof(*kept));
    if (!kept) {
        free_prepared_changes(prepared, change_count);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    i = 0;
    while (i < change_count) {
        size_t j = i + 1;
        size_t keep;
        size_t k;

        while (j < change_count &&
               prepared[j].change->row == prepared[i].change->row &&
               prepared[j].change->col == prepared[i].change->col) {
            j++;
        }

        keep = j - 1;
        kept[kept_count++] = prepared[keep];
        prepared[keep].ref = NULL;
        prepared[keep].row_ref = NULL;

        for (k = i; k < j; k++) {
            free(prepared[k].ref);
            free(prepared[k].row_ref);
        }
        i = j;
    }

    free(prepared);
    *out = kept;
    *out_count = kept_count;
    return LXLSX_NO_ERROR;
}

static const char *self_closing_prefix_end(const char *tag_start,
                                           const char *tag_end)
{
    const char *p = tag_end;

    while (p > tag_start) {
        p--;
        if (isspace((unsigned char)*p))
            continue;
        if (*p == '/')
            return p;
        break;
    }

    return tag_end;
}

static lxlsx_error append_cells(lxlsx_edit_buf *buf,
                                const lxlsx_edit_prepared_change *changes,
                                size_t count)
{
    size_t i;

    for (i = 0; i < count; i++) {
        char *cell_xml = build_cell_xml(changes[i].change, changes[i].ref, NULL);
        if (!cell_xml)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        if (buf_append_s(buf, cell_xml) != 0) {
            free(cell_xml);
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
        free(cell_xml);
    }

    return LXLSX_NO_ERROR;
}

static lxlsx_error append_change_cell(lxlsx_edit_buf *buf,
                                      const lxlsx_edit_prepared_change *change,
                                      const char *style)
{
    char *cell_xml = build_cell_xml(change->change, change->ref, style);
    if (!cell_xml)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    if (buf_append_s(buf, cell_xml) != 0) {
        free(cell_xml);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    free(cell_xml);
    return LXLSX_NO_ERROR;
}

static lxlsx_error append_cells_before_col(
    lxlsx_edit_buf *buf,
    const lxlsx_edit_prepared_change *changes,
    size_t count,
    size_t *index,
    lxlsx_col_t before_col)
{
    while (*index < count && changes[*index].change->col < before_col) {
        lxlsx_error err = append_change_cell(buf, &changes[*index], NULL);
        if (err != LXLSX_NO_ERROR)
            return err;
        (*index)++;
    }

    return LXLSX_NO_ERROR;
}

static lxlsx_error append_remaining_cells(
    lxlsx_edit_buf *buf,
    const lxlsx_edit_prepared_change *changes,
    size_t count,
    size_t *index)
{
    while (*index < count) {
        lxlsx_error err = append_change_cell(buf, &changes[*index], NULL);
        if (err != LXLSX_NO_ERROR)
            return err;
        (*index)++;
    }

    return LXLSX_NO_ERROR;
}

static char *build_new_row_xml(const char *row_ref,
                               const lxlsx_edit_prepared_change *changes,
                               size_t count)
{
    lxlsx_edit_buf row = {0};

    if (buf_appendf(&row, "<row r=\"%s\">", row_ref) != 0)
        goto fail;
    if (append_cells(&row, changes, count) != LXLSX_NO_ERROR)
        goto fail;
    if (buf_append_s(&row, LXLSX_EDIT_ROW_CLOSE) != 0)
        goto fail;

    return row.data;

fail:
    free(row.data);
    return NULL;
}

static char *build_existing_row_xml(const char *row_start,
                                    const char *row_end,
                                    const char *row_tag_end,
                                    const char *row_close_start,
                                    const lxlsx_edit_prepared_change *changes,
                                    size_t count)
{
    lxlsx_edit_buf row = {0};
    size_t change_index = 0;

    if (!row_close_start) {
        const char *prefix_end = self_closing_prefix_end(row_start, row_tag_end);
        if (buf_append(&row, row_start, (size_t)(prefix_end - row_start)) != 0 ||
            buf_append_s(&row, ">") != 0)
            goto fail;
        if (append_remaining_cells(&row, changes, count, &change_index)
            != LXLSX_NO_ERROR)
            goto fail;
        if (buf_append_matching_end_tag(&row, row_start, row_tag_end) != 0)
            goto fail;
    } else {
        const char *pos = row_start;
        const char *p = row_tag_end + 1;
        const char *scan_limit = row_close_start;
        lxlsx_edit_xml_tag tag;

        while (xml_next_tag(p, scan_limit, &tag)) {
            const char *cell_end;
            char *ref = NULL;
            lxlsx_row_t cell_row = 0;
            lxlsx_col_t cell_col = 0;
            int valid_ref = 0;

            p = tag.end + 1;
            if (tag.is_end || !tag_name_is(&tag, "c"))
                continue;

            if (tag.is_self_closing) {
                cell_end = tag.end + 1;
            } else {
                lxlsx_edit_xml_tag close_tag;
                if (!find_matching_end_tag(tag.end + 1, row_close_start,
                                           "c", &close_tag))
                    goto fail;
                cell_end = close_tag.end + 1;
            }

            ref = extract_attr(tag.start, tag.end, "r");
            valid_ref = parse_cell_ref(ref, &cell_row, &cell_col);
            free(ref);

            if (!valid_ref || cell_row != changes[0].change->row) {
                p = cell_end;
                continue;
            }

            if (buf_append(&row, pos, (size_t)(tag.start - pos)) != 0)
                goto fail;

            if (append_cells_before_col(&row, changes, count, &change_index,
                                        cell_col) != LXLSX_NO_ERROR)
                goto fail;

            if (change_index < count &&
                changes[change_index].change->col == cell_col) {
                char *style = extract_attr(tag.start, tag.end, "s");
                lxlsx_error err = append_change_cell(
                    &row, &changes[change_index], style);
                free(style);
                if (err != LXLSX_NO_ERROR)
                    goto fail;
                change_index++;
            } else {
                if (buf_append(&row, tag.start,
                               (size_t)(cell_end - tag.start)) != 0)
                    goto fail;
            }

            pos = cell_end;
            p = cell_end;
        }

        if (buf_append(&row, pos, (size_t)(row_close_start - pos)) != 0)
            goto fail;

        if (append_remaining_cells(&row, changes, count, &change_index)
            != LXLSX_NO_ERROR)
            goto fail;

        if (buf_append(&row, row_close_start,
                       (size_t)(row_end - row_close_start)) != 0)
            goto fail;
    }

    return row.data;

fail:
    free(row.data);
    return NULL;
}

static lxlsx_error append_missing_rows_until(
    lxlsx_edit_buf *out,
    const lxlsx_edit_prepared_change *prepared,
    size_t prepared_count,
    size_t *index,
    lxlsx_row_t before_row,
    int bounded)
{
    while (*index < prepared_count &&
           (!bounded || prepared[*index].change->row < before_row)) {
        lxlsx_row_t row = prepared[*index].change->row;
        const char *row_ref = prepared[*index].row_ref;
        char *row_xml;
        size_t next = *index + 1;

        while (next < prepared_count && prepared[next].change->row == row)
            next++;

        row_xml = build_new_row_xml(row_ref, prepared + *index,
                                    next - *index);
        if (!row_xml)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        if (buf_append_s(out, row_xml) != 0) {
            free(row_xml);
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
        free(row_xml);
        *index = next;
    }

    return LXLSX_NO_ERROR;
}

static lxlsx_error patch_xml_with_changes(unsigned char **xml, size_t *xml_len,
                                          const lxlsx_edit_change **changes,
                                          size_t change_count)
{
    const char *xml_start = (const char *)*xml;
    const char *xml_end = xml_start + *xml_len;
    lxlsx_edit_prepared_change *prepared = NULL;
    size_t prepared_count = 0;
    lxlsx_edit_buf out = {0};
    const char *sheet_start = NULL;
    const char *sheet_end = NULL;
    const char *sheet_tag_end = NULL;
    const char *sheet_close_start = NULL;
    lxlsx_error err;
    size_t change_index = 0;

    err = prepare_sheet_changes(changes, change_count, &prepared,
                                &prepared_count);
    if (err != LXLSX_NO_ERROR)
        return err;

    if (!find_sheet_data(xml_start, *xml_len, &sheet_start, &sheet_end,
                         &sheet_tag_end, &sheet_close_start)) {
        err = LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
        goto done;
    }

    if (!sheet_close_start) {
        const char *prefix_end = self_closing_prefix_end(sheet_start,
                                                         sheet_tag_end);
        if (buf_append(&out, xml_start, (size_t)(prefix_end - xml_start)) != 0 ||
            buf_append_s(&out, ">") != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }

        err = append_missing_rows_until(&out, prepared, prepared_count,
                                        &change_index, 0, 0);
        if (err != LXLSX_NO_ERROR)
            goto done;

        if (buf_append_matching_end_tag(&out, sheet_start, sheet_tag_end) != 0 ||
            buf_append(&out, sheet_end, (size_t)(xml_end - sheet_end)) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }
    } else {
        const char *pos = xml_start;
        const char *p = sheet_tag_end + 1;
        lxlsx_edit_xml_tag tag;

        while (xml_next_tag(p, sheet_close_start, &tag)) {
            const char *row_end;
            const char *row_close_start = NULL;
            char *row_ref_attr = NULL;
            lxlsx_row_t xml_row = 0;
            int valid_row = 0;

            p = tag.end + 1;
            if (tag.is_end || !tag_name_is(&tag, "row"))
                continue;

            if (tag.is_self_closing) {
                row_end = tag.end + 1;
            } else {
                lxlsx_edit_xml_tag close_tag;
                if (!find_matching_end_tag(tag.end + 1, sheet_close_start,
                                           "row", &close_tag))
                    goto malformed;
                row_close_start = close_tag.start;
                row_end = close_tag.end + 1;
            }

            row_ref_attr = extract_attr(tag.start, tag.end, "r");
            valid_row = parse_row_ref(row_ref_attr, &xml_row);
            free(row_ref_attr);

            if (!valid_row) {
                p = row_end;
                continue;
            }

            if (buf_append(&out, pos, (size_t)(tag.start - pos)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto done;
            }

            if (append_missing_rows_until(&out, prepared, prepared_count,
                                          &change_index, xml_row, 1)
                != LXLSX_NO_ERROR) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto done;
            }

            if (change_index < prepared_count &&
                prepared[change_index].change->row == xml_row) {
                char *row_xml;
                size_t next = change_index + 1;
                while (next < prepared_count &&
                       prepared[next].change->row == xml_row)
                    next++;

                row_xml = build_existing_row_xml(tag.start, row_end, tag.end,
                                                 row_close_start,
                                                 prepared + change_index,
                                                 next - change_index);
                if (!row_xml) {
                    err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                    goto done;
                }
                if (buf_append_s(&out, row_xml) != 0) {
                    free(row_xml);
                    err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                    goto done;
                }
                free(row_xml);
                change_index = next;
            } else if (buf_append(&out, tag.start,
                                  (size_t)(row_end - tag.start)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto done;
            }

            pos = row_end;
            p = row_end;
        }

        if (buf_append(&out, pos, (size_t)(sheet_close_start - pos)) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }

        err = append_missing_rows_until(&out, prepared, prepared_count,
                                        &change_index, 0, 0);
        if (err != LXLSX_NO_ERROR)
            goto done;

        if (buf_append(&out, sheet_close_start,
                       (size_t)(xml_end - sheet_close_start)) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }
    }

    free(*xml);
    *xml = (unsigned char *)out.data;
    *xml_len = out.len;
    out.data = NULL;
    err = LXLSX_NO_ERROR;
    goto done;

malformed:
    err = LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

done:
    free(out.data);
    free_prepared_changes(prepared, prepared_count);
    return err;
}

static const char *sheet_target(lxlsx_edit_session *session, const char *sheet_name)
{
    size_t i;
    lxlsx_reader_workbook *wb = session ? session->workbook : NULL;
    if (!wb || wb->sheet_count == 0)
        return NULL;
    if (!sheet_name)
        return wb->sheets[0].target;
    for (i = 0; i < wb->sheet_count; i++) {
        if (wb->sheets[i].name && strcmp(wb->sheets[i].name, sheet_name) == 0)
            return wb->sheets[i].target;
    }
    return NULL;
}

static lxlsx_error append_change(lxlsx_edit_session *session,
                                 const char *sheet_name,
                                 lxlsx_row_t row,
                                 lxlsx_col_t col,
                                 lxlsx_edit_change_type type,
                                 double number,
                                 const char *string,
                                 int boolean,
                                 const char *formula,
                                 const char *cached_result)
{
    const char *target;
    lxlsx_edit_change *change;

    if (!session)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (row >= LXLSX_ROW_MAX || col >= LXLSX_COL_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
    if (type == LXLSX_EDIT_CHANGE_FORMULA && (!formula || !*formula))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    if (session->change_count >= session->change_cap) {
        size_t cap = session->change_cap ? session->change_cap * 2 : 8;
        lxlsx_edit_change *next =
            (lxlsx_edit_change *)realloc(session->changes, cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        session->changes = next;
        session->change_cap = cap;
    }

    change = &session->changes[session->change_count];
    memset(change, 0, sizeof(*change));
    change->target = strdup(target);
    change->row = row;
    change->col = col;
    change->type = type;
    change->number = number;
    change->boolean = boolean ? 1 : 0;
    if (string)
        change->string = strdup(string);
    if (formula)
        change->formula = strdup(formula);
    if (cached_result)
        change->cached_result = strdup(cached_result);

    if (!change->target || (string && !change->string) ||
        (formula && !change->formula) ||
        (cached_result && !change->cached_result)) {
        free(change->target);
        free(change->string);
        free(change->formula);
        free(change->cached_result);
        memset(change, 0, sizeof(*change));
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    session->change_count++;
    return LXLSX_NO_ERROR;
}

lxlsx_edit_session *lxlsx_edit_open(const char *path)
{
    lxlsx_edit_session *session;
    size_t data_len = 0;
    const unsigned char *data;

    if (!path)
        return NULL;

    session = (lxlsx_edit_session *)calloc(1, sizeof(*session));
    if (!session)
        return NULL;

    if (lxlsx_source_package_open(path, &session->package) != LXLSX_NO_ERROR) {
        free(session);
        return NULL;
    }

    data = lxlsx_source_package_data(session->package, &data_len);
    if (lxlsx_reader_workbook_open_memory_borrowed(data, data_len, &session->workbook)
        != LXLSX_READER_NO_ERROR) {
        lxlsx_source_package_close(session->package);
        free(session);
        return NULL;
    }

    return session;
}

void lxlsx_edit_close(lxlsx_edit_session *session)
{
    size_t i;
    if (!session)
        return;
    for (i = 0; i < session->change_count; i++) {
        free(session->changes[i].target);
        free(session->changes[i].string);
        free(session->changes[i].formula);
        free(session->changes[i].cached_result);
    }
    free(session->changes);
    lxlsx_reader_workbook_close(session->workbook);
    lxlsx_source_package_close(session->package);
    free(session);
}

size_t lxlsx_edit_sheet_count(lxlsx_edit_session *session)
{
    if (!session || !session->workbook)
        return 0;
    return session->workbook->sheet_count;
}

const char *lxlsx_edit_sheet_name(lxlsx_edit_session *session, size_t index)
{
    if (!session || !session->workbook || index >= session->workbook->sheet_count)
        return NULL;
    return session->workbook->sheets[index].name;
}

lxlsx_error lxlsx_edit_set_number(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  double value)
{
    return append_change(session, sheet_name, row, col,
                         LXLSX_EDIT_CHANGE_NUMBER, value, NULL, 0, NULL, NULL);
}

lxlsx_error lxlsx_edit_set_string(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  const char *value)
{
    return append_change(session, sheet_name, row, col,
                         LXLSX_EDIT_CHANGE_STRING, 0.0,
                         value ? value : "", 0, NULL, NULL);
}

lxlsx_error lxlsx_edit_set_boolean(lxlsx_edit_session *session,
                                   const char *sheet_name,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col,
                                   int value)
{
    return append_change(session, sheet_name, row, col,
                         LXLSX_EDIT_CHANGE_BOOLEAN, 0.0, NULL,
                         value, NULL, NULL);
}

lxlsx_error lxlsx_edit_set_formula(lxlsx_edit_session *session,
                                   const char *sheet_name,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col,
                                   const char *formula,
                                   const char *cached_result)
{
    return append_change(session, sheet_name, row, col,
                         LXLSX_EDIT_CHANGE_FORMULA, 0.0, NULL, 0, formula,
                         cached_result);
}

static void free_dirty_sheets(lxlsx_dirty_sheet *sheets, size_t count)
{
    size_t i;
    for (i = 0; i < count; i++) {
        free(sheets[i].target);
        free(sheets[i].xml);
        free(sheets[i].changes);
    }
    free(sheets);
}

static lxlsx_dirty_sheet *find_dirty_sheet(lxlsx_dirty_sheet *sheets,
                                           size_t count,
                                           const char *target)
{
    size_t i;
    for (i = 0; i < count; i++) {
        if (strcmp(sheets[i].target, target) == 0)
            return &sheets[i];
    }
    return NULL;
}

static lxlsx_error get_dirty_sheet(lxlsx_edit_session *session,
                                   lxlsx_dirty_sheet **sheets,
                                   size_t *count,
                                   size_t *cap,
                                   const char *target,
                                   lxlsx_dirty_sheet **out)
{
    lxlsx_dirty_sheet *sheet = find_dirty_sheet(*sheets, *count, target);
    int entry_index;
    lxlsx_error err;

    if (sheet) {
        *out = sheet;
        return LXLSX_NO_ERROR;
    }

    if (*count >= *cap) {
        size_t next_cap = *cap ? *cap * 2 : 4;
        lxlsx_dirty_sheet *next =
            (lxlsx_dirty_sheet *)realloc(*sheets, next_cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        *sheets = next;
        *cap = next_cap;
    }

    sheet = &(*sheets)[*count];
    memset(sheet, 0, sizeof(*sheet));
    entry_index = lxlsx_source_package_find_first(session->package, target);
    if (entry_index < 0)
        return LXLSX_ERROR_ZIP_FILE_ADD;

    sheet->target = strdup(target);
    if (!sheet->target)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    sheet->entry_index = (size_t)entry_index;

    err = lxlsx_source_package_read_entry(session->package, sheet->entry_index,
                                          &sheet->xml, &sheet->xml_len);
    if (err != LXLSX_NO_ERROR) {
        free(sheet->target);
        memset(sheet, 0, sizeof(*sheet));
        return err;
    }

    (*count)++;
    *out = sheet;
    return LXLSX_NO_ERROR;
}

static lxlsx_error dirty_sheet_add_change(lxlsx_dirty_sheet *sheet,
                                          const lxlsx_edit_change *change)
{
    const lxlsx_edit_change **next;
    size_t next_cap;

    if (sheet->change_count >= sheet->change_cap) {
        next_cap = sheet->change_cap ? sheet->change_cap * 2 : 8;
        next = (const lxlsx_edit_change **)realloc(
            sheet->changes, next_cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        sheet->changes = next;
        sheet->change_cap = next_cap;
    }

    sheet->changes[sheet->change_count++] = change;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_edit_save_as(lxlsx_edit_session *session, const char *path)
{
    lxlsx_dirty_sheet *dirty_sheets = NULL;
    size_t dirty_count = 0;
    size_t dirty_cap = 0;
    lxlsx_source_package_replacement *replacements = NULL;
    lxlsx_error err = LXLSX_NO_ERROR;
    size_t i;

    if (!session || !path)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (session->change_count == 0)
        return lxlsx_source_package_save_copy(session->package, path);

    for (i = 0; i < session->change_count; i++) {
        lxlsx_dirty_sheet *sheet = NULL;
        err = get_dirty_sheet(session, &dirty_sheets, &dirty_count, &dirty_cap,
                              session->changes[i].target, &sheet);
        if (err != LXLSX_NO_ERROR)
            goto done;
        err = dirty_sheet_add_change(sheet, &session->changes[i]);
        if (err != LXLSX_NO_ERROR)
            goto done;
    }

    for (i = 0; i < dirty_count; i++) {
        err = patch_xml_with_changes(&dirty_sheets[i].xml,
                                     &dirty_sheets[i].xml_len,
                                     dirty_sheets[i].changes,
                                     dirty_sheets[i].change_count);
        if (err != LXLSX_NO_ERROR)
            goto done;
    }

    replacements = (lxlsx_source_package_replacement *)calloc(
        dirty_count, sizeof(*replacements));
    if (!replacements) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto done;
    }

    for (i = 0; i < dirty_count; i++) {
        replacements[i].entry_index = dirty_sheets[i].entry_index;
        replacements[i].data = dirty_sheets[i].xml;
        replacements[i].size = dirty_sheets[i].xml_len;
    }

    err = lxlsx_source_package_save_with_replacements(session->package, path,
                                                      replacements, dirty_count);

done:
    free(replacements);
    free_dirty_sheets(dirty_sheets, dirty_count);
    return err;
}
