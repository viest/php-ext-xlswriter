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
#include "libxlsx/utility.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/workbook.h"
#include "libxlsx/xmlwriter.h"

#include "xlsx_private.h"
#include "xlsx_util.h"

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
    char *number_format;  /* requested number-format code, or NULL */
    lxlsx_format *style;   /* shallow snapshot of an applied format, or NULL */
    int   style_index;    /* xf index assigned on save (-1 = keep existing) */
} lxlsx_edit_change;

typedef struct {
    char *target;
    lxlsx_row_t first_row;
    lxlsx_col_t first_col;
    lxlsx_row_t last_row;
    lxlsx_col_t last_col;
} lxlsx_edit_merge;

typedef struct {
    char *target;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;
    double width;
} lxlsx_edit_col;

typedef struct {
    char *target;
    lxlsx_row_t row;
    double height;
} lxlsx_edit_row_dim;

typedef struct {
    char  *name;      /* sheet display name */
    char  *xml;       /* assembled sheetN.xml content (owned) */
    size_t xml_len;
    char  *filename;  /* assigned part name, e.g. xl/worksheets/sheet3.xml */
} lxlsx_edit_new_sheet;

struct lxlsx_edit_session {
    lxlsx_source_package *package;
    lxlsx_reader_workbook *workbook;
    lxlsx_edit_change *changes;
    size_t change_count;
    size_t change_cap;
    lxlsx_edit_merge *merges;
    size_t merge_count;
    size_t merge_cap;
    lxlsx_edit_col *cols;
    size_t col_count;
    size_t col_cap;
    lxlsx_edit_row_dim *row_dims;
    size_t row_dim_count;
    size_t row_dim_cap;
    lxlsx_edit_new_sheet *new_sheets;
    size_t new_sheet_count;
    size_t new_sheet_cap;
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
    const lxlsx_edit_merge **merges;
    size_t merge_count;
    size_t merge_cap;
    const lxlsx_edit_col **cols;
    size_t col_count;
    size_t col_cap;
    const lxlsx_edit_row_dim **row_dims;
    size_t row_dim_count;
    size_t row_dim_cap;
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
    return lxlsx_reader_strndup(start, len);
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
    char ref[LXLSX_MAX_CELL_NAME_LENGTH];
    lxlsx_rowcol_to_cell(ref, row, col);
    return lxlsx_reader_strdup(ref);
}

static char *make_row_ref(lxlsx_row_t row)
{
    char ref[32];
    snprintf(ref, sizeof(ref), "%u", (unsigned int)row + 1);
    return lxlsx_reader_strdup(ref);
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

static lxlsx_error append_cell_xml(lxlsx_edit_buf *buf,
                                   const lxlsx_edit_change *change,
                                   const char *ref,
                                   const char *style)
{
    char number[64];
    char style_buf[16];

    /* A restyle change overrides the original cell's s= index. */
    if (change->style_index >= 0) {
        snprintf(style_buf, sizeof(style_buf), "%d", change->style_index);
        style = style_buf;
    }

    if (buf_appendf(buf, "<c r=\"%s\"", ref) != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    if (style && *style) {
        if (buf_appendf(buf, " s=\"%s\"", style) != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (change->type == LXLSX_EDIT_CHANGE_STRING) {
        if (buf_append_s(buf, " t=\"inlineStr\"") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    } else if (change->type == LXLSX_EDIT_CHANGE_BOOLEAN) {
        if (buf_append_s(buf, " t=\"b\"") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (change->type == LXLSX_EDIT_CHANGE_FORMULA &&
        change->cached_result && !looks_numeric(change->cached_result)) {
        if (buf_append_s(buf, " t=\"str\"") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (buf_append_s(buf, ">") != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    if (change->type == LXLSX_EDIT_CHANGE_NUMBER) {
        snprintf(number, sizeof(number), "%.17g", change->number);
        if (buf_append_s(buf, "<v>") != 0 ||
            buf_append_s(buf, number) != 0 ||
            buf_append_s(buf, "</v>") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    } else if (change->type == LXLSX_EDIT_CHANGE_STRING) {
        const char *string = change->string ? change->string : "";
        if (buf_append_s(buf, "<is><t") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        if (needs_xml_space_preserve(string) &&
            buf_append_s(buf, " xml:space=\"preserve\"") != 0) {
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
        if (buf_append_s(buf, ">") != 0 ||
            append_xml_escaped(buf, string) != 0 ||
            buf_append_s(buf, "</t></is>") != 0) {
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
    } else if (change->type == LXLSX_EDIT_CHANGE_BOOLEAN) {
        if (buf_appendf(buf, "<v>%d</v>", change->boolean ? 1 : 0) != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    } else {
        const char *formula = change->formula;
        const char *cached = change->cached_result ? change->cached_result : "0";
        if (formula && formula[0] == '=')
            formula++;
        if (buf_append_s(buf, "<f>") != 0 ||
            append_xml_escaped(buf, formula ? formula : "") != 0 ||
            buf_append_s(buf, "</f><v>") != 0 ||
            append_xml_escaped(buf, cached) != 0 ||
            buf_append_s(buf, "</v>") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (buf_append_s(buf, LXLSX_EDIT_CELL_CLOSE) != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    return LXLSX_NO_ERROR;
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
        lxlsx_error err = append_cell_xml(buf, changes[i].change,
                                          changes[i].ref, NULL);
        if (err != LXLSX_NO_ERROR)
            return err;
    }

    return LXLSX_NO_ERROR;
}

static lxlsx_error append_change_cell(lxlsx_edit_buf *buf,
                                      const lxlsx_edit_prepared_change *change,
                                      const char *style)
{
    return append_cell_xml(buf, change->change, change->ref, style);
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

static lxlsx_error append_new_row_xml(lxlsx_edit_buf *buf,
                                      const char *row_ref,
                                      const lxlsx_edit_prepared_change *changes,
                                      size_t count)
{
    if (buf_appendf(buf, "<row r=\"%s\">", row_ref) != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    if (append_cells(buf, changes, count) != LXLSX_NO_ERROR)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    if (buf_append_s(buf, LXLSX_EDIT_ROW_CLOSE) != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    return LXLSX_NO_ERROR;
}

static lxlsx_error append_existing_row_xml(
    lxlsx_edit_buf *row,
    const char *row_start,
    const char *row_end,
    const char *row_tag_end,
    const char *row_close_start,
    const lxlsx_edit_prepared_change *changes,
    size_t count)
{
    size_t change_index = 0;

    if (!row_close_start) {
        const char *prefix_end = self_closing_prefix_end(row_start, row_tag_end);
        if (buf_append(row, row_start, (size_t)(prefix_end - row_start)) != 0 ||
            buf_append_s(row, ">") != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        if (append_remaining_cells(row, changes, count, &change_index)
            != LXLSX_NO_ERROR)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        if (buf_append_matching_end_tag(row, row_start, row_tag_end) != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
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
                    return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
                cell_end = close_tag.end + 1;
            }

            ref = extract_attr(tag.start, tag.end, "r");
            valid_ref = parse_cell_ref(ref, &cell_row, &cell_col);
            free(ref);

            if (!valid_ref || cell_row != changes[0].change->row) {
                p = cell_end;
                continue;
            }

            if (buf_append(row, pos, (size_t)(tag.start - pos)) != 0)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

            if (append_cells_before_col(row, changes, count, &change_index,
                                        cell_col) != LXLSX_NO_ERROR)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

            if (change_index < count &&
                changes[change_index].change->col == cell_col) {
                char *style = extract_attr(tag.start, tag.end, "s");
                lxlsx_error err = append_change_cell(
                    row, &changes[change_index], style);
                free(style);
                if (err != LXLSX_NO_ERROR)
                    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                change_index++;
            } else {
                if (buf_append(row, tag.start,
                               (size_t)(cell_end - tag.start)) != 0)
                    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            }

            pos = cell_end;
            p = cell_end;
        }

        if (buf_append(row, pos, (size_t)(row_close_start - pos)) != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

        if (append_remaining_cells(row, changes, count, &change_index)
            != LXLSX_NO_ERROR)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

        if (buf_append(row, row_close_start,
                       (size_t)(row_end - row_close_start)) != 0)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    return LXLSX_NO_ERROR;
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
        size_t next = *index + 1;
        lxlsx_error err;

        while (next < prepared_count && prepared[next].change->row == row)
            next++;

        err = append_new_row_xml(out, row_ref, prepared + *index,
                                 next - *index);
        if (err != LXLSX_NO_ERROR)
            return err;
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
                size_t next = change_index + 1;
                lxlsx_error row_err;
                while (next < prepared_count &&
                       prepared[next].change->row == xml_row)
                    next++;

                row_err = append_existing_row_xml(&out, tag.start, row_end,
                                                  tag.end, row_close_start,
                                                  prepared + change_index,
                                                  next - change_index);
                if (row_err != LXLSX_NO_ERROR) {
                    err = row_err;
                    goto done;
                }
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
    change->style_index = -1;
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
        free(session->changes[i].number_format);
        free(session->changes[i].style);
    }
    free(session->changes);
    for (i = 0; i < session->merge_count; i++)
        free(session->merges[i].target);
    free(session->merges);
    for (i = 0; i < session->col_count; i++)
        free(session->cols[i].target);
    free(session->cols);
    for (i = 0; i < session->row_dim_count; i++)
        free(session->row_dims[i].target);
    free(session->row_dims);
    for (i = 0; i < session->new_sheet_count; i++) {
        free(session->new_sheets[i].name);
        free(session->new_sheets[i].xml);
        free(session->new_sheets[i].filename);
    }
    free(session->new_sheets);
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

lxlsx_error lxlsx_edit_set_number_format(lxlsx_edit_session *session,
                                         const char *sheet_name,
                                         lxlsx_row_t row,
                                         lxlsx_col_t col,
                                         const char *format_code)
{
    const char *target;
    size_t i;

    if (!session || !format_code || !*format_code)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    /* Attach the format to the most recent value change for this cell (the
     * caller writes the value immediately before requesting the format). */
    for (i = session->change_count; i > 0; i--) {
        lxlsx_edit_change *c = &session->changes[i - 1];
        if (c->row == row && c->col == col && strcmp(c->target, target) == 0) {
            char *code = strdup(format_code);
            if (!code)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            free(c->number_format);
            c->number_format = code;
            return LXLSX_NO_ERROR;
        }
    }
    return LXLSX_ERROR_PARAMETER_VALIDATION;
}

lxlsx_error lxlsx_edit_set_format(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  const lxlsx_format *format)
{
    const char *target;
    size_t i;

    if (!session || !format)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    for (i = session->change_count; i > 0; i--) {
        lxlsx_edit_change *c = &session->changes[i - 1];
        if (c->row == row && c->col == col && strcmp(c->target, target) == 0) {
            /* Shallow snapshot: only scalar style fields are read later, so the
             * format object's lifetime no longer matters. */
            lxlsx_format *copy = (lxlsx_format *)malloc(sizeof(lxlsx_format));
            if (!copy)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            memcpy(copy, format, sizeof(lxlsx_format));

            /* Replicate the writer's solid-fill fg/bg normalization so a
             * Format::background() color lands in fgColor like create mode
             * (see lxlsx_workbook fill prep). */
            if (copy->pattern == LXLSX_PATTERN_SOLID &&
                copy->bg_color != LXLSX_COLOR_UNSET &&
                copy->fg_color != LXLSX_COLOR_UNSET) {
                lxlsx_color_t t = copy->fg_color;
                copy->fg_color = copy->bg_color;
                copy->bg_color = t;
            }
            if (copy->pattern <= LXLSX_PATTERN_SOLID &&
                copy->bg_color != LXLSX_COLOR_UNSET &&
                copy->fg_color == LXLSX_COLOR_UNSET) {
                copy->fg_color = copy->bg_color;
                copy->bg_color = LXLSX_COLOR_UNSET;
                copy->pattern = LXLSX_PATTERN_SOLID;
            }
            if (copy->pattern <= LXLSX_PATTERN_SOLID &&
                copy->bg_color == LXLSX_COLOR_UNSET &&
                copy->fg_color != LXLSX_COLOR_UNSET) {
                copy->pattern = LXLSX_PATTERN_SOLID;
            }

            free(c->style);
            c->style = copy;
            return LXLSX_NO_ERROR;
        }
    }
    return LXLSX_ERROR_PARAMETER_VALIDATION;
}

lxlsx_error lxlsx_edit_set_merge(lxlsx_edit_session *session,
                                 const char *sheet_name,
                                 lxlsx_row_t first_row,
                                 lxlsx_col_t first_col,
                                 lxlsx_row_t last_row,
                                 lxlsx_col_t last_col)
{
    const char *target;
    lxlsx_edit_merge *merge;

    if (!session)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    /* Normalise so first <= last. */
    if (first_row > last_row) { lxlsx_row_t t = first_row; first_row = last_row; last_row = t; }
    if (first_col > last_col) { lxlsx_col_t t = first_col; first_col = last_col; last_col = t; }
    if (last_row >= LXLSX_ROW_MAX || last_col >= LXLSX_COL_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
    if (first_row == last_row && first_col == last_col)
        return LXLSX_ERROR_PARAMETER_VALIDATION;  /* a single cell can't merge */

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    if (session->merge_count >= session->merge_cap) {
        size_t cap = session->merge_cap ? session->merge_cap * 2 : 8;
        lxlsx_edit_merge *next =
            (lxlsx_edit_merge *)realloc(session->merges, cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        session->merges = next;
        session->merge_cap = cap;
    }

    merge = &session->merges[session->merge_count];
    memset(merge, 0, sizeof(*merge));
    merge->target = strdup(target);
    if (!merge->target)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    merge->first_row = first_row;
    merge->first_col = first_col;
    merge->last_row = last_row;
    merge->last_col = last_col;

    session->merge_count++;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_edit_add_sheet(lxlsx_edit_session *session,
                                 const char *name,
                                 const char *xml,
                                 size_t xml_len)
{
    lxlsx_edit_new_sheet *s;

    if (!session || !name || !*name || !xml)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (session->new_sheet_count >= session->new_sheet_cap) {
        size_t cap = session->new_sheet_cap ? session->new_sheet_cap * 2 : 4;
        lxlsx_edit_new_sheet *next =
            (lxlsx_edit_new_sheet *)realloc(session->new_sheets, cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        session->new_sheets = next;
        session->new_sheet_cap = cap;
    }
    s = &session->new_sheets[session->new_sheet_count];
    memset(s, 0, sizeof(*s));
    s->name = strdup(name);
    s->xml = (char *)malloc(xml_len + 1);
    if (!s->name || !s->xml) { free(s->name); free(s->xml); return LXLSX_ERROR_MEMORY_MALLOC_FAILED; }
    memcpy(s->xml, xml, xml_len);
    s->xml[xml_len] = '\0';
    s->xml_len = xml_len;
    session->new_sheet_count++;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_edit_set_column(lxlsx_edit_session *session,
                                 const char *sheet_name,
                                 lxlsx_col_t first_col,
                                 lxlsx_col_t last_col,
                                 double width)
{
    const char *target;
    lxlsx_edit_col *c;

    if (!session)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (first_col > last_col) { lxlsx_col_t t = first_col; first_col = last_col; last_col = t; }
    if (last_col >= LXLSX_COL_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    if (session->col_count >= session->col_cap) {
        size_t cap = session->col_cap ? session->col_cap * 2 : 8;
        lxlsx_edit_col *next = (lxlsx_edit_col *)realloc(session->cols, cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        session->cols = next;
        session->col_cap = cap;
    }
    c = &session->cols[session->col_count];
    memset(c, 0, sizeof(*c));
    c->target = strdup(target);
    if (!c->target)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    c->first_col = first_col;
    c->last_col = last_col;
    c->width = width;
    session->col_count++;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_edit_set_row_height(lxlsx_edit_session *session,
                                     const char *sheet_name,
                                     lxlsx_row_t row,
                                     double height)
{
    const char *target;
    lxlsx_edit_row_dim *r;

    if (!session)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (row >= LXLSX_ROW_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    target = sheet_target(session, sheet_name);
    if (!target)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    if (session->row_dim_count >= session->row_dim_cap) {
        size_t cap = session->row_dim_cap ? session->row_dim_cap * 2 : 8;
        lxlsx_edit_row_dim *next = (lxlsx_edit_row_dim *)realloc(session->row_dims, cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        session->row_dims = next;
        session->row_dim_cap = cap;
    }
    r = &session->row_dims[session->row_dim_count];
    memset(r, 0, sizeof(*r));
    r->target = strdup(target);
    if (!r->target)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    r->row = row;
    r->height = height;
    session->row_dim_count++;
    return LXLSX_NO_ERROR;
}

static void free_dirty_sheets(lxlsx_dirty_sheet *sheets, size_t count)
{
    size_t i;
    for (i = 0; i < count; i++) {
        free(sheets[i].target);
        free(sheets[i].xml);
        free(sheets[i].changes);
        free(sheets[i].merges);
        free(sheets[i].cols);
        free(sheets[i].row_dims);
    }
    free(sheets);
}

static lxlsx_error dirty_sheet_add_merge(lxlsx_dirty_sheet *sheet,
                                         const lxlsx_edit_merge *merge)
{
    const lxlsx_edit_merge **next;
    size_t next_cap;

    if (sheet->merge_count >= sheet->merge_cap) {
        next_cap = sheet->merge_cap ? sheet->merge_cap * 2 : 8;
        next = (const lxlsx_edit_merge **)realloc(
            sheet->merges, next_cap * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        sheet->merges = next;
        sheet->merge_cap = next_cap;
    }

    sheet->merges[sheet->merge_count++] = merge;
    return LXLSX_NO_ERROR;
}

static lxlsx_error dirty_sheet_add_col(lxlsx_dirty_sheet *sheet,
                                       const lxlsx_edit_col *col)
{
    if (sheet->col_count >= sheet->col_cap) {
        size_t nc = sheet->col_cap ? sheet->col_cap * 2 : 8;
        const lxlsx_edit_col **next =
            (const lxlsx_edit_col **)realloc(sheet->cols, nc * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        sheet->cols = next;
        sheet->col_cap = nc;
    }
    sheet->cols[sheet->col_count++] = col;
    return LXLSX_NO_ERROR;
}

static lxlsx_error dirty_sheet_add_row_dim(lxlsx_dirty_sheet *sheet,
                                           const lxlsx_edit_row_dim *rd)
{
    if (sheet->row_dim_count >= sheet->row_dim_cap) {
        size_t nc = sheet->row_dim_cap ? sheet->row_dim_cap * 2 : 8;
        const lxlsx_edit_row_dim **next =
            (const lxlsx_edit_row_dim **)realloc(sheet->row_dims, nc * sizeof(*next));
        if (!next)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        sheet->row_dims = next;
        sheet->row_dim_cap = nc;
    }
    sheet->row_dims[sheet->row_dim_count++] = rd;
    return LXLSX_NO_ERROR;
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

/* Length-aware substring search. */
static const char *mem_find(const char *hay, size_t haylen, const char *needle)
{
    size_t nl = strlen(needle);
    size_t i;
    if (nl == 0 || haylen < nl)
        return NULL;
    for (i = 0; i + nl <= haylen; i++) {
        if (memcmp(hay + i, needle, nl) == 0)
            return hay + i;
    }
    return NULL;
}

/*
 * ECMA-376 child element order for <worksheet>. Used to insert a new child at
 * the schema-correct position when it is absent — so we never place, say,
 * <mergeCells> before <sheetProtection> just because it followed </sheetData>.
 */
static const char *const WS_CHILD_ORDER[] = {
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
    "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
    "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
    "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
    "rowBreaks", "colBreaks", "customProperties", "cellWatches",
    "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF",
    "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
    "extLst"
};

static int ws_child_rank(const char *tag)
{
    size_t i;
    for (i = 0; i < sizeof(WS_CHILD_ORDER) / sizeof(WS_CHILD_ORDER[0]); i++)
        if (strcmp(WS_CHILD_ORDER[i], tag) == 0)
            return (int)i;
    return -1;
}

/* Find an element start "<name" whose name is delimited (not a prefix match). */
static const char *find_element(const char *src, size_t len, const char *name)
{
    size_t nl = strlen(name);
    const char *p = src, *end = src + len;
    while ((p = mem_find(p, (size_t)(end - p), "<")) != NULL) {
        if (p + 1 + nl <= end && memcmp(p + 1, name, nl) == 0) {
            char c = (p + 1 + nl < end) ? p[1 + nl] : '\0';
            if (c == ' ' || c == '>' || c == '/' || c == '\t' || c == '\n' || c == '\r')
                return p;
        }
        p += 1;
    }
    return NULL;
}

/*
 * Byte position within [src, src+len) at which a new <tag> worksheet child
 * should be inserted: before the first existing child that must sort after it,
 * else before </worksheet>.
 */
static const char *ws_insert_point(const char *src, size_t len, const char *tag)
{
    int rank = ws_child_rank(tag);
    const char *best = NULL;
    size_t i;
    if (rank >= 0) {
        for (i = (size_t)rank + 1; i < sizeof(WS_CHILD_ORDER) / sizeof(WS_CHILD_ORDER[0]); i++) {
            const char *pos = find_element(src, len, WS_CHILD_ORDER[i]);
            if (pos && (!best || pos < best))
                best = pos;
        }
    }
    return best ? best : mem_find(src, len, "</worksheet>");
}

static void col_to_letters(lxlsx_col_t col, char *out)
{
    char tmp[8];
    int n = 0, j = 0;
    unsigned v = (unsigned)col + 1;
    while (v) { tmp[n++] = (char)('A' + (v - 1) % 26); v = (v - 1) / 26; }
    while (n) out[j++] = tmp[--n];
    out[j] = '\0';
}

static void merge_ref(char *buf, const lxlsx_edit_merge *m)
{
    char c1[8], c2[8];
    col_to_letters(m->first_col, c1);
    col_to_letters(m->last_col, c2);
    sprintf(buf, "%s%u:%s%u", c1, (unsigned)(m->first_row + 1),
            c2, (unsigned)(m->last_row + 1));
}

/*
 * Inject merged ranges into a worksheet part's <mergeCells>. Creates the
 * element at its schema-correct position (per the worksheet child order) if
 * absent, or splices into an existing block and bumps its count. Everything
 * else is preserved verbatim.
 */
static lxlsx_error patch_xml_with_merges(unsigned char **xml, size_t *xml_len,
                                         const lxlsx_edit_merge **merges,
                                         size_t count)
{
    const char *src = (const char *)*xml;
    size_t srclen = *xml_len;
    lxlsx_edit_buf entries = {0};
    lxlsx_edit_buf out = {0};
    const char *mc, *gt;
    char ref[32], open[48];
    size_t i, existing = 0, total;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (count == 0)
        return LXLSX_NO_ERROR;

    for (i = 0; i < count; i++) {
        merge_ref(ref, merges[i]);
        if (buf_append_s(&entries, "<mergeCell ref=\"") != 0 ||
            buf_append_s(&entries, ref) != 0 ||
            buf_append_s(&entries, "\"/>") != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }
    }

    mc = mem_find(src, srclen, "<mergeCells");
    if (mc) {
        const char *cnt;
        int self_closing;
        gt = (const char *)memchr(mc, '>', srclen - (size_t)(mc - src));
        if (!gt) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        self_closing = (gt > mc && *(gt - 1) == '/');

        cnt = mem_find(mc, (size_t)(gt - mc), "count=\"");
        if (cnt) {
            cnt += strlen("count=\"");
            while (cnt < gt && *cnt >= '0' && *cnt <= '9')
                existing = existing * 10 + (size_t)(*cnt++ - '0');
        }
        total = existing + count;
        snprintf(open, sizeof(open), "<mergeCells count=\"%zu\">", total);

        if (buf_append(&out, src, (size_t)(mc - src)) != 0 ||
            buf_append_s(&out, open) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
        }
        if (self_closing) {
            /* <mergeCells .../> -> <mergeCells count="K">entries</mergeCells> */
            if (buf_append(&out, entries.data, entries.len) != 0 ||
                buf_append_s(&out, "</mergeCells>") != 0 ||
                buf_append(&out, gt + 1, srclen - (size_t)(gt + 1 - src)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
            }
        } else {
            const char *close = mem_find(gt, srclen - (size_t)(gt - src),
                                         "</mergeCells>");
            if (!close) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
            /* keep existing entries, append ours before </mergeCells> */
            if (buf_append(&out, gt + 1, (size_t)(close - (gt + 1))) != 0 ||
                buf_append(&out, entries.data, entries.len) != 0 ||
                buf_append(&out, close, srclen - (size_t)(close - src)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
            }
        }
    } else {
        const char *at = ws_insert_point(src, srclen, "mergeCells");
        if (!at) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        snprintf(open, sizeof(open), "<mergeCells count=\"%zu\">", count);
        if (buf_append(&out, src, (size_t)(at - src)) != 0 ||
            buf_append_s(&out, open) != 0 ||
            buf_append(&out, entries.data, entries.len) != 0 ||
            buf_append_s(&out, "</mergeCells>") != 0 ||
            buf_append(&out, at, srclen - (size_t)(at - src)) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
        }
    }

    free(*xml);
    *xml = (unsigned char *)out.data;
    *xml_len = out.len;
    out.data = NULL;

done:
    free(entries.data);
    free(out.data);
    return err;
}

/*
 * Inject column widths into a worksheet's <cols> (no count attr). Creates the
 * element at its schema-correct position (per the worksheet child order) if absent.
 */
static lxlsx_error patch_xml_with_cols(unsigned char **xml, size_t *xml_len,
                                       const lxlsx_edit_col **cols, size_t count)
{
    const char *src = (const char *)*xml;
    size_t srclen = *xml_len;
    lxlsx_edit_buf entries = {0}, out = {0};
    const char *cc, *gt;
    char ent[96];
    size_t i;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (count == 0)
        return LXLSX_NO_ERROR;

    for (i = 0; i < count; i++) {
        snprintf(ent, sizeof(ent),
                 "<col min=\"%u\" max=\"%u\" width=\"%g\" customWidth=\"1\"/>",
                 (unsigned)(cols[i]->first_col + 1), (unsigned)(cols[i]->last_col + 1),
                 cols[i]->width);
        if (buf_append_s(&entries, ent) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
    }

    cc = mem_find(src, srclen, "<cols");
    if (cc) {
        int self_closing;
        gt = (const char *)memchr(cc, '>', srclen - (size_t)(cc - src));
        if (!gt) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        self_closing = (gt > cc && gt[-1] == '/');
        if (self_closing) {
            if (buf_append(&out, src, (size_t)(cc - src)) != 0 ||
                buf_append_s(&out, "<cols>") != 0 ||
                buf_append(&out, entries.data, entries.len) != 0 ||
                buf_append_s(&out, "</cols>") != 0 ||
                buf_append(&out, gt + 1, srclen - (size_t)(gt + 1 - src)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
            }
        } else {
            const char *close = mem_find(gt, srclen - (size_t)(gt - src), "</cols>");
            if (!close) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
            if (buf_append(&out, src, (size_t)(close - src)) != 0 ||
                buf_append(&out, entries.data, entries.len) != 0 ||
                buf_append(&out, close, srclen - (size_t)(close - src)) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
            }
        }
    } else {
        const char *at = ws_insert_point(src, srclen, "cols");
        if (!at) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        if (buf_append(&out, src, (size_t)(at - src)) != 0 ||
            buf_append_s(&out, "<cols>") != 0 ||
            buf_append(&out, entries.data, entries.len) != 0 ||
            buf_append_s(&out, "</cols>") != 0 ||
            buf_append(&out, at, srclen - (size_t)(at - src)) != 0) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
        }
    }

    free(*xml);
    *xml = (unsigned char *)out.data;
    *xml_len = out.len;
    out.data = NULL;

done:
    free(entries.data);
    free(out.data);
    return err;
}

/*
 * Apply row heights by rewriting each existing <row r="N"> opening tag's ht /
 * customHeight attributes. Rows that don't exist in the sheet are skipped.
 */
static lxlsx_error patch_xml_with_row_dims(unsigned char **xml, size_t *xml_len,
                                           const lxlsx_edit_row_dim **rows,
                                           size_t count)
{
    size_t i;
    for (i = 0; i < count; i++) {
        const char *src = (const char *)*xml;
        size_t srclen = *xml_len;
        char needle[24], attr[48];
        const char *ts, *gt, *p, *limit;
        lxlsx_edit_buf out = {0};
        int self_closing;

        snprintf(needle, sizeof(needle), "<row r=\"%u\"", (unsigned)(rows[i]->row + 1));
        ts = mem_find(src, srclen, needle);
        if (!ts)
            continue;  /* only existing rows */
        gt = (const char *)memchr(ts, '>', srclen - (size_t)(ts - src));
        if (!gt)
            return LXLSX_ERROR_PARAMETER_VALIDATION;
        self_closing = (gt > ts && gt[-1] == '/');
        limit = self_closing ? gt - 1 : gt;

        if (buf_append(&out, src, (size_t)(ts - src)) != 0 ||
            buf_append_s(&out, "<row") != 0) { free(out.data); return LXLSX_ERROR_MEMORY_MALLOC_FAILED; }

        /* Copy existing attributes, dropping any ht / customHeight. */
        p = ts + 4;  /* past "<row" */
        while (p < limit) {
            if ((size_t)(limit - p) >= 4 && memcmp(p, " ht=", 4) == 0) {
                p += 4;
                if (p < limit && *p == '"') { p++; while (p < limit && *p != '"') p++; if (p < limit) p++; }
                continue;
            }
            if ((size_t)(limit - p) >= 14 && memcmp(p, " customHeight=", 14) == 0) {
                p += 14;
                if (p < limit && *p == '"') { p++; while (p < limit && *p != '"') p++; if (p < limit) p++; }
                continue;
            }
            if (buf_append(&out, p, 1) != 0) { free(out.data); return LXLSX_ERROR_MEMORY_MALLOC_FAILED; }
            p++;
        }

        snprintf(attr, sizeof(attr), " ht=\"%g\" customHeight=\"1\"", rows[i]->height);
        if (buf_append_s(&out, attr) != 0 ||
            buf_append_s(&out, self_closing ? "/>" : ">") != 0 ||
            buf_append(&out, gt + 1, srclen - (size_t)(gt + 1 - src)) != 0) {
            free(out.data); return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }

        free(*xml);
        *xml = (unsigned char *)out.data;
        *xml_len = out.len;
    }
    return LXLSX_NO_ERROR;
}

static int append_attr_escaped(lxlsx_edit_buf *buf, const char *s)
{
    for (; *s; s++) {
        const char *rep = NULL;
        switch (*s) {
        case '&': rep = "&amp;"; break;
        case '<': rep = "&lt;"; break;
        case '>': rep = "&gt;"; break;
        case '"': rep = "&quot;"; break;
        default: break;
        }
        if (rep) { if (buf_append_s(buf, rep) != 0) return -1; }
        else if (buf_append(buf, s, 1) != 0) return -1;
    }
    return 0;
}

static size_t parse_count_attr(const char *open, const char *gt)
{
    const char *c = mem_find(open, (size_t)(gt - open), "count=\"");
    size_t v = 0;
    if (c) {
        c += strlen("count=\"");
        while (c < gt && *c >= '0' && *c <= '9')
            v = v * 10 + (size_t)(*c++ - '0');
    }
    return v;
}

/*
 * Insert `entries` into the <tag> collection of `src`, bumping its count by
 * `added`. If the collection is absent it is created right after the opening
 * tag of `create_after` (a parent element). Returns a fresh buffer.
 */
static lxlsx_error splice_collection(const char *src, size_t srclen,
                                     const char *tag, const char *entries,
                                     size_t added, const char *create_after,
                                     char **out_data, size_t *out_len)
{
    lxlsx_edit_buf out = {0};
    char openpat[32], openbuf[48], closebuf[32];
    const char *t, *gt;
    lxlsx_error err = LXLSX_ERROR_PARAMETER_VALIDATION;

    snprintf(openpat, sizeof(openpat), "<%s", tag);
    snprintf(closebuf, sizeof(closebuf), "</%s>", tag);
    t = mem_find(src, srclen, openpat);

    if (t) {
        size_t total;
        int self_closing;
        gt = (const char *)memchr(t, '>', srclen - (size_t)(t - src));
        if (!gt) goto done;
        self_closing = (gt > t && gt[-1] == '/');
        total = parse_count_attr(t, gt) + added;
        snprintf(openbuf, sizeof(openbuf), "<%s count=\"%zu\">", tag, total);
        if (buf_append(&out, src, (size_t)(t - src)) != 0 ||
            buf_append_s(&out, openbuf) != 0) goto oom;
        if (self_closing) {
            if (buf_append_s(&out, entries) != 0 ||
                buf_append_s(&out, closebuf) != 0 ||
                buf_append(&out, gt + 1, srclen - (size_t)(gt + 1 - src)) != 0) goto oom;
        } else {
            const char *close = mem_find(gt, srclen - (size_t)(gt - src), closebuf);
            if (!close) goto done;
            if (buf_append(&out, gt + 1, (size_t)(close - (gt + 1))) != 0 ||
                buf_append_s(&out, entries) != 0 ||
                buf_append(&out, close, srclen - (size_t)(close - src)) != 0) goto oom;
        }
    } else {
        char parentpat[32];
        const char *a, *agt, *after;
        snprintf(parentpat, sizeof(parentpat), "<%s", create_after);
        a = mem_find(src, srclen, parentpat);
        if (!a) goto done;
        agt = (const char *)memchr(a, '>', srclen - (size_t)(a - src));
        if (!agt) goto done;
        after = agt + 1;
        snprintf(openbuf, sizeof(openbuf), "<%s count=\"%zu\">", tag, added);
        if (buf_append(&out, src, (size_t)(after - src)) != 0 ||
            buf_append_s(&out, openbuf) != 0 ||
            buf_append_s(&out, entries) != 0 ||
            buf_append_s(&out, closebuf) != 0 ||
            buf_append(&out, after, srclen - (size_t)(after - src)) != 0) goto oom;
    }

    *out_data = out.data;
    *out_len = out.len;
    return LXLSX_NO_ERROR;

oom:
    err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
done:
    free(out.data);
    return err;
}

/*
 * Inject one numFmt + one cellXf per unique requested number-format code into
 * styles.xml, assign each requesting change its new xf index, and produce a
 * styles.xml replacement. *produced is set to 1 when a replacement was made.
 */
static const char *h_align_str(uint8_t a)
{
    switch (a) {
    case 1: return "left";   case 2: return "center"; case 3: return "right";
    case 4: return "fill";   case 5: return "justify";
    case 6: return "centerContinuous"; case 7: return "distributed";
    default: return NULL;
    }
}

static const char *v_align_str(uint8_t a)
{
    switch (a) {
    case 8: return "top";    case 9: return "bottom"; case 10: return "center";
    case 11: return "justify"; case 12: return "distributed";
    default: return NULL;
    }
}

static int fmt_has_fill(const lxlsx_format *f)
{
    return f->pattern > 0 || f->fg_color != LXLSX_COLOR_UNSET;
}

static int emit_font_fragment(lxlsx_edit_buf *b, const lxlsx_format *f)
{
    char tmp[48];
    const char *name = f->font_name[0] ? f->font_name : "Calibri";
    if (buf_append_s(b, "<font>") != 0) return -1;
    if (f->bold && buf_append_s(b, "<b/>") != 0) return -1;
    if (f->italic && buf_append_s(b, "<i/>") != 0) return -1;
    if (f->font_strikeout && buf_append_s(b, "<strike/>") != 0) return -1;
    if (f->underline) {
        if (buf_append_s(b, f->underline == 2 ? "<u val=\"double\"/>" : "<u/>") != 0)
            return -1;
    }
    snprintf(tmp, sizeof(tmp), "<sz val=\"%g\"/>", f->font_size > 0 ? f->font_size : 11.0);
    if (buf_append_s(b, tmp) != 0) return -1;
    if (f->font_color != LXLSX_COLOR_UNSET) {
        snprintf(tmp, sizeof(tmp), "<color rgb=\"FF%06X\"/>",
                 (unsigned)(f->font_color & LXLSX_COLOR_MASK));
        if (buf_append_s(b, tmp) != 0) return -1;
    }
    if (buf_append_s(b, "<name val=\"") != 0 ||
        append_attr_escaped(b, name) != 0 ||
        buf_append_s(b, "\"/><family val=\"2\"/></font>") != 0)
        return -1;
    return 0;
}

static int emit_fill_fragment(lxlsx_edit_buf *b, const lxlsx_format *f)
{
    static const char *pats[] = {
        "none", "solid", "mediumGray", "darkGray", "lightGray",
        "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid",
        "darkTrellis", "lightHorizontal", "lightVertical", "lightDown",
        "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"
    };
    uint8_t p = f->pattern ? f->pattern : 1;  /* fg set but no pattern -> solid */
    char tmp[48];
    if (p >= sizeof(pats) / sizeof(pats[0])) p = 1;
    if (buf_append_s(b, "<fill><patternFill patternType=\"") != 0 ||
        buf_append_s(b, pats[p]) != 0 || buf_append_s(b, "\">") != 0)
        return -1;
    if (f->fg_color != LXLSX_COLOR_UNSET) {
        snprintf(tmp, sizeof(tmp), "<fgColor rgb=\"FF%06X\"/>",
                 (unsigned)(f->fg_color & LXLSX_COLOR_MASK));
        if (buf_append_s(b, tmp) != 0) return -1;
    }
    if (f->bg_color != LXLSX_COLOR_UNSET) {
        snprintf(tmp, sizeof(tmp), "<bgColor rgb=\"FF%06X\"/>",
                 (unsigned)(f->bg_color & LXLSX_COLOR_MASK));
        if (buf_append_s(b, tmp) != 0) return -1;
    } else if (buf_append_s(b, "<bgColor indexed=\"64\"/>") != 0) {
        return -1;
    }
    return buf_append_s(b, "</patternFill></fill>");
}

static int fmt_has_border(const lxlsx_format *f)
{
    return f->top || f->bottom || f->left || f->right ||
           f->diag_border || f->diag_type;
}

static const char *border_style_str(uint8_t s)
{
    static const char *styles[] = {
        "none", "thin", "medium", "dashed", "dotted", "thick", "double",
        "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
        "mediumDashDotDot", "slantDashDot"
    };
    if (s >= sizeof(styles) / sizeof(styles[0])) s = 1;
    return styles[s];
}

static int emit_sub_border(lxlsx_edit_buf *b, const char *type, uint8_t style,
                           lxlsx_color_t color)
{
    char tmp[48];
    if (!style) {
        snprintf(tmp, sizeof(tmp), "<%s/>", type);
        return buf_append_s(b, tmp);
    }
    snprintf(tmp, sizeof(tmp), "<%s style=\"%s\">", type, border_style_str(style));
    if (buf_append_s(b, tmp) != 0) return -1;
    if (color != LXLSX_COLOR_UNSET)
        snprintf(tmp, sizeof(tmp), "<color rgb=\"FF%06X\"/>", (unsigned)(color & LXLSX_COLOR_MASK));
    else
        snprintf(tmp, sizeof(tmp), "<color auto=\"1\"/>");
    if (buf_append_s(b, tmp) != 0) return -1;
    snprintf(tmp, sizeof(tmp), "</%s>", type);
    return buf_append_s(b, tmp);
}

static int emit_border_fragment(lxlsx_edit_buf *b, const lxlsx_format *f)
{
    uint8_t diag = f->diag_border;
    if (f->diag_type && !diag) diag = 1;  /* thin default, mirrors writer */
    if (buf_append_s(b, "<border") != 0) return -1;
    if (f->diag_type == 1 && buf_append_s(b, " diagonalUp=\"1\"") != 0) return -1;
    if (f->diag_type == 2 && buf_append_s(b, " diagonalDown=\"1\"") != 0) return -1;
    if (f->diag_type == 3 &&
        buf_append_s(b, " diagonalUp=\"1\" diagonalDown=\"1\"") != 0) return -1;
    if (buf_append_s(b, ">") != 0) return -1;
    if (emit_sub_border(b, "left", f->left, f->left_color) != 0) return -1;
    if (emit_sub_border(b, "right", f->right, f->right_color) != 0) return -1;
    if (emit_sub_border(b, "top", f->top, f->top_color) != 0) return -1;
    if (emit_sub_border(b, "bottom", f->bottom, f->bottom_color) != 0) return -1;
    if (emit_sub_border(b, "diagonal", diag, f->diag_color) != 0) return -1;
    return buf_append_s(b, "</border>");
}

static int emit_xf_fragment(lxlsx_edit_buf *b, size_t num_id, size_t font_id,
                            size_t fill_id, size_t border_id,
                            const lxlsx_format *f, int has_num)
{
    char head[160];
    const char *h = f ? h_align_str(f->text_h_align) : NULL;
    const char *v = f ? v_align_str(f->text_v_align) : NULL;
    int wrap = f ? f->text_wrap : 0;
    int has_align = (h || v || wrap);

    snprintf(head, sizeof(head),
             "<xf numFmtId=\"%zu\" fontId=\"%zu\" fillId=\"%zu\" borderId=\"%zu\" xfId=\"0\"",
             num_id, font_id, fill_id, border_id);
    if (buf_append_s(b, head) != 0) return -1;
    if (has_num && buf_append_s(b, " applyNumberFormat=\"1\"") != 0) return -1;
    if (f && buf_append_s(b, " applyFont=\"1\"") != 0) return -1;
    if (f && fmt_has_fill(f) && buf_append_s(b, " applyFill=\"1\"") != 0) return -1;
    if (f && fmt_has_border(f) && buf_append_s(b, " applyBorder=\"1\"") != 0) return -1;
    if (has_align && buf_append_s(b, " applyAlignment=\"1\"") != 0) return -1;
    if (!has_align)
        return buf_append_s(b, "/>");
    if (buf_append_s(b, "><alignment") != 0) return -1;
    if (h && (buf_append_s(b, " horizontal=\"") != 0 || buf_append_s(b, h) != 0 ||
              buf_append_s(b, "\"") != 0)) return -1;
    if (v && (buf_append_s(b, " vertical=\"") != 0 || buf_append_s(b, v) != 0 ||
              buf_append_s(b, "\"") != 0)) return -1;
    if (wrap && buf_append_s(b, " wrapText=\"1\"") != 0) return -1;
    return buf_append_s(b, "/></xf>");
}

/*
 * Inject the styles required by restyle/number-format edits into styles.xml:
 * one font/fill/numFmt/xf per unique applied format (deduped), plus a numFmt+xf
 * per unique standalone number-format code. Assigns each change its new xf index
 * and produces a styles.xml replacement.
 */
static lxlsx_error inject_cell_styles(
    lxlsx_edit_session *session, lxlsx_source_package_replacement *rep,
    int *produced)
{
    unsigned char *xml = NULL;
    size_t xmllen = 0;
    char *b1 = NULL, *b2 = NULL, *b3 = NULL, *final = NULL;
    size_t l1, l2, l3, finallen;
    lxlsx_edit_buf fonts = {0}, fills = {0}, borders = {0}, numfmts = {0}, xfents = {0};
    char *b4 = NULL; size_t l4;
    /* dedup tables */
    const lxlsx_format **fmts = NULL; size_t *fmt_xf = NULL, nfmt = 0, capfmt = 0;
    const char **codes = NULL; size_t *code_xf = NULL, ncode = 0, capcode = 0;
    size_t next_font, next_fill, next_border, next_xf, next_numid;
    size_t nfonts_add = 0, nfills_add = 0, nborders_add = 0, nnum_add = 0, nxf_add = 0;
    lxlsx_error err = LXLSX_NO_ERROR;
    int entry;
    size_t i, j;
    const char *mc, *gt, *p, *end;

    *produced = 0;

    /* Nothing to do unless some change carries a style or number format. */
    for (i = 0; i < session->change_count; i++)
        if (session->changes[i].style || session->changes[i].number_format)
            break;
    if (i == session->change_count)
        goto done;

    entry = lxlsx_source_package_find_first(session->package, "xl/styles.xml");
    if (entry < 0) { err = LXLSX_ERROR_ZIP_FILE_ADD; goto done; }
    err = lxlsx_source_package_read_entry(session->package, (size_t)entry,
                                          &xml, &xmllen);
    if (err != LXLSX_NO_ERROR) goto done;

    {
        const char *t;
        t = mem_find((const char *)xml, xmllen, "<fonts");
        gt = t ? (const char *)memchr(t, '>', xmllen - (size_t)(t - (const char *)xml)) : NULL;
        next_font = (t && gt) ? parse_count_attr(t, gt) : 1;
        t = mem_find((const char *)xml, xmllen, "<fills");
        gt = t ? (const char *)memchr(t, '>', xmllen - (size_t)(t - (const char *)xml)) : NULL;
        next_fill = (t && gt) ? parse_count_attr(t, gt) : 2;
        t = mem_find((const char *)xml, xmllen, "<borders");
        gt = t ? (const char *)memchr(t, '>', xmllen - (size_t)(t - (const char *)xml)) : NULL;
        next_border = (t && gt) ? parse_count_attr(t, gt) : 1;
        mc = mem_find((const char *)xml, xmllen, "<cellXfs");
        if (!mc) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        gt = (const char *)memchr(mc, '>', xmllen - (size_t)(mc - (const char *)xml));
        if (!gt) { err = LXLSX_ERROR_PARAMETER_VALIDATION; goto done; }
        next_xf = parse_count_attr(mc, gt);
    }

    /* Next custom numFmtId (custom ids are >= 164). */
    next_numid = 163;
    p = (const char *)xml;
    end = p + xmllen;
    while ((p = mem_find(p, (size_t)(end - p), "numFmtId=\"")) != NULL) {
        size_t val = 0;
        p += strlen("numFmtId=\"");
        while (p < end && *p >= '0' && *p <= '9')
            val = val * 10 + (size_t)(*p++ - '0');
        if (val > next_numid) next_numid = val;
    }
    next_numid += 1;

    for (i = 0; i < session->change_count; i++) {
        lxlsx_edit_change *ch = &session->changes[i];

        if (ch->style) {
            size_t font_id, fill_id = 0, border_id = 0, num_id = 0;
            int has_num = 0, found = 0;
            for (j = 0; j < nfmt; j++) {
                if (memcmp(fmts[j], ch->style, sizeof(lxlsx_format)) == 0) {
                    ch->style_index = (int)fmt_xf[j];
                    found = 1; break;
                }
            }
            if (found) continue;

            font_id = next_font++; nfonts_add++;
            if (emit_font_fragment(&fonts, ch->style) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            if (fmt_has_fill(ch->style)) {
                fill_id = next_fill++; nfills_add++;
                if (emit_fill_fragment(&fills, ch->style) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            }
            if (fmt_has_border(ch->style)) {
                border_id = next_border++; nborders_add++;
                if (emit_border_fragment(&borders, ch->style) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            }
            if (ch->style->num_format[0]) {
                char head[48];
                num_id = next_numid++; nnum_add++; has_num = 1;
                snprintf(head, sizeof(head), "<numFmt numFmtId=\"%zu\" formatCode=\"", num_id);
                if (buf_append_s(&numfmts, head) != 0 ||
                    append_attr_escaped(&numfmts, ch->style->num_format) != 0 ||
                    buf_append_s(&numfmts, "\"/>") != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            }
            if (emit_xf_fragment(&xfents, num_id, font_id, fill_id, border_id, ch->style, has_num) != 0) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done;
            }
            nxf_add++;
            ch->style_index = (int)next_xf++;

            if (nfmt >= capfmt) {
                size_t nc = capfmt ? capfmt * 2 : 8;
                const lxlsx_format **nf = (const lxlsx_format **)realloc(fmts, nc * sizeof(*nf));
                size_t *nx = (size_t *)realloc(fmt_xf, nc * sizeof(*nx));
                if (!nf || !nx) { free(nf); free(nx); err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
                fmts = nf; fmt_xf = nx; capfmt = nc;
            }
            fmts[nfmt] = ch->style; fmt_xf[nfmt] = (size_t)ch->style_index; nfmt++;
        } else if (ch->number_format) {
            char head[48];
            size_t num_id;
            int found = 0;
            for (j = 0; j < ncode; j++) {
                if (strcmp(codes[j], ch->number_format) == 0) {
                    ch->style_index = (int)code_xf[j]; found = 1; break;
                }
            }
            if (found) continue;

            num_id = next_numid++; nnum_add++;
            snprintf(head, sizeof(head), "<numFmt numFmtId=\"%zu\" formatCode=\"", num_id);
            if (buf_append_s(&numfmts, head) != 0 ||
                append_attr_escaped(&numfmts, ch->number_format) != 0 ||
                buf_append_s(&numfmts, "\"/>") != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            if (emit_xf_fragment(&xfents, num_id, 0, 0, 0, NULL, 1) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
            nxf_add++;
            ch->style_index = (int)next_xf++;

            if (ncode >= capcode) {
                size_t nc = capcode ? capcode * 2 : 8;
                const char **ncc = (const char **)realloc(codes, nc * sizeof(*ncc));
                size_t *nx = (size_t *)realloc(code_xf, nc * sizeof(*nx));
                if (!ncc || !nx) { free(ncc); free(nx); err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
                codes = ncc; code_xf = nx; capcode = nc;
            }
            codes[ncode] = ch->number_format; code_xf[ncode] = (size_t)ch->style_index; ncode++;
        }
    }

    if (nxf_add == 0)
        goto done;

    /* Splice each collection (order doesn't matter between distinct tags). */
    err = splice_collection((const char *)xml, xmllen, "fonts",
                            fonts.data ? fonts.data : "", nfonts_add, "styleSheet", &b1, &l1);
    if (err != LXLSX_NO_ERROR) goto done;
    err = splice_collection(b1, l1, "fills",
                            fills.data ? fills.data : "", nfills_add, "styleSheet", &b2, &l2);
    if (err != LXLSX_NO_ERROR) goto done;
    err = splice_collection(b2, l2, "borders",
                            borders.data ? borders.data : "", nborders_add, "styleSheet", &b3, &l3);
    if (err != LXLSX_NO_ERROR) goto done;
    err = splice_collection(b3, l3, "numFmts",
                            numfmts.data ? numfmts.data : "", nnum_add, "styleSheet", &b4, &l4);
    if (err != LXLSX_NO_ERROR) goto done;
    err = splice_collection(b4, l4, "cellXfs",
                            xfents.data ? xfents.data : "", nxf_add, "styleSheet", &final, &finallen);
    if (err != LXLSX_NO_ERROR) goto done;

    rep->entry_index = (size_t)entry;
    rep->data = (const unsigned char *)final;
    rep->size = finallen;
    final = NULL;
    *produced = 1;

done:
    free(fmts); free(fmt_xf); free(codes); free(code_xf);
    free(xml); free(b1); free(b2); free(b3); free(b4); free(final);
    free(fonts.data); free(fills.data); free(borders.data);
    free(numfmts.data); free(xfents.data);
    return err;
}

/* Largest N among existing xl/worksheets/sheetN.xml parts. */
static size_t max_existing_sheet_num(lxlsx_edit_session *session)
{
    size_t n = lxlsx_source_package_entry_count(session->package), i, maxv = 0;
    for (i = 0; i < n; i++) {
        const lxlsx_source_package_entry_info *info =
            lxlsx_source_package_entry_info_at(session->package, i);
        size_t pos = 19, v = 0;
        int has = 0;
        if (!info || info->name_len < 24) continue;
        if (memcmp(info->name, "xl/worksheets/sheet", 19) != 0) continue;
        while (pos < info->name_len && info->name[pos] >= '0' && info->name[pos] <= '9') {
            v = v * 10 + (size_t)(info->name[pos] - '0'); pos++; has = 1;
        }
        if (has && info->name_len - pos == 4 &&
            memcmp(info->name + pos, ".xml", 4) == 0 && v > maxv)
            maxv = v;
    }
    return maxv;
}

/* Largest integer following each occurrence of `needle` in [buf, buf+len). */
static size_t scan_max_after(const unsigned char *buf, size_t len, const char *needle)
{
    const char *p = (const char *)buf, *end = p + len;
    size_t maxv = 0;
    while ((p = mem_find(p, (size_t)(end - p), needle)) != NULL) {
        size_t v = 0;
        p += strlen(needle);
        while (p < end && *p >= '0' && *p <= '9') v = v * 10 + (size_t)(*p++ - '0');
        if (v > maxv) maxv = v;
    }
    return maxv;
}

/* Read a package part, insert `insert` before `close_tag`, produce a replacement. */
static lxlsx_error part_insert_before(lxlsx_edit_session *session,
                                      const char *part, const char *close_tag,
                                      const char *insert,
                                      lxlsx_source_package_replacement *rep)
{
    unsigned char *xml = NULL;
    size_t xmllen = 0;
    int idx;
    const char *pos;
    lxlsx_edit_buf out = {0};
    lxlsx_error err;

    idx = lxlsx_source_package_find_first(session->package, part);
    if (idx < 0) return LXLSX_ERROR_ZIP_FILE_ADD;
    err = lxlsx_source_package_read_entry(session->package, (size_t)idx, &xml, &xmllen);
    if (err != LXLSX_NO_ERROR) return err;
    pos = mem_find((const char *)xml, xmllen, close_tag);
    if (!pos) { free(xml); return LXLSX_ERROR_PARAMETER_VALIDATION; }

    if (buf_append(&out, (const char *)xml, (size_t)(pos - (const char *)xml)) != 0 ||
        buf_append_s(&out, insert) != 0 ||
        buf_append(&out, pos, xmllen - (size_t)(pos - (const char *)xml)) != 0) {
        free(xml); free(out.data); return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    free(xml);
    rep->entry_index = (size_t)idx;
    rep->data = (const unsigned char *)out.data;
    rep->size = out.len;
    return LXLSX_NO_ERROR;
}

#define LXLSX_WS_REL_TYPE "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
#define LXLSX_WS_CT "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

/*
 * Multi-part composition accumulator. Collects brand-new parts (additions) and
 * per-part fragment insertions (materialised as replacements at finalize time),
 * aggregating every insertion that targets the same part into ONE replacement.
 *
 * That aggregation is the point: several features patch the same metadata part
 * (e.g. [Content_Types].xml gets a worksheet Override from add-sheet and an
 * image Default from add-image). Patching it twice independently — each off the
 * original bytes — would drop one. A brand-new rels part is just an addition, so
 * no separate "create rels" path is needed. add-sheet uses this today;
 * add-image/chart build on the same API. Borrowed `data` pointers handed to
 * composer_add_part must outlive the save (callers pass session-owned buffers).
 */
typedef struct {
    char          *part;       /* host part to patch (owned) */
    char          *close_tag;  /* token to insert the fragment before (owned) */
    lxlsx_edit_buf frag;       /* aggregated fragment */
} lxlsx_composer_patch;

typedef struct {
    lxlsx_edit_session            *session;
    lxlsx_source_package_addition *adds;   /* names owned; data borrowed */
    size_t                         add_count, add_cap;
    lxlsx_composer_patch          *patches;
    size_t                         patch_count, patch_cap;
    int                            oom;    /* sticky out-of-memory flag */
} lxlsx_composer;

static void composer_init(lxlsx_composer *c, lxlsx_edit_session *session)
{
    memset(c, 0, sizeof(*c));
    c->session = session;
}

static void composer_free(lxlsx_composer *c)
{
    size_t i;
    for (i = 0; i < c->add_count; i++)
        free(c->adds[i].name);
    free(c->adds);
    for (i = 0; i < c->patch_count; i++) {
        free(c->patches[i].part);
        free(c->patches[i].close_tag);
        free(c->patches[i].frag.data);
    }
    free(c->patches);
    memset(c, 0, sizeof(*c));
}

/* Append a brand-new part. `name` is copied; `data` is borrowed. */
static void composer_add_part(lxlsx_composer *c, const char *name,
                              const unsigned char *data, size_t size)
{
    lxlsx_source_package_addition *a;
    if (c->oom)
        return;
    if (c->add_count >= c->add_cap) {
        size_t cap = c->add_cap ? c->add_cap * 2 : 4;
        lxlsx_source_package_addition *next = (lxlsx_source_package_addition *)
            realloc(c->adds, cap * sizeof(*next));
        if (!next) { c->oom = 1; return; }
        c->adds = next;
        c->add_cap = cap;
    }
    a = &c->adds[c->add_count];
    a->name = strdup(name);
    if (!a->name) { c->oom = 1; return; }
    a->data = data;
    a->size = size;
    c->add_count++;
}

/* Return the aggregation buffer for `part` (created on first use), so the caller
 * can append fragments to it. Insertions accumulate before `close_tag`. NULL on
 * OOM (the sticky flag is also set, so callers may defer the check). */
static lxlsx_edit_buf *composer_part_buf(lxlsx_composer *c, const char *part,
                                         const char *close_tag)
{
    size_t i;
    lxlsx_composer_patch *p;
    if (c->oom)
        return NULL;
    for (i = 0; i < c->patch_count; i++) {
        if (strcmp(c->patches[i].part, part) == 0)
            return &c->patches[i].frag;
    }
    if (c->patch_count >= c->patch_cap) {
        size_t cap = c->patch_cap ? c->patch_cap * 2 : 4;
        lxlsx_composer_patch *next = (lxlsx_composer_patch *)
            realloc(c->patches, cap * sizeof(*next));
        if (!next) { c->oom = 1; return NULL; }
        c->patches = next;
        c->patch_cap = cap;
    }
    p = &c->patches[c->patch_count];
    memset(p, 0, sizeof(*p));
    p->part = strdup(part);
    p->close_tag = strdup(close_tag);
    if (!p->part || !p->close_tag) {
        free(p->part);
        free(p->close_tag);
        c->oom = 1;
        return NULL;
    }
    c->patch_count++;
    return &p->frag;
}

/* Materialise: additions array (caller frees the array, not the borrowed data)
 * plus one replacement per patched part (caller frees each .data and the
 * array). */
static lxlsx_error composer_finalize(lxlsx_composer *c,
                                     lxlsx_source_package_addition **add_out,
                                     size_t *n_add_out,
                                     lxlsx_source_package_replacement **rep_out,
                                     size_t *n_rep_out)
{
    lxlsx_source_package_addition *adds = NULL;
    lxlsx_source_package_replacement *reps = NULL;
    size_t i, produced = 0;
    lxlsx_error err;

    *add_out = NULL; *n_add_out = 0; *rep_out = NULL; *n_rep_out = 0;
    if (c->oom)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    if (c->add_count) {
        adds = (lxlsx_source_package_addition *)calloc(c->add_count, sizeof(*adds));
        if (!adds)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        memcpy(adds, c->adds, c->add_count * sizeof(*adds));
    }
    if (c->patch_count) {
        reps = (lxlsx_source_package_replacement *)
            calloc(c->patch_count, sizeof(*reps));
        if (!reps) { free(adds); return LXLSX_ERROR_MEMORY_MALLOC_FAILED; }
    }
    for (i = 0; i < c->patch_count; i++) {
        err = part_insert_before(c->session, c->patches[i].part,
                                 c->patches[i].close_tag,
                                 c->patches[i].frag.data ? c->patches[i].frag.data : "",
                                 &reps[produced]);
        if (err != LXLSX_NO_ERROR) {
            size_t j;
            for (j = 0; j < produced; j++)
                free((void *)reps[j].data);
            free(reps);
            free(adds);
            return err;
        }
        produced++;
    }

    *add_out = adds;
    *n_add_out = c->add_count;
    *rep_out = reps;
    *n_rep_out = produced;
    return LXLSX_NO_ERROR;
}

/*
 * Register the session's new worksheets onto a composer: a new sheetN.xml part
 * each, plus <sheet>/<Relationship>/<Override> fragments for workbook.xml, its
 * rels and [Content_Types].xml.
 */
static lxlsx_error prepare_new_sheets(lxlsx_edit_session *session,
                                      lxlsx_composer *c)
{
    lxlsx_edit_buf *wb, *rel, *ct;
    unsigned char *wbxml = NULL, *relxml = NULL;
    size_t wblen = 0, rellen = 0;
    size_t next_num, next_sid, next_rid, i;
    int idx;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (session->new_sheet_count == 0)
        return LXLSX_NO_ERROR;

    next_num = max_existing_sheet_num(session) + 1;

    /* Parse next sheetId from workbook.xml and next rId from its rels. */
    idx = lxlsx_source_package_find_first(session->package, "xl/workbook.xml");
    if (idx < 0) { err = LXLSX_ERROR_ZIP_FILE_ADD; goto done; }
    err = lxlsx_source_package_read_entry(session->package, (size_t)idx, &wbxml, &wblen);
    if (err != LXLSX_NO_ERROR) goto done;
    next_sid = scan_max_after(wbxml, wblen, "sheetId=\"") + 1;

    idx = lxlsx_source_package_find_first(session->package, "xl/_rels/workbook.xml.rels");
    if (idx < 0) { err = LXLSX_ERROR_ZIP_FILE_ADD; goto done; }
    err = lxlsx_source_package_read_entry(session->package, (size_t)idx, &relxml, &rellen);
    if (err != LXLSX_NO_ERROR) goto done;
    next_rid = scan_max_after(relxml, rellen, "Id=\"rId") + 1;

    wb  = composer_part_buf(c, "xl/workbook.xml", "</sheets>");
    rel = composer_part_buf(c, "xl/_rels/workbook.xml.rels", "</Relationships>");
    ct  = composer_part_buf(c, "[Content_Types].xml", "</Types>");
    if (!wb || !rel || !ct) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }

    for (i = 0; i < session->new_sheet_count; i++) {
        lxlsx_edit_new_sheet *s = &session->new_sheets[i];
        char fn[48], frag[256];
        size_t num = next_num + i, sid = next_sid + i, rid = next_rid + i;

        snprintf(fn, sizeof(fn), "xl/worksheets/sheet%zu.xml", num);
        composer_add_part(c, fn, (const unsigned char *)s->xml, s->xml_len);

        if (buf_append_s(wb, "<sheet name=\"") != 0 ||
            append_attr_escaped(wb, s->name) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
        snprintf(frag, sizeof(frag), "\" sheetId=\"%zu\" r:id=\"rId%zu\"/>", sid, rid);
        if (buf_append_s(wb, frag) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }

        snprintf(frag, sizeof(frag),
                 "<Relationship Id=\"rId%zu\" Type=\"%s\" Target=\"worksheets/sheet%zu.xml\"/>",
                 rid, LXLSX_WS_REL_TYPE, num);
        if (buf_append_s(rel, frag) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }

        snprintf(frag, sizeof(frag),
                 "<Override PartName=\"/xl/worksheets/sheet%zu.xml\" ContentType=\"%s\"/>",
                 num, LXLSX_WS_CT);
        if (buf_append_s(ct, frag) != 0) { err = LXLSX_ERROR_MEMORY_MALLOC_FAILED; goto done; }
    }

    if (c->oom)
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;

done:
    free(wbxml);
    free(relxml);
    return err;
}

lxlsx_error lxlsx_edit_save_as(lxlsx_edit_session *session, const char *path)
{
    lxlsx_dirty_sheet *dirty_sheets = NULL;
    size_t dirty_count = 0;
    size_t dirty_cap = 0;
    lxlsx_source_package_replacement *replacements = NULL;
    lxlsx_source_package_replacement styles_rep;
    lxlsx_source_package_replacement *meta_reps = NULL;
    lxlsx_source_package_addition *additions = NULL;
    lxlsx_composer composer;
    size_t add_count = 0, meta_count = 0;
    int styles_produced = 0;
    size_t rep_count;
    lxlsx_error err = LXLSX_NO_ERROR;
    size_t i;

    memset(&styles_rep, 0, sizeof(styles_rep));

    if (!session || !path)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (session->change_count == 0 && session->merge_count == 0 &&
        session->col_count == 0 && session->row_dim_count == 0 &&
        session->new_sheet_count == 0)
        return lxlsx_source_package_save_copy(session->package, path);

    composer_init(&composer, session);

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

    for (i = 0; i < session->merge_count; i++) {
        lxlsx_dirty_sheet *sheet = NULL;
        err = get_dirty_sheet(session, &dirty_sheets, &dirty_count, &dirty_cap,
                              session->merges[i].target, &sheet);
        if (err != LXLSX_NO_ERROR)
            goto done;
        err = dirty_sheet_add_merge(sheet, &session->merges[i]);
        if (err != LXLSX_NO_ERROR)
            goto done;
    }

    for (i = 0; i < session->col_count; i++) {
        lxlsx_dirty_sheet *sheet = NULL;
        err = get_dirty_sheet(session, &dirty_sheets, &dirty_count, &dirty_cap,
                              session->cols[i].target, &sheet);
        if (err != LXLSX_NO_ERROR)
            goto done;
        err = dirty_sheet_add_col(sheet, &session->cols[i]);
        if (err != LXLSX_NO_ERROR)
            goto done;
    }

    for (i = 0; i < session->row_dim_count; i++) {
        lxlsx_dirty_sheet *sheet = NULL;
        err = get_dirty_sheet(session, &dirty_sheets, &dirty_count, &dirty_cap,
                              session->row_dims[i].target, &sheet);
        if (err != LXLSX_NO_ERROR)
            goto done;
        err = dirty_sheet_add_row_dim(sheet, &session->row_dims[i]);
        if (err != LXLSX_NO_ERROR)
            goto done;
    }

    /* Resolve number-format restyles into styles.xml (sets per-change style
     * indices that the cell patcher below reads). */
    err = inject_cell_styles(session, &styles_rep, &styles_produced);
    if (err != LXLSX_NO_ERROR)
        goto done;

    /* Build new-sheet parts + their workbook/rels/content-type registrations
     * onto the composer (shared with future add-image/chart compositions). */
    err = prepare_new_sheets(session, &composer);
    if (err != LXLSX_NO_ERROR)
        goto done;

    for (i = 0; i < dirty_count; i++) {
        if (dirty_sheets[i].change_count > 0) {
            err = patch_xml_with_changes(&dirty_sheets[i].xml,
                                         &dirty_sheets[i].xml_len,
                                         dirty_sheets[i].changes,
                                         dirty_sheets[i].change_count);
            if (err != LXLSX_NO_ERROR)
                goto done;
        }
        if (dirty_sheets[i].merge_count > 0) {
            err = patch_xml_with_merges(&dirty_sheets[i].xml,
                                        &dirty_sheets[i].xml_len,
                                        dirty_sheets[i].merges,
                                        dirty_sheets[i].merge_count);
            if (err != LXLSX_NO_ERROR)
                goto done;
        }
        if (dirty_sheets[i].col_count > 0) {
            err = patch_xml_with_cols(&dirty_sheets[i].xml,
                                      &dirty_sheets[i].xml_len,
                                      dirty_sheets[i].cols,
                                      dirty_sheets[i].col_count);
            if (err != LXLSX_NO_ERROR)
                goto done;
        }
        if (dirty_sheets[i].row_dim_count > 0) {
            err = patch_xml_with_row_dims(&dirty_sheets[i].xml,
                                          &dirty_sheets[i].xml_len,
                                          dirty_sheets[i].row_dims,
                                          dirty_sheets[i].row_dim_count);
            if (err != LXLSX_NO_ERROR)
                goto done;
        }
    }

    /* Turn the composer's accumulated parts/fragments into additions + one
     * replacement per patched metadata part. */
    err = composer_finalize(&composer, &additions, &add_count,
                            &meta_reps, &meta_count);
    if (err != LXLSX_NO_ERROR)
        goto done;

    replacements = (lxlsx_source_package_replacement *)calloc(
        dirty_count + 1 + meta_count, sizeof(*replacements));
    if (!replacements) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto done;
    }

    for (i = 0; i < dirty_count; i++) {
        replacements[i].entry_index = dirty_sheets[i].entry_index;
        replacements[i].data = dirty_sheets[i].xml;
        replacements[i].size = dirty_sheets[i].xml_len;
    }
    rep_count = dirty_count;
    if (styles_produced) {
        replacements[rep_count++] = styles_rep;
    }
    for (i = 0; i < meta_count; i++) {
        replacements[rep_count++] = meta_reps[i];
    }

    err = lxlsx_source_package_save_with_changes(session->package, path,
                                                 replacements, rep_count,
                                                 additions, add_count);

done:
    free(replacements);
    free((void *)styles_rep.data);
    for (i = 0; i < meta_count; i++)
        free((void *)meta_reps[i].data);
    free(meta_reps);
    free(additions);
    composer_free(&composer);
    free_dirty_sheets(dirty_sheets, dirty_count);
    return err;
}
