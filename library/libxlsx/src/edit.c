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
} lxlsx_dirty_sheet;

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

static int buf_append_c(lxlsx_edit_buf *buf, char c)
{
    return buf_append(buf, &c, 1);
}

static int buf_append_s(lxlsx_edit_buf *buf, const char *str)
{
    return buf_append(buf, str, strlen(str));
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

static int append_xml_escaped(lxlsx_edit_buf *buf, const char *str)
{
    const char *p;
    if (!str)
        return 0;
    for (p = str; *p; p++) {
        switch (*p) {
        case '&':
            if (buf_append_s(buf, "&amp;") != 0) return -1;
            break;
        case '<':
            if (buf_append_s(buf, "&lt;") != 0) return -1;
            break;
        case '>':
            if (buf_append_s(buf, "&gt;") != 0) return -1;
            break;
        default:
            if (buf_append_c(buf, *p) != 0) return -1;
            break;
        }
    }
    return 0;
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

static int xml_name_eq(const char *start, size_t len, const char *name)
{
    return strlen(name) == len && strncmp(start, name, len) == 0;
}

static char *extract_attr(const char *tag_start, const char *tag_end,
                          const char *name)
{
    const char *p = tag_start;

    while (p < tag_end && *p != '<')
        p++;
    while (p < tag_end && !isspace((unsigned char)*p) && *p != '>')
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
        if (xml_name_eq(attr_start, (size_t)(attr_end - attr_start), name))
            return dup_range(value_start, (size_t)(p - value_start));
        p++;
    }

    return NULL;
}

static int attr_equals(const char *tag_start, const char *tag_end,
                       const char *name, const char *value)
{
    char *attr = extract_attr(tag_start, tag_end, name);
    int ok = attr && strcmp(attr, value) == 0;
    free(attr);
    return ok;
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

static const char *find_tag_end(const char *start, const char *limit)
{
    const char *p;
    for (p = start; p < limit; p++) {
        if (*p == '>')
            return p;
    }
    return NULL;
}

static int find_cell(const char *xml, size_t len, const char *ref,
                     const char **cell_start, const char **cell_end,
                     const char **tag_end)
{
    const char *limit = xml + len;
    const char *p = xml;

    while ((p = strstr(p, "<c")) != NULL && p < limit) {
        const char *end;
        const char *close;
        if (p + 2 >= limit)
            return 0;
        if (!(isspace((unsigned char)p[2]) || p[2] == '>' || p[2] == '/')) {
            p += 2;
            continue;
        }
        end = find_tag_end(p, limit);
        if (!end)
            return 0;
        if (!attr_equals(p, end, "r", ref)) {
            p = end + 1;
            continue;
        }
        *cell_start = p;
        *tag_end = end;
        if (is_self_closing(p, end)) {
            *cell_end = end + 1;
            return 1;
        }
        close = strstr(end + 1, "</c>");
        if (!close || close > limit)
            return 0;
        *cell_end = close + 4;
        return 1;
    }

    return 0;
}

static int find_row(const char *xml, size_t len, const char *row_ref,
                    const char **row_start, const char **row_end,
                    const char **tag_end)
{
    const char *limit = xml + len;
    const char *p = xml;

    while ((p = strstr(p, "<row")) != NULL && p < limit) {
        const char *end;
        const char *close;
        if (p + 4 >= limit)
            return 0;
        if (!(isspace((unsigned char)p[4]) || p[4] == '>' || p[4] == '/')) {
            p += 4;
            continue;
        }
        end = find_tag_end(p, limit);
        if (!end)
            return 0;
        if (!attr_equals(p, end, "r", row_ref)) {
            p = end + 1;
            continue;
        }
        *row_start = p;
        *tag_end = end;
        if (is_self_closing(p, end)) {
            *row_end = end + 1;
            return 1;
        }
        close = strstr(end + 1, "</row>");
        if (!close || close > limit)
            return 0;
        *row_end = close + 6;
        return 1;
    }

    return 0;
}

static int find_sheet_data(const char *xml, size_t len,
                           const char **sheet_start,
                           const char **sheet_end,
                           const char **tag_end)
{
    const char *limit = xml + len;
    const char *p = strstr(xml, "<sheetData");
    const char *end;
    const char *close;

    if (!p || p >= limit)
        return 0;
    end = find_tag_end(p, limit);
    if (!end)
        return 0;
    *sheet_start = p;
    *tag_end = end;
    if (is_self_closing(p, end)) {
        *sheet_end = end + 1;
        return 1;
    }
    close = strstr(end + 1, "</sheetData>");
    if (!close || close > limit)
        return 0;
    *sheet_end = close + 12;
    return 1;
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

    if (buf_append_s(&buf, "</c>") != 0)
        goto fail;
    return buf.data;

fail:
    free(buf.data);
    return NULL;
}

static lxlsx_error replace_range(unsigned char **xml, size_t *xml_len,
                                 const char *start, const char *end,
                                 const char *replacement)
{
    const char *old_xml = (const char *)*xml;
    size_t prefix = (size_t)(start - old_xml);
    size_t suffix = *xml_len - (size_t)(end - old_xml);
    size_t repl_len = strlen(replacement);
    unsigned char *next;

    next = (unsigned char *)malloc(prefix + repl_len + suffix + 1);
    if (!next)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    memcpy(next, old_xml, prefix);
    memcpy(next + prefix, replacement, repl_len);
    memcpy(next + prefix + repl_len, end, suffix);
    next[prefix + repl_len + suffix] = 0;

    free(*xml);
    *xml = next;
    *xml_len = prefix + repl_len + suffix;
    return LXLSX_NO_ERROR;
}

static lxlsx_error insert_at(unsigned char **xml, size_t *xml_len,
                             const char *at, const char *insert)
{
    return replace_range(xml, xml_len, at, at, insert);
}

static lxlsx_error patch_xml_with_change(unsigned char **xml, size_t *xml_len,
                                         const lxlsx_edit_change *change)
{
    char *ref = NULL;
    char *row_ref = NULL;
    char *style = NULL;
    char *cell_xml = NULL;
    const char *cell_start = NULL;
    const char *cell_end = NULL;
    const char *cell_tag_end = NULL;
    const char *row_start = NULL;
    const char *row_end = NULL;
    const char *row_tag_end = NULL;
    const char *sheet_start = NULL;
    const char *sheet_end = NULL;
    const char *sheet_tag_end = NULL;
    lxlsx_error err = LXLSX_NO_ERROR;

    ref = make_cell_ref(change->row, change->col);
    row_ref = make_row_ref(change->row);
    if (!ref || !row_ref) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto done;
    }

    if (find_cell((const char *)*xml, *xml_len, ref,
                  &cell_start, &cell_end, &cell_tag_end)) {
        style = extract_attr(cell_start, cell_tag_end, "s");
        cell_xml = build_cell_xml(change, ref, style);
        if (!cell_xml) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }
        err = replace_range(xml, xml_len, cell_start, cell_end, cell_xml);
        goto done;
    }

    cell_xml = build_cell_xml(change, ref, NULL);
    if (!cell_xml) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto done;
    }

    if (find_row((const char *)*xml, *xml_len, row_ref,
                 &row_start, &row_end, &row_tag_end)) {
        if (is_self_closing(row_start, row_tag_end)) {
            lxlsx_edit_buf row = {0};
            const char *slash = row_tag_end;
            while (slash > row_start && *slash != '/')
                slash--;
            if (buf_append(&row, row_start, (size_t)(slash - row_start)) != 0 ||
                buf_append_s(&row, ">") != 0 ||
                buf_append_s(&row, cell_xml) != 0 ||
                buf_append_s(&row, "</row>") != 0) {
                free(row.data);
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto done;
            }
            err = replace_range(xml, xml_len, row_start, row_end, row.data);
            free(row.data);
        } else {
            const char *insert_pos = row_end - 6;
            err = insert_at(xml, xml_len, insert_pos, cell_xml);
        }
        goto done;
    }

    if (!find_sheet_data((const char *)*xml, *xml_len,
                         &sheet_start, &sheet_end, &sheet_tag_end)) {
        err = LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
        goto done;
    }

    {
        lxlsx_edit_buf row = {0};
        if (buf_appendf(&row, "<row r=\"%s\">", row_ref) != 0 ||
            buf_append_s(&row, cell_xml) != 0 ||
            buf_append_s(&row, "</row>") != 0) {
            free(row.data);
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto done;
        }

        if (is_self_closing(sheet_start, sheet_tag_end)) {
            lxlsx_edit_buf sheet = {0};
            if (buf_append(&sheet, sheet_start,
                           (size_t)(sheet_tag_end - sheet_start - 1)) != 0 ||
                buf_append_s(&sheet, ">") != 0 ||
                buf_append_s(&sheet, row.data) != 0 ||
                buf_append_s(&sheet, "</sheetData>") != 0) {
                free(row.data);
                free(sheet.data);
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto done;
            }
            err = replace_range(xml, xml_len, sheet_start, sheet_end, sheet.data);
            free(sheet.data);
        } else {
            err = insert_at(xml, xml_len, sheet_end - 12, row.data);
        }
        free(row.data);
    }

done:
    free(ref);
    free(row_ref);
    free(style);
    free(cell_xml);
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
        err = patch_xml_with_change(&sheet->xml, &sheet->xml_len,
                                    &session->changes[i]);
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
