#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"

/* ------------------------------------------------------------------------- */
/* Helpers                                                                   */
/* ------------------------------------------------------------------------- */

static int append_buf(char **buf, size_t *len, size_t *cap, const char *src, size_t n)
{
    size_t need = *len + n + 1;
    if (need > *cap) {
        size_t nc = *cap ? *cap : 64;
        char *nb;
        while (nc < need) nc *= 2;
        nb = (char *)realloc(*buf, nc);
        if (!nb) return -1;
        *buf = nb;
        *cap = nc;
    }
    memcpy(*buf + *len, src, n);
    *len += n;
    (*buf)[*len] = 0;
    return 0;
}

static void parse_cell_ref(const char *ref, size_t *out_row, size_t *out_col)
{
    size_t col = 0, row = 0;
    const char *p = ref;
    if (out_row) *out_row = 0;
    if (out_col) *out_col = 0;
    if (!p) return;

    while (*p && ((*p >= 'A' && *p <= 'Z') || (*p >= 'a' && *p <= 'z'))) {
        col = col * 26 + (toupper((unsigned char)*p) - 'A' + 1);
        p++;
    }
    while (*p && *p >= '0' && *p <= '9') {
        row = row * 10 + (size_t)(*p - '0');
        p++;
    }
    if (out_row) *out_row = row;
    if (out_col) *out_col = col;
}

/* Excel serial date -> Unix timestamp.
 * 1900 system: serial 1 == 1900-01-01, but Excel treats 1900 as a leap year
 * (Lotus 1-2-3 bug), so we anchor at 1899-12-30 to keep all serials >= 61
 * correct. Serials 0..59 inherit the historical Lotus interpretation.
 * 1904 system: serial 0 == 1904-01-01. */
int64_t lxr_excel_serial_to_unix(double serial, int uses_1904)
{
    static const int64_t EPOCH_1900 = -2209161600LL; /* 1899-12-30T00:00:00Z */
    static const int64_t EPOCH_1904 = -2082844800LL; /* 1904-01-01T00:00:00Z */
    int64_t base = uses_1904 ? EPOCH_1904 : EPOCH_1900;
    return base + (int64_t)(serial * 86400.0);
}

/* ------------------------------------------------------------------------- */
/* Cell type inference                                                       */
/* ------------------------------------------------------------------------- */

static void emit_cell(lxr_worksheet *ws, lxr_cell *out)
{
    const lxr_styles *st = ws->wb ? ws->wb->styles : NULL;
    const lxr_xf     *xf = NULL;
    const char       *t  = ws->cell_t;

    memset(out, 0, sizeof(*out));
    out->row      = ws->cell_row;
    out->col      = ws->cell_col;
    out->style_id = ws->cell_style_id;
    out->raw.ptr  = ws->cell_value;
    out->raw.len  = ws->cell_value_len;

    if (st) xf = lxr_styles_get_xf(st, ws->cell_style_id);

    /* Empty cell (no <v> and no inline string and no formula) */
    if (!ws->cell_has_formula && !ws->cell_has_inline &&
        ws->cell_value_len == 0) {
        out->type = LXR_CELL_BLANK;
        return;
    }

    if (ws->cell_has_formula) {
        out->type = LXR_CELL_FORMULA;
        out->value.formula.formula.ptr = ws->cell_formula;
        out->value.formula.formula.len = ws->cell_formula_len;
        out->value.formula.cached.ptr  = ws->cell_value;
        out->value.formula.cached.len  = ws->cell_value_len;
        return;
    }

    if (t[0] == 0 || strcmp(t, "n") == 0) {
        if (xf && (xf->category == LXR_FMT_CATEGORY_DATE ||
                   xf->category == LXR_FMT_CATEGORY_TIME ||
                   xf->category == LXR_FMT_CATEGORY_DATETIME)) {
            double serial = ws->cell_value ? strtod(ws->cell_value, NULL) : 0.0;
            out->type = LXR_CELL_DATETIME;
            out->value.unix_timestamp =
                lxr_excel_serial_to_unix(serial,
                                         ws->wb ? ws->wb->uses_1904 : 0);
        } else {
            out->type = LXR_CELL_NUMBER;
            out->value.number = ws->cell_value ? strtod(ws->cell_value, NULL) : 0.0;
        }
        return;
    }

    if (strcmp(t, "s") == 0) {
        uint32_t idx = ws->cell_value
            ? (uint32_t)strtoul(ws->cell_value, NULL, 10) : 0;
        const char *s = ws->wb && ws->wb->sst
            ? lxr_sst_get(ws->wb->sst, idx) : NULL;
        out->type = LXR_CELL_STRING;
        if (s) {
            out->value.string.ptr = s;
            out->value.string.len = strlen(s);
        } else {
            out->value.string.ptr = "";
            out->value.string.len = 0;
        }
        return;
    }

    if (strcmp(t, "inlineStr") == 0) {
        out->type = LXR_CELL_INLINE_STRING;
        out->value.string.ptr = ws->cell_inline;
        out->value.string.len = ws->cell_inline_len;
        return;
    }

    if (strcmp(t, "str") == 0) {
        out->type = LXR_CELL_STRING;
        out->value.string.ptr = ws->cell_value ? ws->cell_value : "";
        out->value.string.len = ws->cell_value_len;
        return;
    }

    if (strcmp(t, "b") == 0) {
        out->type = LXR_CELL_BOOLEAN;
        out->value.boolean = (ws->cell_value && ws->cell_value[0] == '1') ? 1 : 0;
        return;
    }

    if (strcmp(t, "e") == 0) {
        size_t n = ws->cell_value_len;
        if (n >= sizeof(out->value.error_code))
            n = sizeof(out->value.error_code) - 1;
        out->type = LXR_CELL_ERROR;
        if (ws->cell_value) memcpy(out->value.error_code, ws->cell_value, n);
        out->value.error_code[n] = 0;
        return;
    }

    /* Unknown 't': fall back to string. */
    out->type = LXR_CELL_STRING;
    out->value.string.ptr = ws->cell_value ? ws->cell_value : "";
    out->value.string.len = ws->cell_value_len;
}

/* ------------------------------------------------------------------------- */
/* Cell reset                                                                */
/* ------------------------------------------------------------------------- */

static void reset_cell(lxr_worksheet *ws)
{
    ws->cell_t[0] = 0;
    ws->cell_ref[0] = 0;
    ws->cell_style_id = 0;
    ws->cell_value_len = 0;
    if (ws->cell_value) ws->cell_value[0] = 0;
    ws->cell_formula_len = 0;
    if (ws->cell_formula) ws->cell_formula[0] = 0;
    ws->cell_inline_len = 0;
    if (ws->cell_inline) ws->cell_inline[0] = 0;
    ws->cell_has_formula = 0;
    ws->cell_has_inline = 0;
    ws->cell_row = 0;
    ws->cell_col = 0;
}

/* ------------------------------------------------------------------------- */
/* SAX dispatch                                                              */
/* ------------------------------------------------------------------------- */

static void copy_attr(char *dst, size_t cap, const char *src)
{
    size_t n;
    if (!dst || cap == 0) return;
    if (!src) { dst[0] = 0; return; }
    n = strlen(src);
    if (n >= cap) n = cap - 1;
    memcpy(dst, src, n);
    dst[n] = 0;
}

static void deliver_cell(lxr_worksheet *ws)
{
    if (ws->pull_mode == LXR_WS_PULL_CELL) {
        ws->pending_cell = 1;
        lxr_xml_pump_suspend(ws->pump);
        return;
    }
    if (ws->user_cell_cb) {
        lxr_cell c;
        emit_cell(ws, &c);
        if (ws->user_cell_cb(&c, ws->user_data) != 0) {
            ws->callback_stop = 1;
            lxr_xml_pump_suspend(ws->pump);
        }
    }
}

static void deliver_row_end(lxr_worksheet *ws)
{
    if (ws->pull_mode == LXR_WS_PULL_CELL) {
        /* No more cells in this row */
        ws->pending_row_end = 1;
        lxr_xml_pump_suspend(ws->pump);
        return;
    }
    if (ws->user_row_cb) {
        if (ws->user_row_cb(ws->row_nr, ws->max_col_seen, ws->user_data) != 0) {
            ws->callback_stop = 1;
            lxr_xml_pump_suspend(ws->pump);
        }
    }
}

static void on_start(void *ud, const char *name, const char **attrs)
{
    lxr_worksheet *ws = (lxr_worksheet *)ud;

    if (ws->state == LXR_WS_SKIP) {
        ws->skip_depth++;
        return;
    }

    switch (ws->state) {
    case LXR_WS_INIT:
        if (lxr_xml_name_eq(name, "worksheet")) ws->state = LXR_WS_IN_WORKSHEET;
        break;

    case LXR_WS_IN_WORKSHEET:
        if (lxr_xml_name_eq(name, "sheetData")) ws->state = LXR_WS_IN_SHEETDATA;
        break;

    case LXR_WS_IN_SHEETDATA:
        if (lxr_xml_name_eq(name, "row")) {
            const char *r_attr = lxr_xml_attr(attrs, "r");
            const char *hidden = lxr_xml_attr(attrs, "hidden");
            ws->row_nr = r_attr ? (size_t)strtoul(r_attr, NULL, 10) : ws->row_nr + 1;
            ws->row_hidden = (hidden && (strcmp(hidden, "1") == 0 ||
                                         strcmp(hidden, "true") == 0));
            ws->row_in_progress = 1;
            ws->state = LXR_WS_IN_ROW;

            if (ws->row_hidden && (ws->flags & LXR_SKIP_HIDDEN_ROWS)) {
                /* skip the entire row content */
                free(ws->skip_tag);
                ws->skip_tag = strdup("row");
                ws->state_before_skip = LXR_WS_IN_SHEETDATA;
                ws->state = LXR_WS_SKIP;
                ws->skip_depth = 1;
                ws->row_in_progress = 0;
                return;
            }

            if (ws->skip_rows_remaining > 0) {
                ws->skip_rows_remaining--;
                free(ws->skip_tag);
                ws->skip_tag = strdup("row");
                ws->state_before_skip = LXR_WS_IN_SHEETDATA;
                ws->state = LXR_WS_SKIP;
                ws->skip_depth = 1;
                ws->row_in_progress = 0;
                return;
            }

            if (ws->pull_mode == LXR_WS_PULL_ROW) {
                ws->pending_row_start = 1;
                lxr_xml_pump_suspend(ws->pump);
            }
        }
        break;

    case LXR_WS_IN_ROW:
        if (lxr_xml_name_eq(name, "c")) {
            const char *r_attr = lxr_xml_attr(attrs, "r");
            const char *t_attr = lxr_xml_attr(attrs, "t");
            const char *s_attr = lxr_xml_attr(attrs, "s");

            reset_cell(ws);
            copy_attr(ws->cell_ref, sizeof(ws->cell_ref), r_attr);
            copy_attr(ws->cell_t,   sizeof(ws->cell_t),   t_attr);
            ws->cell_style_id = s_attr ? (uint32_t)strtoul(s_attr, NULL, 10) : 0;
            if (r_attr) parse_cell_ref(r_attr, &ws->cell_row, &ws->cell_col);
            else { ws->cell_row = ws->row_nr; ws->cell_col = 0; }
            if (ws->cell_col > ws->max_col_seen) ws->max_col_seen = ws->cell_col;

            ws->state = LXR_WS_IN_CELL;
        }
        break;

    case LXR_WS_IN_CELL:
        if (lxr_xml_name_eq(name, "v")) {
            ws->state = LXR_WS_IN_VALUE;
        } else if (lxr_xml_name_eq(name, "f")) {
            ws->cell_has_formula = 1;
            ws->state = LXR_WS_IN_FORMULA;
        } else if (lxr_xml_name_eq(name, "is")) {
            ws->cell_has_inline = 1;
            ws->state = LXR_WS_IN_INLINE_STR;
        } else {
            /* Skip unknown sub-elements (extLst, etc.) */
            free(ws->skip_tag);
            ws->skip_tag = strdup(name);
            ws->state_before_skip = LXR_WS_IN_CELL;
            ws->state = LXR_WS_SKIP;
            ws->skip_depth = 1;
        }
        break;

    case LXR_WS_IN_INLINE_STR:
        if (lxr_xml_name_eq(name, "t")) {
            ws->state = LXR_WS_IN_INLINE_STR_T;
        } else if (lxr_xml_name_eq(name, "r")) {
            /* rich-text run — wait for inner <t> */
        } else {
            free(ws->skip_tag);
            ws->skip_tag = strdup(name);
            ws->state_before_skip = LXR_WS_IN_INLINE_STR;
            ws->state = LXR_WS_SKIP;
            ws->skip_depth = 1;
        }
        break;

    default:
        break;
    }
}

static void on_text(void *ud, const char *text, int len)
{
    lxr_worksheet *ws = (lxr_worksheet *)ud;
    if (len <= 0) return;

    switch (ws->state) {
    case LXR_WS_IN_VALUE:
        append_buf(&ws->cell_value, &ws->cell_value_len, &ws->cell_value_cap,
                   text, (size_t)len);
        break;
    case LXR_WS_IN_FORMULA:
        append_buf(&ws->cell_formula, &ws->cell_formula_len, &ws->cell_formula_cap,
                   text, (size_t)len);
        break;
    case LXR_WS_IN_INLINE_STR_T:
        append_buf(&ws->cell_inline, &ws->cell_inline_len, &ws->cell_inline_cap,
                   text, (size_t)len);
        break;
    default:
        break;
    }
}

static void on_end(void *ud, const char *name)
{
    lxr_worksheet *ws = (lxr_worksheet *)ud;

    if (ws->state == LXR_WS_SKIP) {
        ws->skip_depth--;
        if (ws->skip_depth == 0 && ws->skip_tag &&
            lxr_xml_name_eq(name, ws->skip_tag)) {
            free(ws->skip_tag);
            ws->skip_tag = NULL;
            ws->state = ws->state_before_skip;
            /* For row-level skips, transition out the same way a real
             * </row> would (back to IN_SHEETDATA). */
        }
        return;
    }

    switch (ws->state) {
    case LXR_WS_IN_VALUE:
        if (lxr_xml_name_eq(name, "v")) ws->state = LXR_WS_IN_CELL;
        break;
    case LXR_WS_IN_FORMULA:
        if (lxr_xml_name_eq(name, "f")) ws->state = LXR_WS_IN_CELL;
        break;
    case LXR_WS_IN_INLINE_STR_T:
        if (lxr_xml_name_eq(name, "t")) ws->state = LXR_WS_IN_INLINE_STR;
        break;
    case LXR_WS_IN_INLINE_STR:
        if (lxr_xml_name_eq(name, "is")) ws->state = LXR_WS_IN_CELL;
        break;
    case LXR_WS_IN_CELL:
        if (lxr_xml_name_eq(name, "c")) {
            ws->state = LXR_WS_IN_ROW;
            deliver_cell(ws);
        }
        break;
    case LXR_WS_IN_ROW:
        if (lxr_xml_name_eq(name, "row")) {
            ws->state = LXR_WS_IN_SHEETDATA;
            ws->row_in_progress = 0;
            deliver_row_end(ws);
        }
        break;
    case LXR_WS_IN_SHEETDATA:
        if (lxr_xml_name_eq(name, "sheetData")) ws->state = LXR_WS_IN_WORKSHEET;
        break;
    case LXR_WS_IN_WORKSHEET:
        if (lxr_xml_name_eq(name, "worksheet")) {
            ws->state = LXR_WS_INIT;
            ws->eof   = 1;
        }
        break;
    default:
        break;
    }
}

/* ------------------------------------------------------------------------- */
/* Open / close                                                              */
/* ------------------------------------------------------------------------- */

static lxr_error open_internal(lxr_workbook *wb, const char *target,
                               uint32_t flags, lxr_worksheet **out)
{
    lxr_worksheet *ws;
    lxr_error      rc;
    if (!wb || !target || !out) return LXR_ERROR_NULL_PARAMETER;

    ws = (lxr_worksheet *)calloc(1, sizeof(*ws));
    if (!ws) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    ws->wb          = wb;
    ws->flags       = flags;
    ws->state       = LXR_WS_INIT;
    ws->target_path = strdup(target);
    if (!ws->target_path) { free(ws); return LXR_ERROR_MEMORY_MALLOC_FAILED; }

    /* Eagerly scan worksheet-level metadata before opening the streaming
     * pump. Only one minizip entry can be open at a time, so the meta pass
     * uses its own open/close cycle. */
    rc = lxr_worksheet_meta_load(ws);
    if (rc != LXR_NO_ERROR) {
        lxr_worksheet_meta_free(&ws->meta);
        free(ws->target_path);
        free(ws);
        return rc;
    }

    ws->zf = lxr_zip_open_entry(wb->zip, target);
    if (!ws->zf) {
        lxr_worksheet_meta_free(&ws->meta);
        free(ws->target_path);
        free(ws);
        return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    }

    ws->pump = lxr_xml_pump_create_zip_file(ws->zf);
    if (!ws->pump) {
        lxr_zip_close_entry(ws->zf);
        lxr_worksheet_meta_free(&ws->meta);
        free(ws->target_path);
        free(ws);
        return LXR_ERROR_MEMORY_MALLOC_FAILED;
    }
    lxr_xml_pump_set_handlers(ws->pump, on_start, on_end, on_text, ws);

    *out = ws;
    return LXR_NO_ERROR;
}

lxr_error lxr_worksheet_open_internal(lxr_workbook *wb, const char *target,
                                      uint32_t flags, lxr_worksheet **out)
{
    return open_internal(wb, target, flags, out);
}

void lxr_worksheet_close(lxr_worksheet *ws)
{
    if (!ws) return;
    if (ws->pump) lxr_xml_pump_destroy(ws->pump);
    if (ws->zf)   lxr_zip_close_entry(ws->zf);
    lxr_worksheet_meta_free(&ws->meta);
    free(ws->cell_value);
    free(ws->cell_formula);
    free(ws->cell_inline);
    free(ws->skip_tag);
    free(ws->target_path);
    free(ws);
}

/* ------------------------------------------------------------------------- */
/* Pull mode                                                                 */
/* ------------------------------------------------------------------------- */

static lxr_error drive(lxr_worksheet *ws)
{
    lxr_error rc;
    if (lxr_xml_pump_is_eof(ws->pump) && !lxr_xml_pump_is_suspended(ws->pump)) {
        return LXR_ERROR_END_OF_DATA;
    }
    if (lxr_xml_pump_is_suspended(ws->pump)) {
        rc = lxr_xml_pump_resume(ws->pump);
    } else {
        rc = lxr_xml_pump_run(ws->pump);
    }
    return rc;
}

lxr_error lxr_worksheet_next_row(lxr_worksheet *ws)
{
    if (!ws) return LXR_ERROR_NULL_PARAMETER;

    /* Drain anything left from a previous row by calling next_cell until it
     * returns END_OF_DATA. next_cell uses the suspend mechanism, so the pump
     * stops cleanly at the </row> boundary instead of running into the next
     * row (which would steal it from this consumer). */
    if (ws->row_in_progress) {
        lxr_cell dummy;
        while (lxr_worksheet_next_cell(ws, &dummy) == LXR_NO_ERROR) ;
    }

    ws->pending_row_start = 0;
    ws->pending_row_end   = 0;
    ws->pending_cell      = 0;
    ws->pull_mode         = LXR_WS_PULL_ROW;

    while (!ws->pending_row_start && !ws->eof) {
        lxr_error rc = drive(ws);
        if (rc != LXR_NO_ERROR) {
            ws->pull_mode = LXR_WS_PULL_NONE;
            return rc;
        }
        if (!ws->pending_row_start && lxr_xml_pump_is_eof(ws->pump)) break;
    }

    ws->pull_mode = LXR_WS_PULL_NONE;
    if (!ws->pending_row_start) return LXR_ERROR_END_OF_DATA;

    ws->pending_row_start = 0;
    ws->max_col_seen = 0;
    return LXR_NO_ERROR;
}

lxr_error lxr_worksheet_next_cell(lxr_worksheet *ws, lxr_cell *out)
{
    if (!ws || !out) return LXR_ERROR_NULL_PARAMETER;
    if (!ws->row_in_progress) return LXR_ERROR_END_OF_DATA;

    ws->pending_row_end = 0;
    ws->pending_cell    = 0;
    ws->pull_mode       = LXR_WS_PULL_CELL;

    while (!ws->pending_cell && !ws->pending_row_end && !ws->eof) {
        lxr_error rc = drive(ws);
        if (rc != LXR_NO_ERROR) {
            ws->pull_mode = LXR_WS_PULL_NONE;
            return rc;
        }
        if (lxr_xml_pump_is_eof(ws->pump) && !ws->pending_cell) break;
    }

    ws->pull_mode = LXR_WS_PULL_NONE;

    if (ws->pending_cell) {
        ws->pending_cell = 0;
        emit_cell(ws, out);
        return LXR_NO_ERROR;
    }
    /* row ended */
    return LXR_ERROR_END_OF_DATA;
}

size_t lxr_worksheet_current_row(const lxr_worksheet *ws)
{
    return ws ? ws->row_nr : 0;
}

size_t lxr_worksheet_max_column_seen(const lxr_worksheet *ws)
{
    return ws ? ws->max_col_seen : 0;
}

uint32_t lxr_worksheet_flags(const lxr_worksheet *ws)
{
    return ws ? ws->flags : 0;
}

/* ------------------------------------------------------------------------- */
/* Push (callback) mode                                                      */
/* ------------------------------------------------------------------------- */

lxr_error lxr_worksheet_process(lxr_worksheet *ws,
                                lxr_cell_cb    cell_cb,
                                lxr_row_end_cb row_cb,
                                void          *userdata)
{
    if (!ws) return LXR_ERROR_NULL_PARAMETER;

    ws->pull_mode      = LXR_WS_PULL_NONE;
    ws->user_cell_cb   = cell_cb;
    ws->user_row_cb    = row_cb;
    ws->user_data      = userdata;
    ws->callback_stop  = 0;

    while (!ws->eof && !ws->callback_stop) {
        lxr_error rc = drive(ws);
        if (rc != LXR_NO_ERROR) {
            ws->user_cell_cb = NULL;
            ws->user_row_cb  = NULL;
            return rc;
        }
        if (lxr_xml_pump_is_eof(ws->pump)) break;
    }
    ws->user_cell_cb = NULL;
    ws->user_row_cb  = NULL;
    return LXR_NO_ERROR;
}

/* ------------------------------------------------------------------------- */
/* Skip                                                                      */
/* ------------------------------------------------------------------------- */

lxr_error lxr_worksheet_skip_rows(lxr_worksheet *ws, size_t n)
{
    if (!ws) return LXR_ERROR_NULL_PARAMETER;
    ws->skip_rows_remaining += n;
    return LXR_NO_ERROR;
}
