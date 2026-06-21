/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension                                                  |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2018 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <dev@service.viest.me>                                 |
  +----------------------------------------------------------------------+
*/

#include "xlswriter.h"

/* ------------------------------------------------------------------------- */
/* Auto-size width estimation                                                */
/* ------------------------------------------------------------------------- */

/* Returns 1 when the Unicode codepoint is East-Asian Wide/Fullwidth (counts
 * as 2 display columns), 0 otherwise. Covers the common CJK/Hangul/fullwidth
 * and emoji ranges; not exhaustive but matches Excel's auto-fit closely
 * enough for typical content. */
static int xls_codepoint_is_wide(uint32_t cp)
{
    if (cp < 0x1100) return 0;
    if (cp <= 0x115F) return 1;                       /* Hangul Jamo */
    if (cp == 0x2329 || cp == 0x232A) return 1;
    if (cp >= 0x2E80 && cp <= 0x303E) return 1;       /* CJK radicals */
    if (cp >= 0x3041 && cp <= 0x33FF) return 1;       /* Hiragana/Katakana/CJK */
    if (cp >= 0x3400 && cp <= 0x4DBF) return 1;       /* CJK Ext A */
    if (cp >= 0x4E00 && cp <= 0x9FFF) return 1;       /* CJK Unified */
    if (cp >= 0xA000 && cp <= 0xA4CF) return 1;       /* Yi */
    if (cp >= 0xAC00 && cp <= 0xD7A3) return 1;       /* Hangul Syllables */
    if (cp >= 0xF900 && cp <= 0xFAFF) return 1;       /* CJK Compat */
    if (cp >= 0xFE10 && cp <= 0xFE19) return 1;       /* Vertical forms */
    if (cp >= 0xFE30 && cp <= 0xFE6F) return 1;       /* CJK Compat Forms */
    if (cp >= 0xFF00 && cp <= 0xFF60) return 1;       /* Fullwidth Forms */
    if (cp >= 0xFFE0 && cp <= 0xFFE6) return 1;       /* Fullwidth signs */
    if (cp >= 0x1F300 && cp <= 0x1FAFF) return 1;     /* Emoji / symbols */
    if (cp >= 0x20000 && cp <= 0x3FFFD) return 1;     /* CJK Ext B+ */
    return 0;
}

/* Decode one UTF-8 sequence from str[*pos]; advance *pos and store the
 * codepoint. Returns 0 on success, -1 on invalid/truncated input. */
static int utf8_next(const unsigned char *str, size_t len, size_t *pos, uint32_t *cp)
{
    size_t i = *pos;
    unsigned char c = str[i];
    if (c < 0x80) { *cp = c; *pos = i + 1; return 0; }
    if ((c & 0xE0) == 0xC0 && i + 1 < len) {
        *cp = ((uint32_t)(c & 0x1F) << 6) | (str[i + 1] & 0x3F);
        *pos = i + 2; return 0;
    }
    if ((c & 0xF0) == 0xE0 && i + 2 < len) {
        *cp = ((uint32_t)(c & 0x0F) << 12) | ((uint32_t)(str[i + 1] & 0x3F) << 6) | (str[i + 2] & 0x3F);
        *pos = i + 3; return 0;
    }
    if ((c & 0xF8) == 0xF0 && i + 3 < len) {
        *cp = ((uint32_t)(c & 0x07) << 18) | ((uint32_t)(str[i + 1] & 0x3F) << 12)
            | ((uint32_t)(str[i + 2] & 0x3F) << 6) | (str[i + 3] & 0x3F);
        *pos = i + 4; return 0;
    }
    /* Invalid/overlong/truncated: treat as a single narrow byte. */
    *cp = c; *pos = i + 1; return -1;
}

/* Sum display columns over a UTF-8 string (wide codepoints count as 2). */
static double utf8_display_width(const char *s, size_t len)
{
    const unsigned char *str = (const unsigned char *)s;
    size_t pos = 0;
    double width = 0;
    while (pos < len) {
        uint32_t cp;
        utf8_next(str, len, &pos, &cp);
        width += xls_codepoint_is_wide(cp) ? 2 : 1;
    }
    return width;
}

/* Estimate a cell's display width in Excel column-width units. Converts the
 * value to its string form and measures it (wide/CJK codepoints count as 2).
 * Clamped to Excel's maximum column width (255). No extra margin is added:
 * libxlsxwriter already bakes in Excel's standard column margin (~0.71) when
 * the width is written, which matches what Excel's own auto-fit produces. */
#define LXLSX_MAX_COL_WIDTH 255.0

double xls_estimate_cell_width(zval *value)
{
    zend_string *s;
    double width;

    if (value == NULL || Z_TYPE_P(value) == IS_NULL) return 0.0;

    s = zval_get_string(value);
    width = utf8_display_width(ZSTR_VAL(s), ZSTR_LEN(s));
    zend_string_release(s);

    if (width > LXLSX_MAX_COL_WIDTH) width = LXLSX_MAX_COL_WIDTH;
    return width;
}

/* Grow the per-column width map to cover `col` and keep the max. */
void xls_track_auto_width(xls_resource_write_t *res, lxlsx_col_t col, double width)
{
    size_t need, i;
    double *grown;

    if (res == NULL || col >= LXLSX_COL_MAX || width <= 0.0) return;

    if (res->auto_widths == NULL) {
        need = (size_t)col + 1;
        if (need < 16) need = 16;
        res->auto_widths = (double *)ecalloc(need, sizeof(double));
        res->auto_widths_n = need;
    } else if ((size_t)col >= res->auto_widths_n) {
        need = res->auto_widths_n;
        while (need <= (size_t)col) need *= 2;
        grown = (double *)erealloc(res->auto_widths, need * sizeof(double));
        for (i = res->auto_widths_n; i < need; i++) grown[i] = 0.0;
        res->auto_widths = grown;
        res->auto_widths_n = need;
    }

    if (width > res->auto_widths[col]) res->auto_widths[col] = width;
}

void xls_auto_widths_reset(xls_resource_write_t *res)
{
    if (res == NULL || res->auto_widths == NULL) return;
    efree(res->auto_widths);
    res->auto_widths = NULL;
    res->auto_widths_n = 0;
}

/* Apply the tracked per-column widths to [first_col, last_col] on the active
 * worksheet. Columns with no tracked content keep their existing width. A
 * set_column failure (e.g. optimize-mode index conflict) is surfaced as a
 * warning rather than silently dropped. */
void xls_auto_widths_apply(xls_resource_write_t *res, lxlsx_col_t first_col, lxlsx_col_t last_col)
{
    lxlsx_col_t c;
    if (res == NULL || res->auto_widths == NULL || res->worksheet == NULL) return;
    if (first_col > last_col) { lxlsx_col_t t = first_col; first_col = last_col; last_col = t; }
    for (c = first_col; c <= last_col; c++) {
        if (c >= res->auto_widths_n) break;
        if (res->auto_widths[c] > 0.0) {
            lxlsx_error err = lxlsx_worksheet_set_column_opt(res->worksheet, c, c,
                                                     res->auto_widths[c], NULL, NULL);
            if (err != LXLSX_NO_ERROR) {
                php_error_docref(NULL, E_WARNING,
                    "autoSize: could not set column %u width (%s)",
                    (unsigned)(c + 1), lxlsx_strerror(err));
            }
        }
    }
}

/* Apply the tracked widths for the configured range to the active worksheet,
 * then clear the per-column map (the enable flag survives so the next sheet
 * keeps tracking). Called at output() and on sheet switch. */
void xls_auto_widths_flush(xls_resource_write_t *res)
{
    if (res == NULL || !res->auto_size_enabled) return;
    xls_auto_widths_apply(res, res->auto_size_first_col, res->auto_size_last_col);
    xls_auto_widths_reset(res);
}

/*
 * According to the zval type written to the file
 */
void type_writer(zval *value, zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *format, lxlsx_format *lxlsx_format_handle)
{
    lxlsx_col_t lxlsx_col = (lxlsx_col_t)columns;
    lxlsx_row_t lxlsx_row = (lxlsx_row_t)row;

    zend_uchar value_type = Z_TYPE_P(value);

    /* Track the cell's display width for autoSize() — only when the user
     * has opted in, so writes pay no cost otherwise. */
    if (res->auto_size_enabled) {
        xls_track_auto_width(res, lxlsx_col, xls_estimate_cell_width(value));
    }

    if (value_type == IS_STRING) {
        zend_string *_zs_value = zval_get_string(value);

        int error = lxlsx_worksheet_write_string(res->worksheet, lxlsx_row, lxlsx_col, ZSTR_VAL(_zs_value), lxlsx_format_handle);

        zend_string_release(_zs_value);
        WORKSHEET_WRITER_EXCEPTION(error);
        return;
    }

    if (value_type == IS_LONG) {
        if (format != NULL && lxlsx_format_handle == NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, (double)zval_get_long(value), lxlsx_format_handle));
            return;
        }

        if (format == NULL && lxlsx_format_handle != NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, (double)zval_get_long(value), lxlsx_format_handle));
            return;
        }

        if(format != NULL && lxlsx_format_handle != NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, (double)zval_get_long(value), lxlsx_format_handle));
            return;
        }

        WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, (double)zval_get_long(value), NULL));
    }

    if (value_type == IS_DOUBLE) {
        if (format != NULL && lxlsx_format_handle == NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, zval_get_double(value), lxlsx_format_handle));
            return;
        }

        if (format == NULL && lxlsx_format_handle != NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, zval_get_double(value), lxlsx_format_handle));
            return;
        }

        if(format != NULL && lxlsx_format_handle != NULL) {
            WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, lxlsx_row, lxlsx_col, zval_get_double(value), lxlsx_format_handle));
            return;
        }

        WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_number(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, zval_get_double(value), NULL));
        return;
    }

    if (value_type == IS_TRUE || value_type == IS_FALSE) {
        WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_boolean(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, zend_is_true(value), lxlsx_format_handle));
        return;
    }
}

/*
 * Write the rich string to the file
 */
void rich_string_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zval *rich_strings, lxlsx_format *format)
{
    int index = 0, resource_count = 0;
    zval *zv_rich_string = NULL;

    lxlsx_col_t lxlsx_col = (lxlsx_col_t)columns;
    lxlsx_row_t lxlsx_row = (lxlsx_row_t)row;

    if (Z_TYPE_P(rich_strings) != IS_ARRAY) {
        return;
    }

    ZEND_HASH_FOREACH_VAL(Z_ARRVAL_P(rich_strings), zv_rich_string)
        if (Z_TYPE_P(zv_rich_string) != IS_OBJECT) {
            continue;
        }

        if (!instanceof_function(Z_OBJCE_P(zv_rich_string), vtiful_rich_string_ce)) {
            zend_throw_exception(vtiful_exception_ce, "The parameter must be an instance of Vtiful\\Kernel\\RichString.", 500);
            return;
        }

        resource_count++;
    ZEND_HASH_FOREACH_END();

    lxlsx_rich_string_tuple **rich_string_list = (lxlsx_rich_string_tuple **)ecalloc(resource_count + 1,sizeof(lxlsx_rich_string_tuple *));

    ZEND_HASH_FOREACH_VAL(Z_ARRVAL_P(rich_strings), zv_rich_string)
        rich_string_object *obj = Z_RICH_STR_P(zv_rich_string);
        rich_string_list[index] = obj->ptr.tuple;
        index++;
    ZEND_HASH_FOREACH_END();

    rich_string_list[index] = NULL;

    WORKSHEET_WRITER_EXCEPTION(lxlsx_worksheet_write_rich_string(res->worksheet, lxlsx_row, lxlsx_col, rich_string_list, format));

    efree(rich_string_list);
}

void lxlsx_format_copy(lxlsx_format *new_format, lxlsx_format *other_format)
{
    /* Font-family string: previously skipped, causing every clone to fall
     * back to the workbook default (Calibri). Reported as #545 / #472:
     * insertText() with both a num-format string and a format resource
     * dropped the caller's font(). num_format/font_scheme/has_font flags
     * are in the same boat — copy them all. */
    memcpy(new_format->font_name,   other_format->font_name,   LXLSX_FORMAT_FIELD_LEN);
    memcpy(new_format->font_scheme, other_format->font_scheme, LXLSX_FORMAT_FIELD_LEN);
    new_format->has_font     = other_format->has_font;
    new_format->has_dxf_font = other_format->has_dxf_font;

    new_format->bold = other_format->bold;
    new_format->bg_color = other_format->bg_color;
    new_format->border_count = other_format->border_count;
    new_format->border_index = other_format->border_index;
    new_format->bottom = other_format->bottom;
    new_format->bottom_color = other_format->bottom_color;
    new_format->color_indexed = other_format->color_indexed;
    new_format->diag_border = other_format->diag_border;
    new_format->diag_color = other_format->diag_color;

    new_format->font_size = other_format->font_size;
    new_format->bold = other_format->bold;
    new_format->italic = other_format->italic;
    new_format->font_color = other_format->font_color;
    new_format->underline = other_format->underline;
    new_format->font_strikeout = other_format->font_strikeout;
    new_format->font_outline = other_format->font_outline;
    new_format->font_shadow = other_format->font_shadow;
    new_format->font_script = other_format->font_script;
    new_format->font_family = other_format->font_family;
    new_format->font_charset = other_format->font_charset;
    new_format->font_condense = other_format->font_condense;
    new_format->font_extend = other_format->font_extend;
    new_format->theme = other_format->theme;
    new_format->hyperlink = other_format->hyperlink;

    new_format->hidden = other_format->hidden;
    new_format->locked = other_format->locked;

    new_format->text_h_align = other_format->text_h_align;
    new_format->text_wrap = other_format->text_wrap;
    new_format->text_v_align = other_format->text_v_align;
    new_format->text_justlast = other_format->text_justlast;
    new_format->rotation = other_format->rotation;

    new_format->fg_color = other_format->fg_color;
    new_format->bg_color = other_format->bg_color;
    new_format->pattern = other_format->pattern;
    new_format->has_fill = other_format->has_fill;
    new_format->has_dxf_fill = other_format->has_dxf_fill;
    new_format->fill_index = other_format->fill_index;
    new_format->fill_count = other_format->fill_count;

    new_format->border_index = other_format->border_index;
    new_format->has_border = other_format->has_border;
    new_format->has_dxf_border = other_format->has_dxf_border;
    new_format->border_count = other_format->border_count;

    new_format->bottom = other_format->bottom;
    new_format->diag_border = other_format->diag_border;
    new_format->diag_type = other_format->diag_type;
    new_format->left = other_format->left;
    new_format->right = other_format->right;
    new_format->top = other_format->top;
    new_format->bottom_color = other_format->bottom_color;
    new_format->diag_color = other_format->diag_color;
    new_format->left_color = other_format->left_color;
    new_format->right_color = other_format->right_color;
    new_format->top_color = other_format->top_color;

    new_format->indent = other_format->indent;
    new_format->shrink = other_format->shrink;
    new_format->merge_range = other_format->merge_range;
    new_format->reading_order = other_format->reading_order;
    new_format->just_distrib = other_format->just_distrib;
    new_format->color_indexed = other_format->color_indexed;
    new_format->font_only = other_format->font_only;
}

void url_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *url, zend_string *text, zend_string *tool_tip, lxlsx_format *format)
{
    int error = lxlsx_worksheet_write_url_opt(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, ZSTR_VAL(url), format,
                                              text != NULL ? ZSTR_VAL(text) : NULL,
                                              tool_tip != NULL ? ZSTR_VAL(tool_tip) : NULL);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Write the image to the file
 */
void image_writer(zval *value, zend_long row, zend_long columns, double width, double height, xls_resource_write_t *res)
{
    lxlsx_image_options options = {.x_offset = 0, .y_offset = 0, .x_scale = width, .y_scale = height, .object_position = 2};
    zend_string *path = zval_get_string(value);
    int error = lxlsx_worksheet_insert_image_opt(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, ZSTR_VAL(path), &options);
    zend_string_release(path);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Write the image with full options struct (insertImageOpt).
 */
void image_opt_writer(zval *value, zend_long row, zend_long columns,
                     lxlsx_image_options *options, xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_insert_image_opt(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns,
                                           ZSTR_VAL(zval_get_string(value)), options);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Write the image to the file
 */
void formula_writer(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format)
{
    WORKSHEET_WRITER_EXCEPTION(
        lxlsx_worksheet_write_formula(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, ZSTR_VAL(value), format));
}

/*
 * Like formula_writer but evaluates the formula against the cells written so
 * far and stores the computed value as the cached result (compute-on-write).
 * References to not-yet-written cells (or already-flushed cells in
 * constant-memory mode) resolve as blank, matching evaluateFormula().
 */
void formula_writer_calc(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format)
{
    const char *f = ZSTR_VAL(value);
    lxlsx_value out;
    lxlsx_error err;

    if (lxlsx_formula_eval(f, formula_resolver, res->worksheet, &out) != LXLSX_NO_ERROR) {
        /* engine-level failure: fall back to a plain formula (cached 0). */
        WORKSHEET_WRITER_EXCEPTION(
            lxlsx_worksheet_write_formula(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, f, format));
        return;
    }

    switch (out.kind) {
    case LXLSX_VAL_NUMBER:
    case LXLSX_VAL_BOOL:
        err = lxlsx_worksheet_write_formula_num(res->worksheet, (lxlsx_row_t)row,
                                                (lxlsx_col_t)columns, f, format, out.number);
        break;
    case LXLSX_VAL_STRING:
        err = lxlsx_worksheet_write_formula_str(res->worksheet, (lxlsx_row_t)row,
                                                (lxlsx_col_t)columns, f, format,
                                                out.string ? out.string : "");
        break;
    case LXLSX_VAL_ERROR:
        err = lxlsx_worksheet_write_formula_str(res->worksheet, (lxlsx_row_t)row,
                                                (lxlsx_col_t)columns, f, format,
                                                lxlsx_formula_error_string(out.error));
        break;
    case LXLSX_VAL_BLANK:
    default:
        err = lxlsx_worksheet_write_formula_num(res->worksheet, (lxlsx_row_t)row,
                                                (lxlsx_col_t)columns, f, format, 0.0);
        break;
    }

    lxlsx_value_free(&out);
    WORKSHEET_WRITER_EXCEPTION(err);
}

void dynamic_formula_writer(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format)
{
    WORKSHEET_WRITER_EXCEPTION(
        lxlsx_worksheet_write_dynamic_formula(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, ZSTR_VAL(value), format));
}

void dynamic_array_formula_writer(zend_string *value, zend_long first_row, zend_long first_col,
                                  zend_long last_row, zend_long last_col,
                                  xls_resource_write_t *res, lxlsx_format *format)
{
    WORKSHEET_WRITER_EXCEPTION(
        lxlsx_worksheet_write_dynamic_array_formula(res->worksheet, (lxlsx_row_t)first_row, (lxlsx_col_t)first_col,
                                              (lxlsx_row_t)last_row, (lxlsx_col_t)last_col,
                                              ZSTR_VAL(value), format));
}

/*
 * Write the chart to the file
 */
void lxlsx_chart_writer(zend_long row, zend_long columns, xls_resource_chart_t *lxlsx_chart_resource, xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_insert_chart(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, lxlsx_chart_resource->chart);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Write the datetime to the file
 */
void datetime_writer(lxlsx_datetime *datetime, zend_long row, zend_long columns, zend_string *format, xls_resource_write_t *res, lxlsx_format *lxlsx_format_handle)
{
    int error = lxlsx_worksheet_write_datetime(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, datetime, lxlsx_format_handle);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Write the comment to the cell
 */
void comment_writer(zend_string *comment, zend_long row, zend_long columns, xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_write_comment(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns, ZSTR_VAL(comment));

    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Show all comments
 */
void comment_show(xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_show_comments(res->worksheet);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Add the autofilter.
 */
void auto_filter(zend_string *range, xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_autofilter(res->worksheet, RANGE(ZSTR_VAL(range)));

    // Cells that have been placed cannot be modified using optimization mode
    WORKSHEET_INDEX_OUT_OF_CHANGE_IN_OPTIMIZE_EXCEPTION(res, error)

    // Worksheet row or column index out of range
    WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)

    // Any other error, e.g. unsupported in edit mode
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Merge cells.
 */
void merge_cells(zend_string *range, zval *value, xls_resource_write_t *res, lxlsx_format *format)
{
    char *_range = ZSTR_VAL(range);

    int error = lxlsx_worksheet_merge_range(res->worksheet, RANGE(_range), "", format);

    // Cells that have been placed cannot be modified using optimization mode
    WORKSHEET_INDEX_OUT_OF_CHANGE_IN_OPTIMIZE_EXCEPTION(res, error)

    // Worksheet row or column index out of range
    WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)

    // Any other error, e.g. unsupported in edit mode
    WORKSHEET_WRITER_EXCEPTION(error);

    // writer merge cell
    type_writer(value, lxlsx_name_to_row(_range), lxlsx_name_to_col(_range), res, NULL, format);
}

/*
 * Set column format
 */
void set_column(zend_string *range, double width, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options)
{
    int error = lxlsx_worksheet_set_column_opt(res->worksheet, COLS(ZSTR_VAL(range)), width, format, user_options);

    // Cells that have been placed cannot be modified using optimization mode
    WORKSHEET_INDEX_OUT_OF_CHANGE_IN_OPTIMIZE_EXCEPTION(res, error)

    // Worksheet row or column index out of range
    WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)

    // Any other error, e.g. unsupported in edit mode
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Set row format
 */
void set_row(zend_string *range, double height, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options)
{
    char *rows = ZSTR_VAL(range);

    if (strchr(rows, ':')) {
        int error = lxlsx_worksheet_set_rows(ROWS(rows), height, res, format, user_options);

        WORKSHEET_WRITER_EXCEPTION(error);
    } else {
        int error = lxlsx_worksheet_set_row_opt(res->worksheet, ROW(rows), height, format, user_options);

        // Cells that have been placed cannot be modified using optimization mode
        WORKSHEET_INDEX_OUT_OF_CHANGE_IN_OPTIMIZE_EXCEPTION(res, error)

        // Worksheet row or column index out of range
        WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)

        // Any other error, e.g. unsupported in edit mode
        WORKSHEET_WRITER_EXCEPTION(error);
    }
}

/*
 * Add data validations to a worksheet
 */
void validation(xls_resource_write_t *res, zend_string *range, lxlsx_data_validation *validation)
{
    char *rangeStr = ZSTR_VAL(range);
    int error;

    if (strchr(rangeStr, ':')) {
	    error = lxlsx_worksheet_data_validation_range(res->worksheet, RANGE(rangeStr), validation);
    } else {
	    error = lxlsx_worksheet_data_validation_cell(res->worksheet, CELL(rangeStr), validation);
    }

    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Set rows format
 */
lxlsx_error lxlsx_worksheet_set_rows(lxlsx_row_t start, lxlsx_row_t end, double height, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options)
{
    lxlsx_error error = LXLSX_NO_ERROR;

    while (1) {
        error = lxlsx_worksheet_set_row_opt(res->worksheet, end, height, format, user_options);
        if (error > LXLSX_NO_ERROR)
            return error;
        if (end == start)
            break;
        end--;
    }

    return LXLSX_NO_ERROR;
}

/*
 * Set freeze panes
 */
void freeze_panes(xls_resource_write_t *res, zend_long row, zend_long column)
{
    int error = lxlsx_worksheet_freeze_panes(res->worksheet, row, column);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Display or hide screen and print gridlines
 */
void gridlines(xls_resource_write_t *res, zend_long option)
{
    lxlsx_worksheet_gridlines(res->worksheet, option);
}

/*
 * Set the worksheet zoom factor
 */
void zoom(xls_resource_write_t *res, zend_long zoom)
{
    lxlsx_worksheet_set_zoom(res->worksheet, zoom);
}

/*
 * Set the worksheet protection
 */
void protection(xls_resource_write_t *res, zend_string *password)
{
    if (password == NULL) {
        lxlsx_worksheet_protect(res->worksheet, NULL, NULL);
    } else {
        lxlsx_worksheet_protect(res->worksheet, ZSTR_VAL(password), NULL);
    }
}

/*
 * Set the worksheet printed direction
 */
void printed_direction(xls_resource_write_t *res, unsigned int direction)
{
    if (direction == XLSWRITER_PRINTED_PORTRAIT) {
        lxlsx_worksheet_set_portrait(res->worksheet);
    }

    lxlsx_worksheet_set_landscape(res->worksheet);
}

/*
 * Set the worksheet printed scale
 */
void printed_scale(xls_resource_write_t *res, zend_long scale)
{
    if (scale < 10) {
        scale = 10;
    }

    if (scale > 400) {
        scale = 400;
    }

    lxlsx_worksheet_set_print_scale(res->worksheet, scale);
}

/*
 * Hide worksheet
 */
void hide_worksheet(xls_resource_write_t *res)
{
    lxlsx_worksheet_hide(res->worksheet);
}

/*
 * First worksheet
 */
void first_worksheet(xls_resource_write_t *res)
{
    lxlsx_worksheet_set_first_sheet(res->worksheet);
}

/*
 * Paper format
 */
void paper(xls_resource_write_t *res, zend_long type)
{
    lxlsx_worksheet_set_paper(res->worksheet, type);
}

/*
 * Set margins
 */
void margins(xls_resource_write_t *res, double left, double right, double top, double bottom)
{
    lxlsx_worksheet_set_margins(res->worksheet, left, right, top, bottom);
}

/* ------------------------------------------------------------------------ */
/* Phase 2 writer helpers                                                    */
/* ------------------------------------------------------------------------ */

/* Comment with full options. Caller fills out the lxlsx_comment_options
 * struct (zero-init for defaults). */
void comment_opt_writer(zend_string *comment, zend_long row, zend_long columns,
                        lxlsx_comment_options *options, xls_resource_write_t *res)
{
    int error = lxlsx_worksheet_write_comment_opt(res->worksheet,
                                            (lxlsx_row_t)row, (lxlsx_col_t)columns,
                                            ZSTR_VAL(comment), options);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Image from in-memory buffer. */
void image_buffer_writer(zend_long row, zend_long columns,
                         const unsigned char *bytes, size_t size,
                         lxlsx_image_options *options,
                         xls_resource_write_t *res)
{
    int error;
    if (options) {
        error = lxlsx_worksheet_insert_image_buffer_opt(res->worksheet,
                                                  (lxlsx_row_t)row, (lxlsx_col_t)columns,
                                                  bytes, size, options);
    } else {
        error = lxlsx_worksheet_insert_image_buffer(res->worksheet,
                                              (lxlsx_row_t)row, (lxlsx_col_t)columns,
                                              bytes, size);
    }
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Header / footer (with optional image filenames). */
void header_writer(xls_resource_write_t *res, const char *value,
                   lxlsx_header_footer_options *options)
{
    int error = options
        ? lxlsx_worksheet_set_header_opt(res->worksheet, value, options)
        : lxlsx_worksheet_set_header(res->worksheet, value);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void footer_writer(xls_resource_write_t *res, const char *value,
                   lxlsx_header_footer_options *options)
{
    int error = options
        ? lxlsx_worksheet_set_footer_opt(res->worksheet, value, options)
        : lxlsx_worksheet_set_footer(res->worksheet, value);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Print: repeat rows / columns / area. Inputs are 0-based. */
void repeat_rows_writer(xls_resource_write_t *res,
                        zend_long first_row, zend_long last_row)
{
    int error = lxlsx_worksheet_repeat_rows(res->worksheet,
                                      (lxlsx_row_t)first_row, (lxlsx_row_t)last_row);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void repeat_columns_writer(xls_resource_write_t *res,
                           zend_long first_col, zend_long last_col)
{
    int error = lxlsx_worksheet_repeat_columns(res->worksheet,
                                         (lxlsx_col_t)first_col, (lxlsx_col_t)last_col);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void print_area_writer(xls_resource_write_t *res,
                       zend_long first_row, zend_long first_col,
                       zend_long last_row,  zend_long last_col)
{
    int error = lxlsx_worksheet_print_area(res->worksheet,
                                     (lxlsx_row_t)first_row, (lxlsx_col_t)first_col,
                                     (lxlsx_row_t)last_row,  (lxlsx_col_t)last_col);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Page breaks. The libxlsxwriter API expects a 0-terminated array. */
void h_pagebreaks_writer(xls_resource_write_t *res, lxlsx_row_t *breaks)
{
    int error = lxlsx_worksheet_set_h_pagebreaks(res->worksheet, breaks);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void v_pagebreaks_writer(xls_resource_write_t *res, lxlsx_col_t *breaks)
{
    int error = lxlsx_worksheet_set_v_pagebreaks(res->worksheet, breaks);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void fit_to_pages_writer(xls_resource_write_t *res,
                         zend_long width, zend_long height)
{
    lxlsx_worksheet_fit_to_pages(res->worksheet, (uint16_t)width, (uint16_t)height);
}

/* Sheet tab color. */
void tab_color_writer(xls_resource_write_t *res, zend_long rgb)
{
    lxlsx_worksheet_set_tab_color(res->worksheet, (lxlsx_color_t)rgb);
}

/* Background image (file or buffer). */
void background_image_writer(xls_resource_write_t *res, const char *path)
{
    int error = lxlsx_worksheet_set_background(res->worksheet, path);
    WORKSHEET_WRITER_EXCEPTION(error);
}

void background_image_buffer_writer(xls_resource_write_t *res,
                                    const unsigned char *bytes, size_t size)
{
    int error = lxlsx_worksheet_set_background_buffer(res->worksheet, bytes, size);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Document properties (workbook-level). */
void lxlsx_workbook_properties_writer(xls_resource_write_t *res,
                                lxlsx_doc_properties *props)
{
    int error = lxlsx_workbook_set_properties(res->workbook, props);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Defined names (workbook-level or sheet-scoped via "Sheet!Name"). */
void define_name_writer(xls_resource_write_t *res,
                        const char *name, const char *formula)
{
    int error = lxlsx_workbook_define_name(res->workbook, name, formula);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Conditional format (cell or range). */
void conditional_format_writer(xls_resource_write_t *res,
                               zend_long first_row, zend_long first_col,
                               zend_long last_row,  zend_long last_col,
                               lxlsx_conditional_format *cf)
{
    int error;
    if (first_row == last_row && first_col == last_col) {
        error = lxlsx_worksheet_conditional_format_cell(res->worksheet,
                                                  (lxlsx_row_t)first_row,
                                                  (lxlsx_col_t)first_col, cf);
    } else {
        error = lxlsx_worksheet_conditional_format_range(res->worksheet,
                                                   (lxlsx_row_t)first_row,
                                                   (lxlsx_col_t)first_col,
                                                   (lxlsx_row_t)last_row,
                                                   (lxlsx_col_t)last_col, cf);
    }
    WORKSHEET_WRITER_EXCEPTION(error);
}

/* Excel Table. */
void add_table_writer(xls_resource_write_t *res,
                      zend_long first_row, zend_long first_col,
                      zend_long last_row,  zend_long last_col,
                      lxlsx_table_options *opts)
{
    int error = lxlsx_worksheet_add_table(res->worksheet,
                                    (lxlsx_row_t)first_row, (lxlsx_col_t)first_col,
                                    (lxlsx_row_t)last_row,  (lxlsx_col_t)last_col,
                                    opts);
    WORKSHEET_WRITER_EXCEPTION(error);
}

/*
 * Set outline settings
 */
void outline_settings(xls_resource_write_t *res, uint8_t visible, uint8_t symbols_below, uint8_t symbols_right, uint8_t auto_style)
{
    lxlsx_worksheet_outline_settings(res->worksheet, visible, symbols_below, symbols_right, auto_style);
}

/*
 * Call finalization code and close file.
 */
lxlsx_error
lxlsx_workbook_file(xls_resource_write_t *self)
{
    lxlsx_sheet *sheet = NULL;
    lxlsx_worksheet *worksheet = NULL;
    lxlsx_packager *packager = NULL;
    lxlsx_error error = LXLSX_NO_ERROR;
    char codename[LXLSX_MAX_SHEETNAME_LENGTH] = { 0 };

    if (lxlsx_workbook_is_edit(self->workbook)) {
        return lxlsx_workbook_save_as(self->workbook, self->workbook->filename);
    }

    /* Add a default worksheet if non have been added. */
    if (!self->workbook->num_sheets)
        lxlsx_workbook_add_worksheet(self->workbook, NULL);

    /* Ensure that at least one worksheet has been selected. */
    if (self->workbook->active_sheet == 0) {
        sheet = STAILQ_FIRST(self->workbook->sheets);
        if (!sheet->is_chartsheet) {
            worksheet = sheet->u.worksheet;
            worksheet->selected = 1;
            worksheet->hidden = 0;
        }
    }

    /* Set the active sheet and check if a metadata file is needed. */
    STAILQ_FOREACH(sheet, self->workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (worksheet->index == self->workbook->active_sheet)
            worksheet->active = 1;

        if (worksheet->has_dynamic_functions) {
            self->workbook->has_metadata = LXLSX_TRUE;
            self->workbook->has_dynamic_functions = 1;
        }

        if (!STAILQ_EMPTY(worksheet->embedded_image_props)) {
            self->workbook->has_metadata = LXLSX_TRUE;
            self->workbook->has_embedded_images = 1;
        }
    }

    /* Set workbook and worksheet VBA codenames if a macro has been added. */
    if (self->workbook->vba_project) {
        if (!self->workbook->vba_codename)
            lxlsx_workbook_set_vba_name(self->workbook, "ThisWorkbook");

        STAILQ_FOREACH(sheet, self->workbook->sheets, list_pointers) {
            if (sheet->is_chartsheet)
                continue;
            else
                worksheet = sheet->u.worksheet;

            if (!worksheet->vba_codename) {
                lxlsx_snprintf(codename, LXLSX_MAX_SHEETNAME_LENGTH, "Sheet%d",
                             worksheet->index + 1);

                lxlsx_worksheet_set_vba_name(worksheet, codename);
            }
        }
    }

    /* Prepare the worksheet VML elements such as comments. */
    _prepare_vml(self->workbook);

    /* Set the defined names for the worksheets such as Print Titles. */
    _prepare_defined_names(self->workbook);

    /* Prepare the drawings, charts and images. */
    _prepare_drawings(self->workbook);

    /* Add cached data to charts. */
    _add_chart_cache_data(self->workbook);

    /* Set the table ids for the worksheet tables. */
    _prepare_tables(self->workbook);

    /* Create a packager object to assemble sub-elements into a zip file. */
    packager = lxlsx_packager_new(self->workbook->filename,
                                self->workbook->options.tmpdir,
                                self->workbook->options.use_zip64);

    /* If the packager fails it is generally due to a zip permission error. */
    if (packager == NULL) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Error creating '%s'. "
                        "Error = %s\n", self->workbook->filename, strerror(errno));

        error = LXLSX_ERROR_CREATING_XLSX_FILE;
        goto mem_error;
    }

    /* Set the workbook object in the packager. */
    packager->workbook = self->workbook;

    /* Assemble all the sub-files in the xlsx package. */
    error = lxlsx_create_package(packager);

    /* Error and non-error conditions fall through to the cleanup code. */
    if (error == LXLSX_ERROR_CREATING_TMPFILE) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Error creating tmpfile(s) to assemble '%s'. "
                        "Error = %s\n", self->workbook->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_FILE_OPERATION then errno is set by zlib. */
    if (error == LXLSX_ERROR_ZIP_FILE_OPERATION) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zlib error while creating xlsx file '%s'. "
                        "Error = %s\n", self->workbook->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_PARAMETER_ERROR then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_PARAMETER_ERROR) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zip ZIP_PARAMERROR error while creating xlsx file '%s'. "
                        "System error = %s\n", self->workbook->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_BAD_ZIP_FILE then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_BAD_ZIP_FILE) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zip ZIP_BADZIPFILE error while creating xlsx file '%s'. "
                        "This may require the use_zip64 option for large files. "
                        "System error = %s\n", self->workbook->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_INTERNAL_ERROR then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_INTERNAL_ERROR) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zip ZIP_INTERNALERROR error while creating xlsx file '%s'. "
                        "System error = %s\n", self->workbook->filename, strerror(errno));
    }

    /* The next 2 error conditions don't set errno. */
    if (error == LXLSX_ERROR_ZIP_FILE_ADD) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zlib error adding file to xlsx file '%s'.\n",
                self->workbook->filename);
    }

    if (error == LXLSX_ERROR_ZIP_CLOSE) {
        fprintf(stderr, "[ERROR] lxlsx_workbook_close(): "
                        "Zlib error closing xlsx file '%s'.\n", self->workbook->filename);
    }

    mem_error:
    lxlsx_packager_free(packager);

    return error;
}

void _php_vtiful_xls_close(zend_resource *rsrc TSRMLS_DC)
{
    //
}

/*
 * Iterate through the worksheets and set up the VML objects.
 */

STATIC void
_prepare_vml(lxlsx_workbook *self)
{
    lxlsx_worksheet *worksheet;
    lxlsx_sheet *sheet;
    uint32_t comment_id = 0;
    uint32_t vml_drawing_id = 0;
    uint32_t vml_data_id = 1;
    uint32_t vml_header_id = 0;
    uint32_t vml_shape_id = 1024;
    uint32_t comment_count = 0;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (!worksheet->has_vml && !worksheet->has_header_vml)
            continue;

        if (worksheet->has_vml) {
            self->has_vml = LXLSX_TRUE;
            if (worksheet->has_comments) {
                self->comment_count += 1;
                comment_id += 1;
                self->has_comments = LXLSX_TRUE;
            }

            vml_drawing_id += 1;

            comment_count = lxlsx_worksheet_prepare_vml_objects(worksheet,
                                                              vml_data_id,
                                                              vml_shape_id,
                                                              vml_drawing_id,
                                                              comment_id);

            /* Each VML should start with a shape id incremented by 1024. */
            vml_data_id += 1 * ((1024 + comment_count) / 1024);
            vml_shape_id += 1024 * ((1024 + comment_count) / 1024);
        }

        /* Header/footer image VML — required for setHeader([image_*=>...])
         * to produce a valid xlsx. Mirrors upstream _prepare_vml(). */
        if (worksheet->has_header_vml) {
            self->has_vml = LXLSX_TRUE;
            vml_drawing_id += 1;
            vml_header_id += 1;
            lxlsx_worksheet_prepare_header_vml_objects(worksheet,
                                                     vml_header_id,
                                                     vml_drawing_id);
        }
    }
}

/*
 * Iterate through the worksheets and store any defined names used for print
 * ranges or repeat rows/columns.
 */
STATIC void
_prepare_defined_names(lxlsx_workbook *self)
{
    lxlsx_worksheet *worksheet;
    char app_name[LXLSX_DEFINED_NAME_LENGTH];
    char range[LXLSX_DEFINED_NAME_LENGTH];
    char area[LXLSX_MAX_CELL_RANGE_LENGTH];
    char first_col[8];
    char last_col[8];

    STAILQ_FOREACH(worksheet, self->worksheets, list_pointers) {

        /*
         * Check for autofilter settings and store them.
         */
        if (worksheet->autofilter.in_use) {

            lxlsx_snprintf(app_name, LXLSX_DEFINED_NAME_LENGTH,
                         "%s!_FilterDatabase", worksheet->quoted_name);

            lxlsx_rowcol_to_range_abs(area,
                                    worksheet->autofilter.first_row,
                                    worksheet->autofilter.first_col,
                                    worksheet->autofilter.last_row,
                                    worksheet->autofilter.last_col);

            lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH, "%s!%s",
                         worksheet->quoted_name, area);

            /* Autofilters are the only defined name to set the hidden flag. */
            _store_defined_name(self, "_xlnm._FilterDatabase", app_name, range, worksheet->index, LXLSX_TRUE);
        }

        /*
         * Check for Print Area settings and store them.
         */
        if (worksheet->print_area.in_use) {

            lxlsx_snprintf(app_name, LXLSX_DEFINED_NAME_LENGTH,
                         "%s!Print_Area", worksheet->quoted_name);

            /* Check for print area that is the max row range. */
            if (worksheet->print_area.first_row == 0
                && worksheet->print_area.last_row == LXLSX_ROW_MAX - 1) {

                lxlsx_col_to_name(first_col,
                                worksheet->print_area.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->print_area.last_col, LXLSX_FALSE);

                lxlsx_snprintf(area, LXLSX_MAX_CELL_RANGE_LENGTH - 1, "$%s:$%s",
                             first_col, last_col);

            }
                /* Check for print area that is the max column range. */
            else if (worksheet->print_area.first_col == 0
                     && worksheet->print_area.last_col == LXLSX_COL_MAX - 1) {

                lxlsx_snprintf(area, LXLSX_MAX_CELL_RANGE_LENGTH - 1, "$%d:$%d",
                             worksheet->print_area.first_row + 1,
                             worksheet->print_area.last_row + 1);

            }
            else {
                lxlsx_rowcol_to_range_abs(area,
                                        worksheet->print_area.first_row,
                                        worksheet->print_area.first_col,
                                        worksheet->print_area.last_row,
                                        worksheet->print_area.last_col);
            }

            lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH, "%s!%s",
                         worksheet->quoted_name, area);

            _store_defined_name(self, "_xlnm.Print_Area", app_name,
                                range, worksheet->index, LXLSX_FALSE);
        }

        /*
         * Check for repeat rows/cols. aka, Print Titles and store them.
         */
        if (worksheet->repeat_rows.in_use || worksheet->repeat_cols.in_use) {
            if (worksheet->repeat_rows.in_use
                && worksheet->repeat_cols.in_use) {
                lxlsx_snprintf(app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_col_to_name(first_col,
                                worksheet->repeat_cols.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->repeat_cols.last_col, LXLSX_FALSE);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%s:$%s,%s!$%d:$%d",
                             worksheet->quoted_name, first_col,
                             last_col, worksheet->quoted_name,
                             worksheet->repeat_rows.first_row + 1,
                             worksheet->repeat_rows.last_row + 1);

                _store_defined_name(self, "_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
            else if (worksheet->repeat_rows.in_use) {

                lxlsx_snprintf(app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%d:$%d", worksheet->quoted_name,
                             worksheet->repeat_rows.first_row + 1,
                             worksheet->repeat_rows.last_row + 1);

                _store_defined_name(self, "_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
            else if (worksheet->repeat_cols.in_use) {
                lxlsx_snprintf(app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_col_to_name(first_col,
                                worksheet->repeat_cols.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->repeat_cols.last_col, LXLSX_FALSE);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%s:$%s", worksheet->quoted_name,
                             first_col, last_col);

                _store_defined_name(self, "_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
        }
    }
}

/*
 * Track which image types the workbook contains so the packager can emit the
 * right MIME entries in [Content_Types].xml.
 */
STATIC void
_store_image_type(lxlsx_workbook *self, uint8_t image_type)
{
    if (image_type == LXLSX_IMAGE_PNG)  self->has_png  = LXLSX_TRUE;
    if (image_type == LXLSX_IMAGE_JPEG) self->has_jpeg = LXLSX_TRUE;
    if (image_type == LXLSX_IMAGE_BMP)  self->has_bmp  = LXLSX_TRUE;
    if (image_type == LXLSX_IMAGE_GIF)  self->has_gif  = LXLSX_TRUE;
}

/*
 * Iterate through the worksheets and set up chart, image, background, and
 * header/footer image drawings. Mirrors libxlsxwriter's static
 * _prepare_drawings() — without the background/header passes our previous
 * version skipped image_type bookkeeping and never called
 * lxlsx_worksheet_prepare_{background,header_image}, which produced malformed
 * xlsx files (missing media + Content_Types entries) when callers used
 * setBackgroundImage / setHeader([image_left|center|right]).
 */
STATIC void
_prepare_drawings(lxlsx_workbook *self)
{
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_object_properties *object_props;
    uint32_t lxlsx_chart_ref_id = 0;
    uint32_t image_ref_id = 0;
    uint32_t ref_id = 0;
    uint32_t drawing_id = 0;
    uint8_t i;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            worksheet = sheet->u.worksheet;

        if (STAILQ_EMPTY(worksheet->image_props)
            && STAILQ_EMPTY(worksheet->lxlsx_chart_data)
            && !worksheet->has_header_vml
            && !worksheet->has_background_image) {
            continue;
        }

        drawing_id++;

        /* Background image (no drawing_id; sheetBackground relationship). */
        if (worksheet->has_background_image) {
            object_props = worksheet->background_image;
            _store_image_type(self, object_props->image_type);
            image_ref_id++;
            ref_id = image_ref_id;
            lxlsx_worksheet_prepare_background(worksheet, ref_id, object_props);
        }

        /* Regular sheet images. */
        STAILQ_FOREACH(object_props, worksheet->image_props, list_pointers) {
            if (object_props->is_background)
                continue;
            _store_image_type(self, object_props->image_type);
            image_ref_id++;
            ref_id = image_ref_id;
            lxlsx_worksheet_prepare_image(worksheet, ref_id, drawing_id,
                                        object_props);
        }

        /* Charts. */
        STAILQ_FOREACH(object_props, worksheet->lxlsx_chart_data, list_pointers) {
            lxlsx_chart_ref_id++;
            lxlsx_worksheet_prepare_chart(worksheet, lxlsx_chart_ref_id, drawing_id,
                                        object_props, sheet->is_chartsheet);
            if (object_props->chart)
                STAILQ_INSERT_TAIL(self->ordered_charts, object_props->chart,
                                   ordered_list_pointers);
        }

        /* Header/footer images — &G tokens in setHeader/setFooter strings. */
        for (i = 0; i < LXLSX_HEADER_FOOTER_OBJS_MAX; i++) {
            object_props = *worksheet->header_footer_objs[i];
            if (!object_props)
                continue;
            _store_image_type(self, object_props->image_type);
            image_ref_id++;
            ref_id = image_ref_id;
            lxlsx_worksheet_prepare_header_image(worksheet, ref_id, object_props);
        }
    }

    self->lxlsx_drawing_count = drawing_id;
}

/*
 * Add "cached" data to charts to provide the numCache and strCache data for
 * series and title/axis ranges.
 */
STATIC void
_add_chart_cache_data(lxlsx_workbook *self)
{
    lxlsx_chart *chart;
    lxlsx_chart_series *series;

    STAILQ_FOREACH(chart, self->ordered_charts, ordered_list_pointers) {

        _populate_range(self, chart->title.range);
        _populate_range(self, chart->x_axis->title.range);
        _populate_range(self, chart->y_axis->title.range);

        if (STAILQ_EMPTY(chart->series_list))
            continue;

        STAILQ_FOREACH(series, chart->series_list, list_pointers) {
            _populate_range(self, series->categories);
            _populate_range(self, series->values);
            _populate_range(self, series->title.range);
        }
    }
}

/*
 * Iterate through the worksheets and assign 1-based ids to each table object,
 * mirroring libxlsxwriter's static _prepare_tables(). Without this, every
 * <table> element is written with id="0", which OnlyOffice / Numbers reject
 * when resolving structured references like [Sales] in total-row formulas.
 */
STATIC void
_prepare_tables(lxlsx_workbook *self)
{
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    uint32_t table_id = 0;
    uint32_t table_count = 0;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        worksheet = sheet->u.worksheet;

        table_count = worksheet->lxlsx_table_count;
        if (table_count == 0)
            continue;

        lxlsx_worksheet_prepare_tables(worksheet, table_id + 1);
        table_id += table_count;
    }
}

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
STATIC lxlsx_error
_store_defined_name(lxlsx_workbook *self, const char *name, const char *app_name, const char *formula, int16_t index, uint8_t hidden)
{
    lxlsx_worksheet *worksheet;
    lxlsx_defined_name *defined_name;
    lxlsx_defined_name *list_defined_name;
    char name_copy[LXLSX_DEFINED_NAME_LENGTH];
    char *tmp_str;
    char *lxlsx_worksheet_name;

    /* Do some checks on the input data */
    if (!name || !formula)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_utf8_strlen(name) > LXLSX_DEFINED_NAME_LENGTH ||
        lxlsx_utf8_strlen(formula) > LXLSX_DEFINED_NAME_LENGTH) {
        return LXLSX_ERROR_128_STRING_LENGTH_EXCEEDED;
    }

    /* Allocate a new defined_name to be added to the linked list of names. */
    defined_name = calloc(1, sizeof(struct lxlsx_defined_name));
    RETURN_ON_MEM_ERROR(defined_name, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Copy the user input string. */
    lxlsx_strcpy(name_copy, name);

    /* Set the worksheet index or -1 for a global defined name. */
    defined_name->index = index;
    defined_name->hidden = hidden;

    /* Check for local defined names like like "Sheet1!name". */
    tmp_str = strchr(name_copy, '!');

    if (tmp_str == NULL) {
        /* The name is global. We just store the defined name string. */
        lxlsx_strcpy(defined_name->name, name_copy);
    }
    else {
        /* The name is worksheet local. We need to extract the sheet name
         * and map it to a sheet index. */

        /* Split the into the worksheet name and defined name. */
        *tmp_str = '\0';
        tmp_str++;
        lxlsx_worksheet_name = name_copy;

        /* Remove any worksheet quoting. */
        if (lxlsx_worksheet_name[0] == '\'')
            lxlsx_worksheet_name++;
        if (lxlsx_worksheet_name[strlen(lxlsx_worksheet_name) - 1] == '\'')
            lxlsx_worksheet_name[strlen(lxlsx_worksheet_name) - 1] = '\0';

        /* Search for worksheet name to get the equivalent worksheet index. */
        STAILQ_FOREACH(worksheet, self->worksheets, list_pointers) {
            if (strcmp(lxlsx_worksheet_name, worksheet->name) == 0) {
                defined_name->index = worksheet->index;
                lxlsx_strcpy(defined_name->normalised_sheetname,
                           lxlsx_worksheet_name);
            }
        }

        /* If we didn't find the worksheet name we exit. */
        if (defined_name->index == -1)
            goto mem_error;

        lxlsx_strcpy(defined_name->name, tmp_str);
    }

    /* Print titles and repeat title pass in the name used for App.xml. */
    if (app_name) {
        lxlsx_strcpy(defined_name->lxlsx_app_name, app_name);
        lxlsx_strcpy(defined_name->normalised_sheetname, app_name);
    }
    else {
        lxlsx_strcpy(defined_name->lxlsx_app_name, name);
    }

    /* We need to normalize the defined names for sorting. This involves
     * removing any _xlnm namespace  and converting it to lowercase. */
    tmp_str = strstr(name_copy, "_xlnm.");

    if (tmp_str)
        lxlsx_strcpy(defined_name->normalised_name, defined_name->name + 6);
    else
        lxlsx_strcpy(defined_name->normalised_name, defined_name->name);

    lxlsx_str_tolower(defined_name->normalised_name);
    lxlsx_str_tolower(defined_name->normalised_sheetname);

    /* Strip leading "=" from the formula. */
    if (formula[0] == '=')
        lxlsx_strcpy(defined_name->formula, formula + 1);
    else
        lxlsx_strcpy(defined_name->formula, formula);

    /* We add the defined name to the list in sorted order. */
    list_defined_name = TAILQ_FIRST(self->defined_names);

    if (list_defined_name == NULL ||
        _compare_defined_names(defined_name, list_defined_name) < 1) {
        /* List is empty or defined name goes to the head. */
        TAILQ_INSERT_HEAD(self->defined_names, defined_name, list_pointers);
        return LXLSX_NO_ERROR;
    }

    TAILQ_FOREACH(list_defined_name, self->defined_names, list_pointers) {
        int res = _compare_defined_names(defined_name, list_defined_name);

        /* The entry already exists. We exit and don't overwrite. */
        if (res == 0)
            goto mem_error;

        /* New defined name is inserted in sorted order before other entries. */
        if (res < 0) {
            TAILQ_INSERT_BEFORE(list_defined_name, defined_name,
                                list_pointers);
            return LXLSX_NO_ERROR;
        }
    }

    /* If the entry wasn't less than any of the entries in the list we add it
     * to the end. */
    TAILQ_INSERT_TAIL(self->defined_names, defined_name, list_pointers);
    return LXLSX_NO_ERROR;

    mem_error:
    free(defined_name);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Compare two defined_name structures.
 */
static int
_compare_defined_names(lxlsx_defined_name *a, lxlsx_defined_name *b)
{
    int res = strcmp(a->normalised_name, b->normalised_name);

    /* Primary comparison based on defined name. */
    if (res)
        return res;

    /* Secondary comparison based on worksheet name. */
    res = strcmp(a->normalised_sheetname, b->normalised_sheetname);

    return res;
}

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
STATIC void
_populate_range_dimensions(lxlsx_workbook *self, lxlsx_series_range *range)
{

    char formula[LXLSX_MAX_FORMULA_RANGE_LENGTH] = { 0 };
    char *tmp_str;
    char *sheetname;

    /* If neither the range formula or sheetname is defined then this probably
     * isn't a valid range.
     */
    if (!range->formula && !range->sheetname) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* If the sheetname is already defined it was already set via
     * lxlsx_chart_series_set_categories() or  lxlsx_chart_series_set_values().
     */
    if (range->sheetname)
        return;

    /* Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5) */
    if (range->formula[0] == '(') {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Create a copy of the formula to modify and parse into parts. */
    lxlsx_snprintf(formula, LXLSX_MAX_FORMULA_RANGE_LENGTH, "%s", range->formula);

    /* Check for valid formula. TODO. This needs stronger validation. */
    tmp_str = strchr(formula, '!');

    if (tmp_str == NULL) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }
    else {
        /* Split the formulas into sheetname and row-col data. */
        *tmp_str = '\0';
        tmp_str++;
        sheetname = formula;

        /* Remove any worksheet quoting. */
        if (sheetname[0] == '\'')
            sheetname++;
        if (sheetname[strlen(sheetname) - 1] == '\'')
            sheetname[strlen(sheetname) - 1] = '\0';

        /* Check that the sheetname exists. */
        if (!lxlsx_workbook_get_worksheet_by_name(self, sheetname)) {
            LXLSX_WARN_FORMAT2("lxlsx_workbook_add_chart(): worksheet name '%s' "
                                     "in chart formula '%s' doesn't exist.",
                             sheetname, range->formula);
            range->ignore_cache = LXLSX_TRUE;
            return;
        }

        range->sheetname = lxlsx_strdup(sheetname);
        range->first_row = lxlsx_name_to_row(tmp_str);
        range->first_col = lxlsx_name_to_col(tmp_str);

        if (strchr(tmp_str, ':')) {
            /* 2D range. */
            range->last_row = lxlsx_name_to_row_2(tmp_str);
            range->last_col = lxlsx_name_to_col_2(tmp_str);
        }
        else {
            /* 1D range. */
            range->last_row = range->first_row;
            range->last_col = range->first_col;
        }

    }
}

/*
 * Populate the data cache of a chart data series by reading the data from the
 * relevant worksheet and adding it to the cached in the range object as a
 * list of points.
 *
 * Note, the data cache isn't strictly required by Excel but it helps if the
 * chart is embedded in another application such as PowerPoint and it also
 * helps with comparison testing.
 */
STATIC void
_populate_range_data_cache(lxlsx_workbook *self, lxlsx_series_range *range)
{
    lxlsx_worksheet *worksheet;
    lxlsx_row_t row_num;
    lxlsx_col_t col_num;
    lxlsx_row *row_obj;
    lxlsx_cell *cell_obj;
    struct lxlsx_series_data_point *data_point;
    uint16_t num_data_points = 0;

    /* If ignore_cache is set then don't try to populate the cache. This flag
     * may be set manually, for testing, or due to a case where the cache
     * can't be calculated.
     */
    if (range->ignore_cache)
        return;

    /* Currently we only handle 2D ranges so ensure either the rows or cols
     * are the same.
     */
    if (range->first_row != range->last_row
        && range->first_col != range->last_col) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Check that the sheetname exists. */
    worksheet = lxlsx_workbook_get_worksheet_by_name(self, range->sheetname);
    if (!worksheet) {
        LXLSX_WARN_FORMAT2("lxlsx_workbook_add_chart(): worksheet name '%s' "
                                 "in chart formula '%s' doesn't exist.",
                         range->sheetname, range->formula);
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* We can't read the data when worksheet optimization is on. */
    if (worksheet->optimize) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Iterate through the worksheet data and populate the range cache. */
    for (row_num = range->first_row; row_num <= range->last_row; row_num++) {
        row_obj = lxlsx_worksheet_find_row(worksheet, row_num);

        for (col_num = range->first_col; col_num <= range->last_col;
             col_num++) {

            data_point = calloc(1, sizeof(struct lxlsx_series_data_point));
            if (!data_point) {
                range->ignore_cache = LXLSX_TRUE;
                return;
            }

#if defined(LXLSX_VERSION_ID) && LXLSX_VERSION_ID >= 93
            cell_obj = lxlsx_worksheet_find_cell_in_row(row_obj, col_num);
#else
            cell_obj = lxlsx_worksheet_find_cell(row_obj, col_num);
#endif

            if (cell_obj) {
                if (cell_obj->type == NUMBER_CELL) {
                    data_point->number = cell_obj->data.writer.value.number;
                }

                if (cell_obj->type == STRING_CELL) {
                    data_point->string = lxlsx_strdup(
                        cell_obj->data.writer.value.shared_string.string);
                    data_point->is_string = LXLSX_TRUE;
                    range->has_string_cache = LXLSX_TRUE;
                }
            }
            else {
                data_point->no_data = LXLSX_TRUE;
            }

            STAILQ_INSERT_TAIL(range->data_cache, data_point, list_pointers);
            num_data_points++;
        }
    }

    range->num_data_points = num_data_points;
}

/* Set the range dimensions and set the data cache.
 */
STATIC void
_populate_range(lxlsx_workbook *self, lxlsx_series_range *range)
{
    _populate_range_dimensions(self, range);
    _populate_range_data_cache(self, range);
}
