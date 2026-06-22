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
    zend_string *path = zval_get_string(value);
    int error = lxlsx_worksheet_insert_image_opt(res->worksheet, (lxlsx_row_t)row, (lxlsx_col_t)columns,
                                           ZSTR_VAL(path), options);
    zend_string_release(path);
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

    /* Normalize reversed ranges (e.g. "A10:A1"): without this the loop would
     * write row `end` then underflow `end` past 0 before the row-max guard
     * trips, leaving one stray row behind. */
    if (start > end) {
        lxlsx_row_t tmp = start;
        start = end;
        end = tmp;
    }

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
 * Run workbook finalization and write the file, without freeing the workbook
 * (the extension's resource owns the workbook and frees it in its object
 * destructor). Edit mode patches the original archive in place; normal mode
 * delegates to libxlsx's shared finalization pipeline via
 * lxlsx_workbook_assemble() — the same code lxlsx_workbook_close() runs, minus
 * the trailing lxlsx_workbook_free(). This replaces the ~840 lines of workbook
 * finalization that used to be copied here.
 */
lxlsx_error
lxlsx_workbook_file(xls_resource_write_t *self)
{
    if (lxlsx_workbook_is_edit(self->workbook)) {
        return lxlsx_workbook_save_as(self->workbook, self->workbook->filename);
    }

    return lxlsx_workbook_assemble(self->workbook);
}

void _php_vtiful_xls_close(zend_resource *rsrc TSRMLS_DC)
{
    //
}

