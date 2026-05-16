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
#include "ext/date/php_date.h"

/* ------------------------------------------------------------------------- */
/* Cell -> string view                                                       */
/* ------------------------------------------------------------------------- */

static void cell_text_view(const lxr_cell *c, const char **out_ptr, size_t *out_len)
{
    if (!c) { *out_ptr = ""; *out_len = 0; return; }
    switch (c->type) {
    case LXR_CELL_BLANK:
        *out_ptr = ""; *out_len = 0; return;
    case LXR_CELL_STRING:
    case LXR_CELL_INLINE_STRING:
        *out_ptr = c->value.string.ptr ? c->value.string.ptr : "";
        *out_len = c->value.string.len;
        return;
    case LXR_CELL_NUMBER:
    case LXR_CELL_DATETIME:
    case LXR_CELL_BOOLEAN:
    case LXR_CELL_ERROR:
    case LXR_CELL_FORMULA:
    default:
        *out_ptr = c->raw.ptr ? c->raw.ptr : "";
        *out_len = c->raw.len;
        return;
    }
}

/* ------------------------------------------------------------------------- */
/* Open helpers                                                              */
/* ------------------------------------------------------------------------- */

lxr_workbook *file_open(const char *directory, const char *file_name) {
    char         *path = (char *)emalloc(strlen(directory) + strlen(file_name) + 2);
    lxr_workbook *wb   = NULL;
    lxr_error     rc;

    strcpy(path, directory);
    strcat(path, "/");
    strcat(path, file_name);

    if (file_exists(path) == XLSWRITER_FALSE) {
        zend_string *message = char_join_to_zend_str("File not found, file path:", path);
        zend_throw_exception(vtiful_exception_ce, ZSTR_VAL(message), 121);
        zend_string_free(message);
        efree(path);
        return NULL;
    }

    rc = lxr_workbook_open(path, &wb);
    if (rc != LXR_NO_ERROR || wb == NULL) {
        zend_string *message = char_join_to_zend_str("Failed to open file, file path:", path);
        zend_throw_exception(vtiful_exception_ce, ZSTR_VAL(message), 100);
        zend_string_free(message);
        efree(path);
        return NULL;
    }
    efree(path);
    return wb;
}

lxr_worksheet *sheet_open(lxr_workbook *wb, const zend_string *zs_sheet_name_t, const zend_long zl_flag)
{
    lxr_worksheet *ws = NULL;
    const char    *name = zs_sheet_name_t ? ZSTR_VAL(zs_sheet_name_t) : NULL;
    if (lxr_workbook_get_worksheet_by_name(wb, name, (uint32_t)zl_flag, &ws) != LXR_NO_ERROR) {
        return NULL;
    }
    return ws;
}

void sheet_list(lxr_workbook *wb, zval *zv_result_t)
{
    size_t i, n;
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }
    n = lxr_workbook_sheet_count(wb);
    for (i = 0; i < n; i++) {
        const char *name = lxr_workbook_sheet_name(wb, i);
        if (name) add_next_index_stringl(zv_result_t, name, strlen(name));
    }
}

void sheet_list_with_meta(lxr_workbook *wb, zval *zv_result_t)
{
    size_t i, n;
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }
    n = lxr_workbook_sheet_count(wb);
    for (i = 0; i < n; i++) {
        const char *name = lxr_workbook_sheet_name(wb, i);
        const char *state;
        zval entry;
        if (!name) continue;
        switch (lxr_workbook_sheet_visibility(wb, i)) {
            case LXR_SHEET_HIDDEN:      state = "hidden";     break;
            case LXR_SHEET_VERY_HIDDEN: state = "veryHidden"; break;
            default:                    state = "visible";    break;
        }
        array_init(&entry);
        add_assoc_stringl(&entry, "name", name, strlen(name));
        add_assoc_string (&entry, "state", state);
        add_next_index_zval(zv_result_t, &entry);
    }
}

/* ------------------------------------------------------------------------- */
/* Type helpers                                                              */
/* ------------------------------------------------------------------------- */

int is_number(const char *value)
{
    return strspn(value, ".0123456789") == strlen(value) ? XLSWRITER_TRUE : XLSWRITER_FALSE;
}

void data_to_null(zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
        add_next_index_null(zv_result_t);
    } else {
        ZVAL_NULL(zv_result_t);
    }
}

void data_to_custom_type(const char *string_value, const size_t string_value_length, const zend_ulong type, zval *zv_result_t, const zend_ulong zv_hashtable_index)
{
    if (type == 0) goto STRING;
    if (!is_number(string_value)) goto STRING;

    if (type & READ_TYPE_DATETIME) {
        if (string_value_length == 0) { data_to_null(zv_result_t); return; }
        zend_long timestamp = date_double_to_timestamp(zend_strtod(string_value, NULL));
        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) add_index_long(zv_result_t, zv_hashtable_index, timestamp);
        else                                   ZVAL_LONG(zv_result_t, timestamp);
        return;
    }

    if (type & READ_TYPE_DOUBLE) {
        if (string_value_length == 0) { data_to_null(zv_result_t); return; }
        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) add_index_double(zv_result_t, zv_hashtable_index, strtod(string_value, NULL));
        else                                   ZVAL_DOUBLE(zv_result_t, strtod(string_value, NULL));
        return;
    }

    if (type & READ_TYPE_INT) {
        if (string_value_length == 0) { data_to_null(zv_result_t); return; }
        zend_long _long_value;
        sscanf(string_value, ZEND_LONG_FMT, &_long_value);
        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) add_index_long(zv_result_t, zv_hashtable_index, _long_value);
        else                                   ZVAL_LONG(zv_result_t, _long_value);
        return;
    }

    STRING:
    {
        if (!(type & READ_TYPE_STRING)) {
            zend_long _long = 0; double _double = 0;
            is_numeric_string(string_value, string_value_length, &_long, &_double, 0);

            if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
                if (_double > 0 && _double <= (double)ZEND_LONG_MAX) {
                    add_index_double(zv_result_t, zv_hashtable_index, _double);
                    return;
                }
                if (_long > 0 && _long <= ZEND_LONG_MAX) {
                    add_index_long(zv_result_t, zv_hashtable_index, _long);
                    return;
                }
            } else {
                if (_double > 0 && _double <= (double)ZEND_LONG_MAX) {
                    ZVAL_DOUBLE(zv_result_t, _double);
                    return;
                }
                if (_long > 0 && _long <= ZEND_LONG_MAX) {
                    ZVAL_LONG(zv_result_t, _long);
                    return;
                }
            }
        }

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_index_stringl(zv_result_t, zv_hashtable_index, string_value, string_value_length);
            return;
        }
        ZVAL_STRINGL(zv_result_t, string_value, string_value_length);
    }
}

/* ------------------------------------------------------------------------- */
/* Row / cell streaming                                                      */
/* ------------------------------------------------------------------------- */

int sheet_read_row(lxr_worksheet *ws)
{
    return lxr_worksheet_next_row(ws) == LXR_NO_ERROR ? 1 : 0;
}

/* Apply user type (or global default) to either a real cell or a synthesised
 * blank, writing the result into zv_result_t at index idx (0-based). */
static void emit_typed_value(zval *zv_result_t, zend_array *za_type, zend_long data_type_default,
                             const char *str, size_t str_len, zend_ulong idx)
{
    zend_long _type = data_type_default;
    if (za_type) {
        zval *t = zend_hash_index_find(za_type, (zend_long)idx);
        if (t && Z_TYPE_P(t) == IS_LONG) _type = Z_LVAL_P(t);
    }
    data_to_custom_type(str, str_len, (zend_ulong)_type, zv_result_t, idx);
}

unsigned int load_sheet_current_row_data(struct xls_resource_read_t *r, zval *zv_result_t,
                                         zval *zv_type_arr_t, zend_long data_type_default,
                                         unsigned int flag)
{
    if (!r || !r->sheet_t) return XLSWRITER_FALSE;

    if (flag && !sheet_read_row(r->sheet_t)) {
        return XLSWRITER_FALSE;
    }

    uint32_t      ws_flags         = lxr_worksheet_flags(r->sheet_t);
    int           skip_empty_cells = (ws_flags & LXR_SKIP_EMPTY_CELLS)   != 0;
    int           skip_empty_value = (ws_flags & SKIP_EMPTY_VALUE)       != 0;
    int           skip_merged_foll = (ws_flags & LXR_SKIP_MERGED_FOLLOW) != 0;
    zend_array   *za_type          = (zv_type_arr_t && Z_TYPE_P(zv_type_arr_t) == IS_ARRAY)
                                       ? Z_ARR_P(zv_type_arr_t) : NULL;
    size_t        expected_col     = 1;
    size_t        row_max_col      = 0;
    int           saw_real_cell    = 0;
    size_t        row_nr           = lxr_worksheet_current_row(r->sheet_t);
    lxr_cell      cell;

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while (lxr_worksheet_next_cell(r->sheet_t, &cell) == LXR_NO_ERROR) {
        const char *str;
        size_t      str_len;
        size_t      cur_col;

        cell_text_view(&cell, &str, &str_len);
        if (skip_empty_value && str_len == 0) continue;

        cur_col = cell.col > 0 ? cell.col : 1;

        /* Lead / intermediate gap blanks. Both SKIP_EMPTY_CELLS and
         * SKIP_EMPTY_VALUE suppress synthesised blanks, mirroring how
         * libxlsxio's empty placeholder cells were filtered. */
        if (!skip_empty_cells && !skip_empty_value) {
            while (expected_col < cur_col) {
                if (skip_merged_foll &&
                    lxr_worksheet_in_merge_follow(r->sheet_t, row_nr, expected_col)) {
                    add_index_null(zv_result_t, (zend_ulong)(expected_col - 1));
                } else {
                    emit_typed_value(zv_result_t, za_type, data_type_default,
                                     "", 0, (zend_ulong)(expected_col - 1));
                }
                expected_col++;
            }
        } else if (skip_empty_value || skip_empty_cells) {
            /* Advance the column cursor without emitting blanks so the index
             * we use for cur_col still matches the cell's real position. */
            expected_col = cur_col;
        }

        if (skip_merged_foll &&
            lxr_worksheet_in_merge_follow(r->sheet_t, row_nr, cur_col)) {
            add_index_null(zv_result_t, (zend_ulong)(cur_col - 1));
        } else {
            emit_typed_value(zv_result_t, za_type, data_type_default,
                             str, str_len, (zend_ulong)(cur_col - 1));
        }

        expected_col  = cur_col + 1;
        if (cur_col > row_max_col) row_max_col = cur_col;
        saw_real_cell = 1;
    }

    /* Trailing blanks up to the cached first-row width. Both
     * SKIP_EMPTY_CELLS and SKIP_EMPTY_VALUE suppress synthesised trailing
     * blanks. */
    if (!skip_empty_cells && !skip_empty_value && saw_real_cell && r->cols > 0) {
        while (expected_col <= r->cols) {
            if (skip_merged_foll &&
                lxr_worksheet_in_merge_follow(r->sheet_t, row_nr, expected_col)) {
                add_index_null(zv_result_t, (zend_ulong)(expected_col - 1));
            } else {
                emit_typed_value(zv_result_t, za_type, data_type_default,
                                 "", 0, (zend_ulong)(expected_col - 1));
            }
            expected_col++;
        }
    }

    /* First non-empty row defines the canonical column count. */
    if (r->cols == 0 && row_max_col > 0) {
        r->cols = row_max_col;
    }

    return XLSWRITER_TRUE;
}

/* ------------------------------------------------------------------------- */
/* Callback bridge                                                           */
/* ------------------------------------------------------------------------- */

static int lxr_row_end_bridge(size_t row, size_t max_col, void *callback_data)
{
    xls_read_callback_data *_cd = (xls_read_callback_data *)callback_data;
    zval args[3], retval;

    if (!_cd || !_cd->fci || !_cd->fci_cache) return 0;

    _cd->fci->retval      = &retval;
    _cd->fci->params      = args;
    _cd->fci->param_count = 3;

    ZVAL_LONG(&args[0], (zend_long)(row - 1));
    ZVAL_LONG(&args[1], (zend_long)(max_col - 1));
    ZVAL_STRING(&args[2], "XLSX_ROW_END");

    zend_call_function(_cd->fci, _cd->fci_cache);
    zval_ptr_dtor(&args[2]);
    zval_ptr_dtor(&retval);
    return 0;
}

static int lxr_cell_bridge(const lxr_cell *c, void *callback_data)
{
    xls_read_callback_data *_cd = (xls_read_callback_data *)callback_data;
    const char *str;
    size_t      str_len;
    zval        args[3], retval;

    if (!_cd || !_cd->fci || !_cd->fci_cache) return 0;
    cell_text_view(c, &str, &str_len);

    _cd->fci->retval      = &retval;
    _cd->fci->params      = args;
    _cd->fci->param_count = 3;

    ZVAL_LONG(&args[0], (zend_long)(c->row - 1));
    ZVAL_LONG(&args[1], (zend_long)(c->col - 1));
    ZVAL_NULL(&args[2]);

    if (c->type == LXR_CELL_BLANK) goto CALL;

    if (Z_TYPE_P(_cd->zv_type_t) != IS_ARRAY && _cd->data_type_default == READ_TYPE_EMPTY) {
        zend_long _long = 0; double _double = 0;
        if (is_numeric_string(str, str_len, &_long, &_double, 0)) {
            /* Both branches expand to macros that PHP 7.4 defines as bare
             * `{ ... }` blocks; without explicit braces the trailing `;`
             * orphans the else. Same shape as the fix in kernel/excel.c. */
            if (_double > 0) {
                ZVAL_DOUBLE(&args[2], _double);
            } else {
                ZVAL_LONG(&args[2], _long);
            }
        } else {
            ZVAL_STRINGL(&args[2], str, str_len);
        }
    }

    if (Z_TYPE_P(_cd->zv_type_t) != IS_ARRAY && _cd->data_type_default != READ_TYPE_EMPTY) {
        data_to_custom_type(str, str_len, _cd->data_type_default, &args[2], 0);
    }

    if (Z_TYPE_P(_cd->zv_type_t) == IS_ARRAY) {
        zval      *t     = zend_hash_index_find(Z_ARR_P(_cd->zv_type_t), (zend_long)(c->col - 1));
        zend_ulong _type = (t && Z_TYPE_P(t) == IS_LONG) ? Z_LVAL_P(t) : READ_TYPE_EMPTY;
        data_to_custom_type(str, str_len, _type, &args[2], 0);
    }

    CALL:

    zend_call_function(_cd->fci, _cd->fci_cache);
    zval_ptr_dtor(&args[2]);
    zval_ptr_dtor(&retval);
    return 0;
}

unsigned int load_sheet_current_row_data_callback(zend_string *zs_sheet_name_t,
                                                  lxr_workbook *wb, void *callback_data)
{
    lxr_worksheet *ws   = NULL;
    const char    *name = zs_sheet_name_t ? ZSTR_VAL(zs_sheet_name_t) : NULL;
    lxr_error      rc;

    if (lxr_workbook_get_worksheet_by_name(wb, name, LXR_SKIP_NONE, &ws) != LXR_NO_ERROR) {
        return 0;
    }
    rc = lxr_worksheet_process(ws, lxr_cell_bridge, lxr_row_end_bridge, callback_data);
    lxr_worksheet_close(ws);
    return rc == LXR_NO_ERROR ? 1 : 0;
}

/* ------------------------------------------------------------------------- */
/* High-level loaders                                                        */
/* ------------------------------------------------------------------------- */

static int row_is_empty(const zval *row)
{
    return Z_TYPE_P(row) == IS_ARRAY && zend_hash_num_elements(Z_ARRVAL_P(row)) == 0;
}

void load_sheet_row_data(struct xls_resource_read_t *r, zend_long sheet_flag, zval *zv_type_t,
                         zend_long data_type_default, zval *zv_result_t)
{
    if (!r || !r->sheet_t) return;
    if (r->expected_row_nr == 0) r->expected_row_nr = 1;

    if (r->pending_synth_rows > 0) {
        r->pending_synth_rows--;
        array_init(zv_result_t);
        return;
    }

    if (Z_TYPE(r->pending_real_row) == IS_ARRAY) {
        ZVAL_COPY_VALUE(zv_result_t, &r->pending_real_row);
        ZVAL_NULL(&r->pending_real_row);
        return;
    }

    while (1) {
        if (!sheet_read_row(r->sheet_t)) {
            return;  /* EOF */
        }

        size_t cur_row = lxr_worksheet_current_row(r->sheet_t);
        zval   row;
        ZVAL_NULL(&row);
        load_sheet_current_row_data(r, &row, zv_type_t, data_type_default, READ_SKIP_ROW);

        if ((sheet_flag & LXR_SKIP_EMPTY_ROWS) && row_is_empty(&row)) {
            zval_ptr_dtor(&row);
            r->expected_row_nr = cur_row + 1;
            continue;
        }

        if (!(sheet_flag & LXR_SKIP_EMPTY_ROWS) && cur_row > r->expected_row_nr) {
            size_t gap = cur_row - r->expected_row_nr;
            array_init(zv_result_t);
            r->pending_synth_rows = gap - 1;
            ZVAL_COPY_VALUE(&r->pending_real_row, &row);
            r->expected_row_nr = cur_row + 1;
            return;
        }

        ZVAL_COPY_VALUE(zv_result_t, &row);
        r->expected_row_nr = cur_row + 1;
        return;
    }
}

void load_sheet_all_data(struct xls_resource_read_t *r, zend_long sheet_flag, zval *zv_type_t,
                         zend_long data_type_default, zval *zv_result_t)
{
    if (!r || !r->sheet_t) {
        if (Z_TYPE_P(zv_result_t) != IS_ARRAY) array_init(zv_result_t);
        return;
    }
    if (r->expected_row_nr == 0) r->expected_row_nr = 1;
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) array_init(zv_result_t);

    while (sheet_read_row(r->sheet_t)) {
        size_t cur_row = lxr_worksheet_current_row(r->sheet_t);
        zval   row;
        ZVAL_NULL(&row);

        load_sheet_current_row_data(r, &row, zv_type_t, data_type_default, READ_SKIP_ROW);

        if ((sheet_flag & LXR_SKIP_EMPTY_ROWS) && row_is_empty(&row)) {
            zval_ptr_dtor(&row);
            r->expected_row_nr = cur_row + 1;
            continue;
        }

        if (!(sheet_flag & LXR_SKIP_EMPTY_ROWS)) {
            while (r->expected_row_nr < cur_row) {
                zval empty;
                array_init(&empty);
                add_next_index_zval(zv_result_t, &empty);
                r->expected_row_nr++;
            }
        }

        add_next_index_zval(zv_result_t, &row);
        r->expected_row_nr = cur_row + 1;
    }
}

void skip_rows(struct xls_resource_read_t *r, zval *zv_type_t, zend_long data_type_default, zend_long zl_skip_row)
{
    (void)zv_type_t;
    (void)data_type_default;
    if (!r || !r->sheet_t || zl_skip_row <= 0) return;
    lxr_worksheet_skip_rows(r->sheet_t, (size_t)zl_skip_row);
    if (r->expected_row_nr == 0) r->expected_row_nr = 1;
    r->expected_row_nr += (size_t)zl_skip_row;
}
