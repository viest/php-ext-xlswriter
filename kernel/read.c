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

/* {{{ */
xlsxioreader file_open(const char *directory, const char *file_name) {
    char path[strlen(directory) + strlen(file_name) + 2];
    xlsxioreader file;

    lxw_strcpy(path, directory);
    strcat(path, "/");
    strcat(path, file_name);

    if ((file = xlsxioread_open(path)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Failed to open file", 100);
        return NULL;
    }

    return file;
}
/* }}} */

/* {{{ */
xlsxioreadersheet sheet_open(xlsxioreader file_t, const zend_string *zs_sheet_name_t)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_sheet_open(file_t, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
    }

    return xlsxioread_sheet_open(file_t, ZSTR_VAL(zs_sheet_name_t), XLSXIOREAD_SKIP_EMPTY_ROWS);
}
/* }}} */

/* {{{ */
int sheet_read_row(xlsxioreadersheet sheet_t)
{
    return xlsxioread_sheet_next_row(sheet_t);
}
/* }}} */

/* {{{ */
unsigned int load_sheet_current_row_data(xlsxioreadersheet sheet_t, zval *zv_result_t, zval *zv_type_arr_t, unsigned int flag)
{
    zend_long _type = READ_TYPE_EMPTY;
    char *_string_value = NULL;
    Bucket *_start = NULL, *_end = NULL;

    if (flag && !sheet_read_row(sheet_t)) {
        return XLSWRITER_FALSE;
    }

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    if (zv_type_arr_t != NULL && Z_TYPE_P(zv_type_arr_t) == IS_ARRAY) {
        _start = zv_type_arr_t->value.arr->arData;
        _end   = _start + zv_type_arr_t->value.arr->nNumUsed;
    }

    while ((_string_value = xlsxioread_sheet_next_cell(sheet_t)) != NULL)
    {
        if (_start != _end) {
            zval *_zv_type = &_start->val;

            if (Z_TYPE_P(_zv_type) == IS_LONG) {
                _type = Z_LVAL_P(_zv_type);
            }

            _start++;
        }

        if (_type & READ_TYPE_DATETIME) {
            double value = strtod(_string_value, NULL);

            if (value != 0) {
                value = (value - 25569) * 86400;
            }

            add_next_index_double(zv_result_t, value);
            continue;
        }

        if (_type & READ_TYPE_DOUBLE) {
            add_next_index_double(zv_result_t, strtod(_string_value, NULL));
            continue;
        }

        if (_type & READ_TYPE_INT) {
            zend_long _long_value;

            sscanf(_string_value, "%" PRIi64, &_long_value);
            add_next_index_long(zv_result_t, _long_value);
            continue;
        }

        add_next_index_stringl(zv_result_t, _string_value, strlen(_string_value));
    }

    return XLSWRITER_TRUE;
}
/* }}} */

/* {{{ */
int sheet_row_callback (size_t row, size_t max_col, void* callback_data)
{
    if (callback_data == NULL) {
        return FAILURE;
    }

    xls_read_callback_data *_callback_data = (xls_read_callback_data *)callback_data;

    zval args[3], retval;

    _callback_data->fci->retval      = &retval;
    _callback_data->fci->params      = args;
    _callback_data->fci->param_count = 3;

    ZVAL_LONG(&args[0], row);
    ZVAL_LONG(&args[1], max_col);
    ZVAL_STRING(&args[2], "XLSX_ROW_END");

    zend_call_function(_callback_data->fci, _callback_data->fci_cache);

    zval_ptr_dtor(&args[2]);
    zval_ptr_dtor(&retval);

    return SUCCESS;
}
/* }}} */

/* {{{ */
int sheet_cell_callback (size_t row, size_t col, const char *value, void *callback_data)
{
    if (callback_data == NULL) {
        return FAILURE;
    }

    xls_read_callback_data *_callback_data = (xls_read_callback_data *)callback_data;

    if (_callback_data->fci == NULL || _callback_data->fci_cache == NULL) {
        return FAILURE;
    }

    zval args[3], retval;

    _callback_data->fci->retval      = &retval;
    _callback_data->fci->params      = args;
    _callback_data->fci->param_count = 3;

    ZVAL_LONG(&args[0], row);
    ZVAL_LONG(&args[1], col);
    ZVAL_STRING(&args[2], value);

    zend_call_function(_callback_data->fci, _callback_data->fci_cache);

    zval_ptr_dtor(&args[2]);
    zval_ptr_dtor(&retval);

    return SUCCESS;
}
/* }}} */

/* {{{ */
unsigned int load_sheet_current_row_data_callback(zend_string *zs_sheet_name_t, xlsxioreader file_t, void *callback_data)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_process(file_t, NULL, XLSXIOREAD_SKIP_NONE, sheet_cell_callback, sheet_row_callback, callback_data);
    }

    return xlsxioread_process(file_t, ZSTR_VAL(zs_sheet_name_t), XLSXIOREAD_SKIP_NONE, sheet_cell_callback, sheet_row_callback, callback_data);
}
/* }}} */

/* {{{ */
void load_sheet_all_data(xlsxioreadersheet sheet_t, zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while (sheet_read_row(sheet_t))
    {
        zval _zv_tmp_row;
        ZVAL_NULL(&_zv_tmp_row);

        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, NULL, READ_SKIP_ROW);
        add_next_index_zval(zv_result_t, &_zv_tmp_row);
    }
}
/* }}} */