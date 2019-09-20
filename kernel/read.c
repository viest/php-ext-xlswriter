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
xlsxioreadersheet sheet_open(xlsxioreader file_t, const zend_string *zs_sheet_name_t, const zend_long zl_flag)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_sheet_open(file_t, NULL, zl_flag);
    }

    return xlsxioread_sheet_open(file_t, ZSTR_VAL(zs_sheet_name_t), zl_flag);
}
/* }}} */

/* {{{ */
int is_number(const char *value)
{
    if (strspn(value, ".0123456789") == strlen(value)) {
        return XLSWRITER_TRUE;
    }

    return XLSWRITER_FALSE;
}
/* }}} */

/* {{{ */
void data_to_custom_type(const char *string_value, zend_ulong type, zval *zv_result_t)
{
    if (type & READ_TYPE_DATETIME) {
        if (!is_number(string_value)) {
            goto STRING;
        }

        double value = strtod(string_value, NULL);

        if (value != 0) {
            value = (value - 25569) * 86400;
        }

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_next_index_long(zv_result_t, (zend_long)(value + 0.5));
        } else {
            ZVAL_LONG(zv_result_t, (zend_long)(value + 0.5));
        }

        return;
    }

    if (type & READ_TYPE_DOUBLE) {
        if (!is_number(string_value)) {
            goto STRING;
        }

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_next_index_double(zv_result_t, strtod(string_value, NULL));
        } else {
            ZVAL_DOUBLE(zv_result_t, strtod(string_value, NULL));
        }

        return;
    }

    if (type & READ_TYPE_INT) {
        if (!is_number(string_value)) {
            goto STRING;
        }

        zend_long _long_value;

        sscanf(string_value, ZEND_LONG_FMT, &_long_value);

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_next_index_long(zv_result_t, _long_value);
        } else {
            ZVAL_LONG(zv_result_t, _long_value);
        }

        return;
    }

    STRING:

    if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
        add_next_index_stringl(zv_result_t, string_value, strlen(string_value));
    } else {
        ZVAL_STRINGL(zv_result_t, string_value, strlen(string_value));
    }
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
    zend_ulong _type, _cell_index = 0;
    zend_array *_za_type_t = NULL;
    char *_string_value = NULL;
    zval *_current_type = NULL;

    if (flag && !sheet_read_row(sheet_t)) {
        return XLSWRITER_FALSE;
    }

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    if (zv_type_arr_t != NULL && Z_TYPE_P(zv_type_arr_t) == IS_ARRAY) {
        _za_type_t = Z_ARR_P(zv_type_arr_t);
    }

    while ((_string_value = xlsxioread_sheet_next_cell(sheet_t)) != NULL)
    {
        _type = READ_TYPE_EMPTY;

        if (_za_type_t != NULL) {
            if ((_current_type = zend_hash_index_find(_za_type_t, _cell_index)) != NULL) {
                if (Z_TYPE_P(_current_type) == IS_LONG) {
                    _type = Z_LVAL_P(_current_type);
                }
            }

            _cell_index++;
        }

        data_to_custom_type(_string_value, _type, zv_result_t);
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

    ZVAL_LONG(&args[0], (row - 1));
    ZVAL_LONG(&args[1], (max_col - 1));
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

    ZVAL_LONG(&args[0], (row - 1));
    ZVAL_LONG(&args[1], (col - 1));

    if (Z_TYPE_P(_callback_data->zv_type_t) != IS_ARRAY) {
        ZVAL_STRING(&args[2], value);
    }

    if (Z_TYPE_P(_callback_data->zv_type_t) == IS_ARRAY) {
        zval *_current_type = NULL;
        zend_ulong _type = READ_TYPE_EMPTY;

        if ((_current_type = zend_hash_index_find(Z_ARR_P(_callback_data->zv_type_t), (col - 1))) != NULL) {
            if (Z_TYPE_P(_current_type) == IS_LONG) {
                _type = Z_LVAL_P(_current_type);
            }
        }

        data_to_custom_type(value, _type, &args[2]);
    }

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
void load_sheet_all_data(xlsxioreadersheet sheet_t, zval *zv_type_t, zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while (sheet_read_row(sheet_t))
    {
        zval _zv_tmp_row;
        ZVAL_NULL(&_zv_tmp_row);

        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, zv_type_t, READ_SKIP_ROW);
        add_next_index_zval(zv_result_t, &_zv_tmp_row);
    }
}
/* }}} */
