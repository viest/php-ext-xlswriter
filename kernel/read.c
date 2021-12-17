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

/* {{{ */
xlsxioreader file_open(const char *directory, const char *file_name) {
    char *path = (char *)emalloc(strlen(directory) + strlen(file_name) + 2);
    xlsxioreader file;

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

    if ((file = xlsxioread_open(path)) == NULL) {
        zend_string *message = char_join_to_zend_str("Failed to open file, file path:", path);
        zend_throw_exception(vtiful_exception_ce, ZSTR_VAL(message), 100);

        zend_string_free(message);
        efree(path);

        return NULL;
    }

    efree(path);
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
void sheet_list(xlsxioreader file_t, zval *zv_result_t)
{
    const char *sheet_name = NULL;
    xlsxioreadersheetlist sheet_list = NULL;

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    if ((sheet_list = xlsxioread_sheetlist_open(file_t)) == NULL) {
        return;
    }

    while ((sheet_name = xlsxioread_sheetlist_next(sheet_list)) != NULL) {
        add_next_index_stringl(zv_result_t, sheet_name, strlen(sheet_name));
    }

    xlsxioread_sheetlist_close(sheet_list);
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
void data_to_null(zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
        add_next_index_null(zv_result_t);
    } else {
        ZVAL_NULL(zv_result_t);
    }
}
/* }}} */

/* {{{ */
void data_to_custom_type(const char *string_value, const size_t string_value_length, const zend_ulong type, zval *zv_result_t, const zend_ulong zv_hashtable_index)
{
    if (type == 0) {
        goto STRING;
    }

    if (!is_number(string_value)) {
        goto STRING;
    }

    if (type & READ_TYPE_DATETIME) {
        if (string_value_length == 0) {
            data_to_null(zv_result_t);

            return;
        }

        zend_long timestamp = date_double_to_timestamp(zend_strtod(string_value, NULL));

        // GMT
        // if (value != 0) {
        //     timestamp = (value - 25569) * 86400;
        // }

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_index_long(zv_result_t, zv_hashtable_index, timestamp);
        } else {
            ZVAL_LONG(zv_result_t, timestamp);
        }

        return;
    }

    if (type & READ_TYPE_DOUBLE) {
        if (string_value_length == 0) {
            data_to_null(zv_result_t);

            return;
        }

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_index_double(zv_result_t, zv_hashtable_index,strtod(string_value, NULL));
        } else {
            ZVAL_DOUBLE(zv_result_t, strtod(string_value, NULL));
        }

        return;
    }

    if (type & READ_TYPE_INT) {
        if (string_value_length == 0) {
            data_to_null(zv_result_t);

            return;
        }

        zend_long _long_value;

        sscanf(string_value, ZEND_LONG_FMT, &_long_value);

        if (Z_TYPE_P(zv_result_t) == IS_ARRAY) {
            add_index_long(zv_result_t, zv_hashtable_index, _long_value);
        } else {
            ZVAL_LONG(zv_result_t, _long_value);
        }

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
/* }}} */

/* {{{ */
int sheet_read_row(xlsxioreadersheet sheet_t)
{
    return xlsxioread_sheet_next_row(sheet_t);
}
/* }}} */

/* {{{ */
unsigned int load_sheet_current_row_data(xlsxioreadersheet sheet_t, zval *zv_result_t, zval *zv_type_arr_t, zend_long data_type_default, unsigned int flag)
{
    zend_long _type, _cell_index = 0, _last_cell_index = 0;
    zend_bool _skip_empty_value_cell = 0;
    zend_array *_za_type_t = NULL;
    char *_string_value = NULL;
    zval *_current_type = NULL;

    if (flag && !sheet_read_row(sheet_t)) {
        return XLSWRITER_FALSE;
    }

    if (xlsxioread_sheet_flags(sheet_t) & SKIP_EMPTY_VALUE) {
        _skip_empty_value_cell = 1;
    }

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    if (zv_type_arr_t != NULL && Z_TYPE_P(zv_type_arr_t) == IS_ARRAY) {
        _za_type_t = Z_ARR_P(zv_type_arr_t);
    }

    while ((_string_value = xlsxioread_sheet_next_cell(sheet_t)) != NULL)
    {
        size_t _string_value_length = strlen(_string_value);

        _type = READ_TYPE_EMPTY;
        _last_cell_index = xlsxioread_sheet_last_column_index(sheet_t) - 1;

        if (_last_cell_index < 0 || (_skip_empty_value_cell && _string_value_length == 0)) {
            goto FREE_TMP_VALUE;
        }

        if (_last_cell_index > _cell_index) {
            _cell_index = _last_cell_index;
        }

        if (_za_type_t != NULL) {
            _current_type = zend_hash_index_find(_za_type_t, _cell_index);

            if (_current_type != NULL && Z_TYPE_P(_current_type) == IS_LONG) {
                _type = Z_LVAL_P(_current_type);
            } else {
                _type = data_type_default;
            }
        } else {
            _type = data_type_default;
        }

        data_to_custom_type(_string_value, _string_value_length, _type, zv_result_t, _cell_index);

        FREE_TMP_VALUE:

        ++_cell_index;
        free(_string_value);
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
    size_t _value_length = 0;

    if (value != NULL) {
        _value_length = strlen(value);
    }

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
    ZVAL_NULL(&args[2]);

    if (value == NULL) {
        goto CALL_USER_FUNCTION;
    }

    if (Z_TYPE_P(_callback_data->zv_type_t) != IS_ARRAY && _callback_data->data_type_default == READ_TYPE_EMPTY) {
        zend_long _long = 0; double _double = 0;

        if (is_numeric_string(value, _value_length, &_long, &_double, 0)) {
            if (_double > 0) {
                ZVAL_DOUBLE(&args[2], _double);
            } else {
                ZVAL_LONG(&args[2], _long);
            }
        } else {
            ZVAL_STRINGL(&args[2], value, _value_length);
        }
    }

    if (Z_TYPE_P(_callback_data->zv_type_t) != IS_ARRAY && _callback_data->data_type_default != READ_TYPE_EMPTY) {
        data_to_custom_type(value, _value_length, _callback_data->data_type_default, &args[2], 0);
    }

    if (Z_TYPE_P(_callback_data->zv_type_t) == IS_ARRAY) {
        zval *_current_type = NULL;
        zend_ulong _type = READ_TYPE_EMPTY;

        if ((_current_type = zend_hash_index_find(Z_ARR_P(_callback_data->zv_type_t), (col - 1))) != NULL) {
            if (Z_TYPE_P(_current_type) == IS_LONG) {
                _type = Z_LVAL_P(_current_type);
            }
        }

        data_to_custom_type(value, _value_length, _type, &args[2], 0);
    }

    CALL_USER_FUNCTION:

    zend_call_function(_callback_data->fci, _callback_data->fci_cache);

    zval_ptr_dtor(&args[2]);
    zval_ptr_dtor(&retval);

    return SUCCESS;
}
/* }}} */

/* {{{ */
unsigned int load_sheet_current_row_data_callback (zend_string *zs_sheet_name_t, xlsxioreader file_t, void *callback_data)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_process(file_t, NULL, XLSXIOREAD_SKIP_NONE, sheet_cell_callback, sheet_row_callback, callback_data);
    }

    return xlsxioread_process(file_t, ZSTR_VAL(zs_sheet_name_t), XLSXIOREAD_SKIP_NONE, sheet_cell_callback, sheet_row_callback, callback_data);
}
/* }}} */

/* {{{ */
void load_sheet_row_data (xlsxioreadersheet sheet_t, zend_long sheet_flag, zval *zv_type_t, zend_long data_type_default, zval *zv_result_t)
{
    size_t row_index = 0;

    do {
        load_sheet_current_row_data(sheet_t, zv_result_t, zv_type_t, data_type_default, READ_ROW);

        if (row_index == xlsxioread_sheet_last_row_index(sheet_t)) {
            return;
        }

        row_index = xlsxioread_sheet_last_row_index(sheet_t);

        if (sheet_flag & XLSXIOREAD_SKIP_EMPTY_ROWS
            && Z_TYPE_P(zv_result_t) == IS_ARRAY
            && zend_hash_num_elements(Z_ARR_P(zv_result_t)) == 0) {
            continue;
        }

        return;
    } while (1);
}
/* }}} */

/* {{{ */
void load_sheet_all_data (xlsxioreadersheet sheet_t, zend_long sheet_flag, zval *zv_type_t, zend_long data_type_default, zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while (sheet_read_row(sheet_t))
    {
        zval _zv_tmp_row;
        ZVAL_NULL(&_zv_tmp_row);

        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, zv_type_t, data_type_default, READ_SKIP_ROW);

        if (sheet_flag & XLSXIOREAD_SKIP_EMPTY_ROWS
            && Z_TYPE(_zv_tmp_row) == IS_ARRAY
            && zend_hash_num_elements(Z_ARR(_zv_tmp_row)) == 0) {
            zval_ptr_dtor(&_zv_tmp_row);
            continue;
        }

        add_next_index_zval(zv_result_t, &_zv_tmp_row);
    }
}
/* }}} */

void skip_rows(xlsxioreadersheet sheet_t, zval *zv_type_t, zend_long data_type_default, zend_long zl_skip_row)
{
    while (sheet_read_row(sheet_t))
    {
        zval _zv_tmp_row;
        ZVAL_NULL(&_zv_tmp_row);

        if (xlsxioread_sheet_last_row_index(sheet_t) < zl_skip_row) {
            sheet_read_row(sheet_t);
        }

        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, zv_type_t, data_type_default, READ_SKIP_ROW);

        zval_ptr_dtor(&_zv_tmp_row);

        if (xlsxioread_sheet_last_row_index(sheet_t) >= zl_skip_row) {
            break;
        }
    }
}
