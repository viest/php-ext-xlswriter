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
xlsxioreadersheet sheet_open(xlsxioreader file, const zend_string *zs_sheet_name_t)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_sheet_open(file, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
    }

    return xlsxioread_sheet_open(file, ZSTR_VAL(zs_sheet_name_t), XLSXIOREAD_SKIP_EMPTY_ROWS);
}
/* }}} */

/* {{{ */
int sheet_read_row(xlsxioreadersheet sheet)
{
    return xlsxioread_sheet_next_row(sheet);
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