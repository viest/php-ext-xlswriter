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
void xls_file_path(zend_string *file_name, zval *dir_path, zval *file_path)
{
    zend_string *full_path, *zstr_path;

    zstr_path = zval_get_string(dir_path);

    if (Z_STRVAL_P(dir_path)[Z_STRLEN_P(dir_path)-1] == '/') {
        full_path = zend_string_extend(zstr_path, ZSTR_LEN(zstr_path) + ZSTR_LEN(file_name), 0);
        memcpy(ZSTR_VAL(full_path)+ZSTR_LEN(zstr_path), ZSTR_VAL(file_name), ZSTR_LEN(file_name)+1);
    } else {
        full_path = zend_string_extend(zstr_path, ZSTR_LEN(zstr_path) + ZSTR_LEN(file_name) + 1, 0);
        ZSTR_VAL(full_path)[ZSTR_LEN(zstr_path)] ='/';
        memcpy(ZSTR_VAL(full_path)+ZSTR_LEN(zstr_path)+1, ZSTR_VAL(file_name), ZSTR_LEN(file_name)+1);
    }

    ZVAL_STR(file_path, full_path);
}
/* }}} */

/* {{{ */
zend_string* str_pick_up(zend_string *left, char *right)
{
    zend_string *full = NULL;

    size_t _left_length = ZSTR_LEN(left);
    size_t _extend_length = _left_length + strlen(right);

    full = zend_string_extend(left, _extend_length, 0);

    memcpy(ZSTR_VAL(full) + _left_length, right, strlen(right));

    ZSTR_VAL(full)[_extend_length] = '\0';

    return full;
}
/* }}} */