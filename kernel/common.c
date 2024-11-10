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
zend_string* str_pick_up(zend_string *left, const char *right, size_t len)
{
    zend_string *full = NULL;

    size_t _left_length = ZSTR_LEN(left);
    size_t _extend_length = _left_length + len;

    full = zend_string_extend(left, _extend_length, 0);

    memcpy(ZSTR_VAL(full) + _left_length, right, len);

    ZSTR_VAL(full)[_extend_length] = '\0';

    return full;
}
/* }}} */

/* {{{ */
zend_string* char_join_to_zend_str(const char *left, const char *right)
{
    size_t _new_len = strlen(left) + strlen(right);

    zend_string *str = zend_string_alloc(_new_len, 0);

    memcpy(ZSTR_VAL(str), left, strlen(left));
    memcpy(ZSTR_VAL(str) + strlen(left), right, strlen(right) + 1);

    ZSTR_VAL(str)[_new_len] = '\0';

    return str;
}

/* }}} */

/* {{{ */
void call_object_method(zval *object, const char *function_name, uint32_t param_count, zval *params, zval *ret_val)
{
    uint32_t index;
    zval z_f_name;

    ZVAL_STRINGL(&z_f_name, function_name, strlen(function_name));
    call_user_function(NULL, object, &z_f_name, ret_val, param_count, params);

    if (Z_ISUNDEF_P(ret_val)) {
        ZVAL_NULL(ret_val);
    }

    for (index = 0; index < param_count; index++) {
        zval_ptr_dtor(&params[index]);
    }

    zval_ptr_dtor(&z_f_name);
}
/* }}} */

/* {{{ */
lxw_datetime timestamp_to_datetime(zend_long timestamp)
{
    int yearLocal   = php_idate('Y', timestamp, 0);
    int monthLocal  = php_idate('m', timestamp, 0);
    int dayLocal    = php_idate('d', timestamp, 0);
    int hourLocal   = php_idate('H', timestamp, 0);
    int minuteLocal = php_idate('i', timestamp, 0);
    int secondLocal = php_idate('s', timestamp, 0);

    lxw_datetime datetime = {
            yearLocal, monthLocal, dayLocal, hourLocal, minuteLocal, secondLocal
    };

    return datetime;
}
/* }}} */

/* {{{ */
lxw_format* object_format(xls_object *obj, zend_string *format, lxw_format *format_handle)
{
    if (format == NULL && format_handle == NULL) {
        return NULL;
    }

    if (format != NULL && format_handle != NULL) {
        if (format->len <= 0) {
            return format_handle;
        }

        zend_string *_format_key = strpprintf(0, "%p|%s", format_handle, format->val);

        void *exit_format = zend_hash_str_find_ptr(obj->formats_cache_ptr.maps, ZEND_STRL(_format_key->val));

        if (exit_format != NULL) {
            zend_string_release(_format_key);

            return (lxw_format *)exit_format;
        }

        lxw_format *new_format = workbook_add_format((&obj->write_ptr)->workbook);
        format_copy(new_format, format_handle);
        format_set_num_format(new_format, ZSTR_VAL(format));

        zend_hash_str_add_ptr(obj->formats_cache_ptr.maps, ZEND_STRL(_format_key->val), new_format);

        zend_string_release(_format_key);

        return new_format;
    }

    if (format != NULL) {
        if (format->len <= 0) {
            return NULL;
        }

        void *exit_format = zend_hash_str_find_ptr(obj->formats_cache_ptr.maps, ZEND_STRL(format->val));

        if (exit_format != NULL) {
            return (lxw_format *)exit_format;
        }

        lxw_format *new_format = workbook_add_format((&obj->write_ptr)->workbook);
        format_set_num_format(new_format, ZSTR_VAL(format));

        zend_hash_str_add_ptr(obj->formats_cache_ptr.maps, ZEND_STRL(format->val), new_format);

        return new_format;
    }

    return format_handle;
}
/* }}} */