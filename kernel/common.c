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

    /* Guard the [len-1] index: Z_STRLEN is size_t, so an empty path would
     * underflow to SIZE_MAX and read out of bounds. An empty dir falls through
     * to the else-branch (which prepends '/'). */
    if (Z_STRLEN_P(dir_path) > 0 && Z_STRVAL_P(dir_path)[Z_STRLEN_P(dir_path)-1] == '/') {
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
lxlsx_datetime timestamp_to_datetime(zend_long timestamp)
{
    int yearLocal   = php_idate('Y', timestamp, 0);
    int monthLocal  = php_idate('m', timestamp, 0);
    int dayLocal    = php_idate('d', timestamp, 0);
    int hourLocal   = php_idate('H', timestamp, 0);
    int minuteLocal = php_idate('i', timestamp, 0);
    int secondLocal = php_idate('s', timestamp, 0);

    lxlsx_datetime datetime = {
            yearLocal, monthLocal, dayLocal, hourLocal, minuteLocal, secondLocal
    };

    return datetime;
}
/* }}} */

/* {{{ */
lxlsx_format* object_format(xls_object *obj, zend_string *format, lxlsx_format *lxlsx_format_handle)
{
    if (format == NULL && lxlsx_format_handle == NULL) {
        return NULL;
    }

    /* The cache keys below MUST use ZSTR_VAL+ZSTR_LEN, not ZEND_STRL(x->val).
     * ZEND_STRL is for string literals — sizeof("...") - 1 — but applied to
     * zend_string's flexible-array `val` it expands to sizeof(char[1]) - 1
     * which is 0. That collapses every distinct format string into the same
     * empty-string cache slot, and the first format wins for every
     * subsequent insertDate / insertText call (reported as #552 / #548 /
     * #544: "$format ignored; everything comes back as the first format"). */
    if (format != NULL && lxlsx_format_handle != NULL) {
        if (format->len <= 0) {
            return lxlsx_format_handle;
        }

        zend_string *_format_key = strpprintf(0, "%p|%s", lxlsx_format_handle, ZSTR_VAL(format));

        void *exit_format = zend_hash_str_find_ptr(obj->formats_cache_ptr.maps, ZSTR_VAL(_format_key), ZSTR_LEN(_format_key));

        if (exit_format != NULL) {
            zend_string_release(_format_key);

            return (lxlsx_format *)exit_format;
        }

        lxlsx_format *new_format = lxlsx_workbook_add_format((&obj->write_ptr)->workbook);
        lxlsx_format_copy(new_format, lxlsx_format_handle);
        lxlsx_format_set_num_format(new_format, ZSTR_VAL(format));

        zend_hash_str_add_ptr(obj->formats_cache_ptr.maps, ZSTR_VAL(_format_key), ZSTR_LEN(_format_key), new_format);

        zend_string_release(_format_key);

        return new_format;
    }

    if (format != NULL) {
        if (format->len <= 0) {
            return NULL;
        }

        void *exit_format = zend_hash_str_find_ptr(obj->formats_cache_ptr.maps, ZSTR_VAL(format), ZSTR_LEN(format));

        if (exit_format != NULL) {
            return (lxlsx_format *)exit_format;
        }

        lxlsx_format *new_format = lxlsx_workbook_add_format((&obj->write_ptr)->workbook);
        lxlsx_format_set_num_format(new_format, ZSTR_VAL(format));

        zend_hash_str_add_ptr(obj->formats_cache_ptr.maps, ZSTR_VAL(format), ZSTR_LEN(format), new_format);

        return new_format;
    }

    return lxlsx_format_handle;
}
/* }}} */