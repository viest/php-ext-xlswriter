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
#include "ext/standard/php_math.h"
#include "ext/standard/php_filestat.h"

/* {{{ */
zend_long date_double_to_timestamp(double value) {
    double days, partDay, hours, minutes, seconds;

    days    = floor(value);
    partDay = value - days;
    hours   = floor(partDay * 24);
    partDay = partDay * 24 - hours;
    minutes = floor(partDay * 60);
    partDay = partDay * 60 - minutes;
    seconds = _php_math_round(partDay * 60, 0, PHP_ROUND_HALF_UP);

    zval datetime;
    php_date_instantiate(php_date_get_date_ce(), &datetime);
    php_date_initialize(Z_PHPDATE_P(&datetime), ZEND_STRL("1899-12-30"), NULL, NULL, 1);

    zval _modify_args[1], _modify_result;
    smart_str _modify_arg_string = {0};
    if (days >= 0) {
        smart_str_appendl(&_modify_arg_string, "+", 1);
    }
    smart_str_append_long(&_modify_arg_string, days);
    smart_str_appendl(&_modify_arg_string, " days", 5);
    ZSTR_VAL(_modify_arg_string.s)[ZSTR_LEN(_modify_arg_string.s)] = '\0';
    ZVAL_STR(&_modify_args[0], _modify_arg_string.s);
    call_object_method(&datetime, "modify", 1, _modify_args, &_modify_result);
    zval_ptr_dtor(&datetime);

    zval _set_time_args[3], _set_time_result;
    ZVAL_LONG(&_set_time_args[0], (zend_long)hours);
    ZVAL_LONG(&_set_time_args[1], (zend_long)minutes);
    ZVAL_LONG(&_set_time_args[2], (zend_long)seconds);
    call_object_method(&_modify_result, "setTime", 3, _set_time_args, &_set_time_result);
    zval_ptr_dtor(&_modify_result);

    zval _format_args[1], _format_result;
    ZVAL_STRING(&_format_args[0], "U");
    call_object_method(&_set_time_result, "format", 1, _format_args, &_format_result);
    zval_ptr_dtor(&_set_time_result);

    zend_long timestamp = ZEND_STRTOL(Z_STRVAL(_format_result), NULL ,10);
    zval_ptr_dtor(&_format_result);

    return timestamp;
}
/* }}} */

/* {{{ */
unsigned int directory_exists(const char *path) {
    zval dir_exists;

#if PHP_VERSION_ID >= 80100
    zend_string *zs_path = zend_string_init(path, strlen(path), 0);
    php_stat(zs_path, FS_IS_DIR, &dir_exists);
    zend_string_release(zs_path);
#else
    php_stat(path, strlen(path), FS_IS_DIR, &dir_exists);
#endif

    if (Z_TYPE(dir_exists) == IS_FALSE) {
        zval_ptr_dtor(&dir_exists);
        return XLSWRITER_FALSE;
    }

    zval_ptr_dtor(&dir_exists);
    return XLSWRITER_TRUE;
}
/* }}} */

/* {{{ */
unsigned int file_exists(const char *path) {
    zval file_exists;

#if PHP_VERSION_ID >= 80100
    zend_string *zs_path = zend_string_init(path, strlen(path), 0);
    php_stat(zs_path, FS_IS_FILE, &file_exists);
    zend_string_release(zs_path);
#else
    php_stat(path, strlen(path), FS_IS_FILE, &file_exists);
#endif

    if (Z_TYPE(file_exists) == IS_FALSE) {
        zval_ptr_dtor(&file_exists);
        return XLSWRITER_FALSE;
    }

    zval_ptr_dtor(&file_exists);
    return XLSWRITER_TRUE;
}
/* }}} */