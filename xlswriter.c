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

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "php.h"
#include "ext/standard/info.h"
#include "xlswriter.h"

#if ENABLE_READER
#include <xlsxio_version.h>
#include <xlsxio_read.h>
#endif

int le_xls_writer;

ZEND_BEGIN_ARG_INFO_EX(xlswriter_get_version_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xlswriter_get_auther_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

/* {{{ xlswriter_get_version
 */
PHP_FUNCTION(xlswriter_get_version)
{
    RETURN_STRINGL(PHP_XLSWRITER_VERSION, strlen(PHP_XLSWRITER_VERSION));
}
/* }}} */

/* {{{ xlswriter_get_author
 */
PHP_FUNCTION(xlswriter_get_author)
{
    RETURN_STRINGL(PHP_XLSWRITER_AUTHOR, strlen(PHP_XLSWRITER_AUTHOR));
}
/* }}} */

/* {{{ PHP_MINIT_FUNCTION
 */
PHP_MINIT_FUNCTION(xlswriter)
{
    VTIFUL_STARTUP_MODULE(exception);
	VTIFUL_STARTUP_MODULE(excel);
	VTIFUL_STARTUP_MODULE(format);
	VTIFUL_STARTUP_MODULE(chart);
    VTIFUL_STARTUP_MODULE(validation);
    VTIFUL_STARTUP_MODULE(rich_string);

	le_xls_writer = zend_register_list_destructors_ex(_php_vtiful_xls_close, NULL, VTIFUL_RESOURCE_NAME, module_number);

	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_MSHUTDOWN_FUNCTION(xlswriter)
{
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_RINIT_FUNCTION
 */
PHP_RINIT_FUNCTION(xlswriter)
{
#if defined(COMPILE_DL_VTIFUL) && defined(ZTS)
	ZEND_TSRMLS_CACHE_UPDATE();
#endif
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_RSHUTDOWN_FUNCTION
 */
PHP_RSHUTDOWN_FUNCTION(xlswriter)
{
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MINFO_FUNCTION
 */
PHP_MINFO_FUNCTION(xlswriter)
{
	php_info_print_table_start();
	php_info_print_table_header(2, "xlswriter support", "enabled");
#if defined(PHP_XLSWRITER_VERSION)
    php_info_print_table_row(2, "Version", PHP_XLSWRITER_VERSION);
#endif
#ifdef LXW_VERSION
#ifdef HAVE_LIBXLSXWRITER
    /* Build time */
    php_info_print_table_row(2, "libxlsxwriter headers version", LXW_VERSION);
    /* Run time, available since 0.7.9 */
    php_info_print_table_row(2, "libxlsxwriter library version", lxw_version());
#else
    php_info_print_table_row(2, "bundled libxlsxwriter version", LXW_VERSION);
#endif
#endif

#if ENABLE_READER
#if HAVE_LIBXLSXIO
    /* Build time */
    php_info_print_table_row(2, "libxlsxio headers version", XLSXIO_VERSION_STRING);
    /* Run time */
    php_info_print_table_row(2, "libxlsxio library version", xlsxioread_get_version_string());
#else
    php_info_print_table_row(2, "bundled libxlsxio version", XLSXIO_VERSION_STRING);
#endif
#endif

	php_info_print_table_end();
}
/* }}} */

/* {{{ xlswriter_functions[]
 *
 * Every user visible function must have an entry in xlswriter_functions[].
 */
const zend_function_entry xlswriter_functions[] = {
    PHP_FE(xlswriter_get_version, xlswriter_get_version_arginfo)
    PHP_FE(xlswriter_get_author,  xlswriter_get_auther_arginfo)
	PHP_FE_END
};
/* }}} */

/* {{{ vtiful_module_entry
 */
zend_module_entry xlswriter_module_entry = {
	STANDARD_MODULE_HEADER,
	"xlswriter",
    xlswriter_functions,
	PHP_MINIT(xlswriter),
	PHP_MSHUTDOWN(xlswriter),
	PHP_RINIT(xlswriter),		/* Replace with NULL if there's nothing to do at request start */
	PHP_RSHUTDOWN(xlswriter),	/* Replace with NULL if there's nothing to do at request end */
	PHP_MINFO(xlswriter),
	PHP_XLSWRITER_VERSION,
	STANDARD_MODULE_PROPERTIES
};
/* }}} */

#ifdef COMPILE_DL_XLSWRITER
#ifdef ZTS
ZEND_TSRMLS_CACHE_DEFINE();
#endif
ZEND_GET_MODULE(xlswriter)
#endif

/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
