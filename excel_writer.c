/*
  +----------------------------------------------------------------------+
  | Vtiful Extension                                                     |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2017 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.vtiful.com                                                |
  +----------------------------------------------------------------------+
  | Author: viest <dev@service.viest.me>                                 |
  +----------------------------------------------------------------------+
*/

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "php.h"
#include "ext/standard/info.h"
#include "kernel/include.h"

excel_resource_t *excel_res;

int le_excel_writer;

/* {{{ PHP_MINIT_FUNCTION
 */
PHP_MINIT_FUNCTION(excel_writer)
{
    VTIFUL_STARTUP_MODULE(exception);
	VTIFUL_STARTUP_MODULE(excel);
	VTIFUL_STARTUP_MODULE(format);

	excel_res = malloc(sizeof(excel_resource_t));

    le_excel_writer = zend_register_list_destructors_ex(_php_vtiful_excel_close, NULL, VTIFUL_RESOURCE_NAME, module_number);

	return SUCCESS;
}
/* }}} */


/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_MSHUTDOWN_FUNCTION(excel_writer)
{
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_RINIT_FUNCTION
 */
PHP_RINIT_FUNCTION(excel_writer)
{
#if defined(COMPILE_DL_VTIFUL) && defined(ZTS)
	ZEND_TSRMLS_CACHE_UPDATE();
#endif
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_RSHUTDOWN_FUNCTION
 */
PHP_RSHUTDOWN_FUNCTION(excel_writer)
{
	return SUCCESS;
}
/* }}} */


/* {{{ PHP_MINFO_FUNCTION
 */
PHP_MINFO_FUNCTION(excel_writer)
{
	php_info_print_table_start();
	php_info_print_table_header(2, "excel_writer support", "enabled");
#if defined(PHP_VTIFUL_VERSION)
    php_info_print_table_row(2, "Version", PHP_VTIFUL_VERSION);
#endif
	php_info_print_table_end();
}
/* }}} */

/* {{{ vtiful_functions[]
 *
 * Every user visible function must have an entry in vtiful_functions[].
 */
const zend_function_entry excel_writer_functions[] = {
	PHP_FE_END
};
/* }}} */

/* {{{ vtiful_module_entry
 */
zend_module_entry excel_writer_module_entry = {
	STANDARD_MODULE_HEADER,
	"excel_writer",
    excel_writer_functions,
	PHP_MINIT(excel_writer),
	PHP_MSHUTDOWN(excel_writer),
	PHP_RINIT(excel_writer),		/* Replace with NULL if there's nothing to do at request start */
	PHP_RSHUTDOWN(excel_writer),	/* Replace with NULL if there's nothing to do at request end */
	PHP_MINFO(excel_writer),
    PHP_EXCEL_WRITER_VERSION,
	STANDARD_MODULE_PROPERTIES
};
/* }}} */

#ifdef COMPILE_DL_EXCEL_WRITER
#ifdef ZTS
ZEND_TSRMLS_CACHE_DEFINE();
#endif
ZEND_GET_MODULE(excel_writer)
#endif

/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
