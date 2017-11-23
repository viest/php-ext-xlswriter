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
#include "php_ini.h"
#include "php_vtiful.h"
#include "ext/standard/info.h"

#include "kernel/excel.h"
#include "kernel/exception.h"

int le_vtiful;

PHP_MINIT_FUNCTION(vtiful)
{
	/* If you have INI entries, uncomment these lines
	REGISTER_INI_ENTRIES();
	*/

    VTIFUL_STARTUP_MODULE(vtiful_exception);
	VTIFUL_STARTUP_MODULE(excel);

	le_vtiful = zend_register_list_destructors_ex(_php_vtiful_excel_close, NULL, "vtiful", module_number);

	return SUCCESS;
}
/* }}} */


/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_MSHUTDOWN_FUNCTION(vtiful)
{
	/* uncomment this line if you have INI entries
	UNREGISTER_INI_ENTRIES();
	*/
	return SUCCESS;
}
/* }}} */


/* Remove if there's nothing to do at request start */
/* {{{ PHP_RINIT_FUNCTION
 */
PHP_RINIT_FUNCTION(vtiful)
{
#if defined(COMPILE_DL_VTIFUL) && defined(ZTS)
	ZEND_TSRMLS_CACHE_UPDATE();
#endif
	return SUCCESS;
}
/* }}} */


/* Remove if there's nothing to do at request end */
/* {{{ PHP_RSHUTDOWN_FUNCTION
 */
PHP_RSHUTDOWN_FUNCTION(vtiful)
{
	return SUCCESS;
}
/* }}} */


/* {{{ PHP_MINFO_FUNCTION
 */
PHP_MINFO_FUNCTION(vtiful)
{
	php_info_print_table_start();
	php_info_print_table_header(2, "vtiful support", "enabled");
	php_info_print_table_end();

	/* Remove comments if you have entries in php.ini
	DISPLAY_INI_ENTRIES();
	*/
}
/* }}} */


/* {{{ _php_fss_close
 *
 * List destructor for FSS handles
 */
static void _php_vtiful_excel_close(zend_resource *rsrc TSRMLS_DC)
{
//	int i;
//	fss_resource_t * res = (fss_resource_t *)rsrc->ptr;
//	/* Destroy the replace strings */
//	for (i = 0; i < res->replace_size; i++) {
//		if (res->replace[i]) {
//			zval_ptr_dtor(res->replace[i]);
//		}
//	}
//	/* Destroy the kwset structure */
//	kwsfree(res->set);
//	/* Destroy the resource structure itself */
//	efree(res);
}
/* }}} */


/* {{{ vtiful_functions[]
 *
 * Every user visible function must have an entry in vtiful_functions[].
 */
const zend_function_entry vtiful_functions[] = {
	PHP_FE_END
};
/* }}} */

/* {{{ vtiful_module_entry
 */
zend_module_entry vtiful_module_entry = {
	STANDARD_MODULE_HEADER,
	"vtiful",
	vtiful_functions,
	PHP_MINIT(vtiful),
	PHP_MSHUTDOWN(vtiful),
	PHP_RINIT(vtiful),		/* Replace with NULL if there's nothing to do at request start */
	PHP_RSHUTDOWN(vtiful),	/* Replace with NULL if there's nothing to do at request end */
	PHP_MINFO(vtiful),
	PHP_VTIFUL_VERSION,
	STANDARD_MODULE_PROPERTIES
};
/* }}} */

#ifdef COMPILE_DL_VTIFUL
#ifdef ZTS
ZEND_TSRMLS_CACHE_DEFINE();
#endif
ZEND_GET_MODULE(vtiful)
#endif

/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
