/*
  +----------------------------------------------------------------------+
  | Vtiful Extension                                                     |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2017 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <dev@service.viest.me>                                 |
  +----------------------------------------------------------------------+
*/

#ifndef PHP_VTIFUL_H
#define PHP_VTIFUL_H

#include "kernel/include.h"

extern zend_module_entry excel_writer_module_entry;
#define phpext_excel_writer_ptr &excel_writer_module_entry

#define PHP_EXCEL_WRITER_VERSION "1.0.0"

#ifdef PHP_WIN32
#	define PHP_VTIFUL_API __declspec(dllexport)
#elif defined(__GNUC__) && __GNUC__ >= 4
#	define PHP_VTIFUL_API __attribute__ ((visibility("default")))
#else
#	define PHP_VTIFUL_API
#endif

#ifdef ZTS
#include "TSRM.h"
#endif

#define VTIFUL_RESOURCE_NAME "vtiful"

extern int le_excel_writer;

#define VTIFUL_STARTUP_MODULE(module) ZEND_MODULE_STARTUP_N(excel_writer_##module)(INIT_FUNC_ARGS_PASSTHRU)
#define VTIFUL_STARTUP_FUNCTION(module) ZEND_MINIT_FUNCTION(excel_writer_##module)
#define VTIFUL_G(v) ZEND_MODULE_GLOBALS_ACCESSOR(vtiful, v)

void _php_vtiful_excel_close(zend_resource *rsrc TSRMLS_DC);

#if defined(ZTS) && defined(COMPILE_DL_VTIFUL)
ZEND_TSRMLS_CACHE_EXTERN();
#endif

PHP_MINIT_FUNCTION(excel_writer);
PHP_MSHUTDOWN_FUNCTION(excel_writer);
PHP_RINIT_FUNCTION(excel_writer);
PHP_RSHUTDOWN_FUNCTION(excel_writer);
PHP_MINFO_FUNCTION(excel_writer);

#endif


/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
