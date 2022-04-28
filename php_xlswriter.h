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

#ifndef PHP_VTIFUL_H
#define PHP_VTIFUL_H

#include "xlswriter.h"

extern zend_module_entry xlswriter_module_entry;
#define phpext_xlswriter_ptr &xlswriter_module_entry

#define PHP_XLSWRITER_VERSION "1.5.2"
#define PHP_XLSWRITER_AUTHOR  "Jiexing.Wang (wjx@php.net)"

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

#if PHP_VERSION_ID >= 80000
#define TSRMLS_D	void
#define TSRMLS_DC
#define TSRMLS_C
#define TSRMLS_CC
#endif

#define VTIFUL_RESOURCE_NAME "xlsx"

extern int le_xls_writer;

#define VTIFUL_STARTUP_MODULE(module) ZEND_MODULE_STARTUP_N(xlsxwriter_##module)(INIT_FUNC_ARGS_PASSTHRU)
#define VTIFUL_STARTUP_FUNCTION(module) ZEND_MINIT_FUNCTION(xlsxwriter_##module)

void _php_vtiful_xls_close(zend_resource *rsrc TSRMLS_DC);

#if defined(ZTS) && defined(COMPILE_DL_VTIFUL)
ZEND_TSRMLS_CACHE_EXTERN();
#endif

PHP_MINIT_FUNCTION(xlswriter);
PHP_MSHUTDOWN_FUNCTION(xlswriter);
PHP_RINIT_FUNCTION(xlswriter);
PHP_RSHUTDOWN_FUNCTION(xlswriter);
PHP_MINFO_FUNCTION(xlswriter);

#endif


/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
