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

#define PHP_XLSWRITER_VERSION "1.5.8"
#define PHP_XLSWRITER_AUTHOR  "Jiexing.Wang (wjx@php.net)"

extern int le_xls_writer;

void _php_vtiful_xls_close(zend_resource *rsrc TSRMLS_DC);

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
