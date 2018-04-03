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

#ifndef PHP_EXT_EXCEL_EXPORT_FORMAT_H
#define PHP_EXT_EXCEL_EXPORT_FORMAT_H

#include "php_excel_writer.h"
#include "xlsxwriter.h"

extern zend_class_entry *vtiful_format_ce;

VTIFUL_STARTUP_FUNCTION(format);

#endif
