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

#ifndef VTIFUL_EXCEL_WRITE_H
#define VTIFUL_EXCEL_WRITE_H

void type_writer(zval *value, zend_long row, zend_long columns, excel_resource_t *res);
lxw_error workbook_file(excel_resource_t *self, zval *handle);

#endif
