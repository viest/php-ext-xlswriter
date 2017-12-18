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

#ifndef VTIFUL_EXCEL_WRITE_H
#define VTIFUL_EXCEL_WRITE_H

STATIC lxw_error _store_defined_name(lxw_workbook *self, const char *name, const char *app_name, const char *formula, int16_t index, uint8_t hidden);

STATIC void _prepare_defined_names(lxw_workbook *self);
STATIC void _prepare_drawings(lxw_workbook *self);
STATIC void _add_chart_cache_data(lxw_workbook *self);
STATIC int  _compare_defined_names(lxw_defined_name *a, lxw_defined_name *b);
STATIC void _populate_range(lxw_workbook *self, lxw_series_range *range);
STATIC void _populate_range_dimensions(lxw_workbook *self, lxw_series_range *range);

void type_writer(zval *value, zend_long row, zend_long columns, excel_resource_t *res, zend_string *format);
void image_writer(zval *value, zend_long row, zend_long columns, excel_resource_t *res);
void formula_writer(zval *value, zend_long row, zend_long columns, excel_resource_t *res);
void auto_filter(zend_string *range, excel_resource_t *res);
void merge_cells(zend_string *range, zend_string *value, excel_resource_t *res);
lxw_error workbook_file(excel_resource_t *self, zval *handle);

#endif
