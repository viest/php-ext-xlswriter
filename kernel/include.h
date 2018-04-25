#ifndef PHP_EXT_EXCEL_EXPORT_INCLUDE_H
#define PHP_EXT_EXCEL_EXPORT_INCLUDE_H

#include <php.h>

#include "zend_exceptions.h"
#include "zend.h"
#include "zend_API.h"
#include "php.h"

#include "xlsxwriter.h"
#include "xlsxwriter/packager.h"
#include "xlsxwriter/format.h"

#include "../php_excel_writer.h"
#include "excel.h"
#include "exception.h"
#include "format.h"

#define REGISTER_CLASS_CONST_LONG(class_name, const_name, value) \
    zend_declare_class_constant_long(class_name, const_name, sizeof(const_name)-1, (zend_long)value);

#define REGISTER_CLASS_PROPERTY_NULL(class_name, property_name, acc) \
    zend_declare_property_null(class_name, ZEND_STRL(property_name), acc);

#define ROW(range) \
    lxw_name_to_row(range)

typedef struct {
    lxw_workbook *workbook;
    lxw_worksheet *worksheet;
} excel_resource_t;

excel_resource_t * zval_get_resource(zval *handle);
lxw_format       * zval_get_format(zval *handle);

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
void set_column(zend_string *range, double width, excel_resource_t *res, lxw_format *format);
void set_row(zend_string *range, double height, excel_resource_t *res, lxw_format *format);
lxw_error workbook_file(excel_resource_t *self, zval *handle);

void excel_file_path(zend_string *file_name, zval *dir_path, zval *file_path);

#endif
