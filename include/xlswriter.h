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

#ifndef PHP_XLS_WRITER_INCLUDE_H
#define PHP_XLS_WRITER_INCLUDE_H

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <php.h>
#include "ext/date/php_date.h"

#include "zend_smart_str.h"
#include "zend_exceptions.h"
#include "zend.h"
#include "zend_API.h"
#include "php.h"

#include "xlsxwriter.h"
#include "xlsxwriter/packager.h"
#include "xlsxwriter/format.h"

#include "php_xlswriter.h"
#include "excel.h"
#include "validation.h"
#include "exception.h"
#include "format.h"
#include "chart.h"
#include "rich_string.h"
#include "help.h"

#ifdef ENABLE_READER
#include "xlsxio_read.h"
#include "read.h"
#include "csv.h"

typedef struct {
    xlsxioreader      file_t;
    xlsxioreadersheet sheet_t;
    zend_long         data_type_default;
    zend_long         sheet_flag;
} xls_resource_read_t;

typedef struct {
    zend_long             data_type_default;
    zval                  *zv_type_t;
    zend_fcall_info       *fci;
    zend_fcall_info_cache *fci_cache;
} xls_read_callback_data;
#endif

#ifndef ENABLE_READER
typedef struct {
    void      *file_t;
    void      *sheet_t;
    zend_long data_type_default;
    zend_long sheet_flag;
} xls_resource_read_t;
#endif

enum xlswriter_boolean {
    XLSWRITER_FALSE,
    XLSWRITER_TRUE
};

enum xlswirter_printed_direction {
    XLSWRITER_PRINTED_LANDSCAPE,
    XLSWRITER_PRINTED_PORTRAIT,
};

typedef struct {
    lxw_workbook  *workbook;
    lxw_worksheet *worksheet;
} xls_resource_write_t;

typedef struct {
    lxw_format  *format;
} xls_resource_format_t;

typedef struct {
    lxw_data_validation *validation;
} xls_resource_validation_t;

typedef struct {
    lxw_chart *chart;
    lxw_chart_series *series;
} xls_resource_chart_t;

typedef struct {
    lxw_rich_string_tuple *tuple;
} xls_resource_rich_string_t;

typedef struct _vtiful_xls_object {
    xls_resource_read_t   read_ptr;
    xls_resource_write_t  write_ptr;
    zend_long             write_line;
    xls_resource_format_t format_ptr;
    zend_object           zo;
} xls_object;

typedef struct _vtiful_format_object {
    xls_resource_format_t ptr;
    zend_object zo;
} format_object;

typedef struct _vtiful_chart_object {
    xls_resource_chart_t ptr;
    zend_object zo;
} chart_object;

typedef struct _vtiful_validation_object {
    xls_resource_validation_t ptr;
    zend_object zo;
} validation_object;

typedef struct _vtiful_rich_string_object {
    xls_resource_rich_string_t ptr;
    zend_object zo;
} rich_string_object;

#define REGISTER_CLASS_CONST_LONG(class_name, const_name, value) \
    zend_declare_class_constant_long(class_name, const_name, sizeof(const_name)-1, (zend_long)value);

#define REGISTER_CLASS_PROPERTY_NULL(class_name, property_name, acc) \
    zend_declare_property_null(class_name, ZEND_STRL(property_name), acc);

#define Z_XLS_P(zv)         php_vtiful_xls_fetch_object(Z_OBJ_P(zv));
#define Z_CHART_P(zv)       php_vtiful_chart_fetch_object(Z_OBJ_P(zv));
#define Z_FORMAT_P(zv)      php_vtiful_format_fetch_object(Z_OBJ_P(zv));
#define Z_VALIDATION_P(zv)  php_vtiful_validation_fetch_object(Z_OBJ_P(zv));
#define Z_RICH_STR_P(zv)    php_vtiful_rich_string_fetch_object(Z_OBJ_P(zv));

#define WORKBOOK_NOT_INITIALIZED(xls_object_t)                                                                       \
    do {                                                                                                             \
        if(xls_object_t->write_ptr.workbook == NULL) {                                                               \
            zend_throw_exception(vtiful_exception_ce, "Please create a file first, use the filename method", 130);   \
            return;                                                                                                  \
        }                                                                                                            \
    } while(0);

#define WORKSHEET_NOT_INITIALIZED(xls_object_t)                                                                     \
    do {                                                                                                            \
        if (xls_object_t->write_ptr.worksheet == NULL) {                                                            \
            zend_throw_exception(vtiful_exception_ce, "worksheet not initialized", 200);                            \
            return;                                                                                                 \
        }                                                                                                           \
    } while(0);

#define WORKSHEET_INDEX_OUT_OF_CHANGE_IN_OPTIMIZE_EXCEPTION(xls_resource_write_t, error)                                \
    do {                                                                                                                \
        if(xls_resource_write_t->worksheet->optimize && error == LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE) {              \
            zend_throw_exception(vtiful_exception_ce, "In const memory mode, you cannot modify the placed cells", 170); \
            return;                                                                                                     \
        }                                                                                                               \
    } while(0);

#define WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)                                                    \
    do {                                                                                                  \
        if(error == LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE) {                                             \
            zend_throw_exception(vtiful_exception_ce, "Worksheet row or column index out of range", 180); \
            return;                                                                                       \
        }                                                                                                 \
    } while(0);

#define WORKSHEET_WRITER_EXCEPTION(error)                                                   \
    do {                                                                                    \
        if(error > LXW_NO_ERROR) {                                                          \
            zend_throw_exception(vtiful_exception_ce, exception_message_map(error), error); \
            return;                                                                         \
        }                                                                                   \
    } while(0)

#define FCALL_TWO_ARGS(bucket)                   \
    ZVAL_COPY_VALUE(&args[0], &bucket->val); \
        if (bucket->key) {                       \
            ZVAL_STR(&args[1], bucket->key);     \
        } else {                                 \
            ZVAL_LONG(&args[1], bucket->h);      \
        }                                        \
        zend_call_function(&fci, &fci_cache);

#define ROW(range) \
    lxw_name_to_row(range)

#define ROWS(range) \
    lxw_name_to_row(range), lxw_name_to_row_2(range)

#define SHEET_LINE_INIT(obj_p) \
    obj_p->write_line = 0;

#define SHEET_LINE_ADD(obj_p) \
    ++obj_p->write_line;

#define SHEET_LINE_SET(obj_p, current_line) \
    obj_p->write_line = current_line;

#define SHEET_CURRENT_LINE(obj_p) obj_p->write_line

#ifdef LXW_HAS_SNPRINTF
#define lxw_snprintf snprintf
#else
#define lxw_snprintf __builtin_snprintf
#endif

#if PHP_VERSION_ID < 80000
#define PROP_OBJ(zv) (zv)
#else
#define PROP_OBJ(zv) Z_OBJ_P(zv)
#endif

#if PHP_VERSION_ID < 80000
#define Z_PARAM_STRING_OR_NULL(dest, dest_len) \
	Z_PARAM_STRING_EX(dest, dest_len, 1, 0)
#define Z_PARAM_STR_OR_NULL(dest) \
	Z_PARAM_STR_EX(dest, 1, 0)
#define Z_PARAM_RESOURCE_OR_NULL(dest) \
	Z_PARAM_RESOURCE_EX(dest, 1, 0)
#define Z_PARAM_DOUBLE_OR_NULL(dest, is_null) \
	Z_PARAM_DOUBLE_EX(dest, is_null, 1, 0)
#define Z_PARAM_LONG_OR_NULL(dest, is_null) \
	Z_PARAM_LONG_EX(dest, is_null, 1, 0)
#define Z_PARAM_ARRAY_OR_NULL(dest) \
	Z_PARAM_ARRAY_EX(dest, 1, 0)
#define Z_PARAM_BOOL_OR_NULL(dest, is_null) \
	Z_PARAM_BOOL_EX(dest, is_null, 1, 0)
#endif

static inline xls_object *php_vtiful_xls_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (xls_object *)((char *)(obj) - XtOffsetOf(xls_object, zo));
}

static inline format_object *php_vtiful_format_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (format_object *)((char *)(obj) - XtOffsetOf(format_object, zo));
}

static inline chart_object *php_vtiful_chart_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (chart_object *)((char *)(obj) - XtOffsetOf(chart_object, zo));
}

static inline validation_object *php_vtiful_validation_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (validation_object *)((char *)(obj) - XtOffsetOf(validation_object, zo));
}

static inline rich_string_object *php_vtiful_rich_string_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (rich_string_object *)((char *)(obj) - XtOffsetOf(validation_object, zo));
}

static inline void php_vtiful_close_resource(zend_object *obj) {
    if (obj == NULL) {
        return;
    }

    xls_object *intern = php_vtiful_xls_fetch_object(obj);

    SHEET_LINE_INIT(intern);

    if (intern->write_ptr.workbook != NULL) {
        lxw_workbook_free(intern->write_ptr.workbook);
        intern->write_ptr.workbook = NULL;
    }

    if (intern->format_ptr.format != NULL) {
        intern->format_ptr.format = NULL;
    }

#ifdef ENABLE_READER
    if (intern->read_ptr.sheet_t != NULL) {
        xlsxioread_sheet_close(intern->read_ptr.sheet_t);
        intern->read_ptr.sheet_t = NULL;
    }

    if (intern->read_ptr.file_t != NULL) {
        xlsxioread_close(intern->read_ptr.file_t);
        intern->read_ptr.file_t = NULL;
    }
#endif

    intern->read_ptr.data_type_default = READ_TYPE_EMPTY;
}

lxw_format            * zval_get_format(zval *handle);
lxw_data_validation   * zval_get_validation(zval *resource);
lxw_rich_string_tuple * zval_get_rich_string(zval *resource);
xls_resource_write_t  * zval_get_resource(zval *handle);
xls_resource_chart_t  * zval_get_chart(zval *resource);

STATIC lxw_error _store_defined_name(lxw_workbook *self, const char *name, const char *app_name, const char *formula, int16_t index, uint8_t hidden);

STATIC int  _compare_defined_names(lxw_defined_name *a, lxw_defined_name *b);

STATIC void _prepare_drawings(lxw_workbook *self);
STATIC void _add_chart_cache_data(lxw_workbook *self);
STATIC void _prepare_vml(lxw_workbook *self);
STATIC void _prepare_defined_names(lxw_workbook *self);
STATIC void _populate_range(lxw_workbook *self, lxw_series_range *range);
STATIC void _populate_range_dimensions(lxw_workbook *self, lxw_series_range *range);

void comment_show(xls_resource_write_t *res);
void hide_worksheet(xls_resource_write_t *res);
void first_worksheet(xls_resource_write_t *res);
void zoom(xls_resource_write_t *res, zend_long zoom);
void paper(xls_resource_write_t *res, zend_long type);
void gridlines(xls_resource_write_t *res, zend_long option);
void auto_filter(zend_string *range, xls_resource_write_t *res);
void protection(xls_resource_write_t *res, zend_string *password);
void format_copy(lxw_format *new_format, lxw_format *other_format);
void printed_direction(xls_resource_write_t *res, unsigned int direction);
void xls_file_path(zend_string *file_name, zval *dir_path, zval *file_path);
void freeze_panes(xls_resource_write_t *res, zend_long row, zend_long column);
void margins(xls_resource_write_t *res, double left, double right, double top, double bottom);
void set_row(zend_string *range, double height, xls_resource_write_t *res, lxw_format *format);
void validation(xls_resource_write_t *res, zend_string *range, lxw_data_validation *validation);
void set_column(zend_string *range, double width, xls_resource_write_t *res, lxw_format *format);
void merge_cells(zend_string *range, zval *value, xls_resource_write_t *res, lxw_format *format);
void comment_writer(zend_string *comment, zend_long row, zend_long columns, xls_resource_write_t *res);
void call_object_method(zval *object, const char *function_name, uint32_t param_count, zval *params, zval *ret_val);
void chart_writer(zend_long row, zend_long columns, xls_resource_chart_t *chart_resource, xls_resource_write_t *res);
void worksheet_set_rows(lxw_row_t start, lxw_row_t end, double height, xls_resource_write_t *res, lxw_format *format);
void image_writer(zval *value, zend_long row, zend_long columns, double width, double height, xls_resource_write_t *res);
void formula_writer(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxw_format *format);
void type_writer(zval *value, zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *format, lxw_format *format_handle);
void rich_string_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zval *rich_strings, lxw_format *format);
void datetime_writer(lxw_datetime *datetime, zend_long row, zend_long columns, zend_string *format, xls_resource_write_t *res, lxw_format *format_handle);
void url_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *url, zend_string *text, zend_string *tool_tip, lxw_format *format);

lxw_error workbook_file(xls_resource_write_t *self);

lxw_datetime timestamp_to_datetime(zend_long timestamp);
zend_string* char_join_to_zend_str(const char *left, const char *right);
zend_string* str_pick_up(zend_string *left, const char *right, size_t len);

#endif
