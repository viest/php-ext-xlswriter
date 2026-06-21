/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension                                                  |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2018 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <wjx@php.net>                                          |
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

#include "libxlsx.h"
#include "libxlsx/packager.h"
#include "libxlsx/format.h"
#include "libxlsx/formula.h"

#include "common.h"
#include "php_xlswriter.h"
#include "excel.h"
#include "validation.h"
#include "exception.h"
#include "format.h"
#include "chart.h"
#include "rich_string.h"
#include "conditional_format.h"
#include "table.h"
#include "help.h"

#include "read.h"
#include "csv.h"

typedef struct xls_resource_read_t {
    lxlsx_reader_workbook  *file_t;
    lxlsx_reader_worksheet *sheet_t;
    zend_long      data_type_default;
    zend_long      sheet_flag;

    /* Synthesised-blanks bookkeeping (matches libxlsxio's compatibility shape):
     *   cols              max column from first non-empty row (trailing fill)
     *   expected_row_nr   next row number we expect to emit (1-based)
     *   pending_synth_rows  empty rows still to yield from a row gap (nextRow)
     *   pending_real_row    real row data buffered while emitting synth rows
     */
    size_t         cols;
    size_t         expected_row_nr;
    size_t         pending_synth_rows;
    zval           pending_real_row;
} xls_resource_read_t;

typedef struct {
    zend_long             data_type_default;
    zval                  *zv_type_t;
    zend_fcall_info       *fci;
    zend_fcall_info_cache *fci_cache;
} xls_read_callback_data;

enum xlswriter_boolean {
    XLSWRITER_FALSE,
    XLSWRITER_TRUE
};

enum xlswirter_printed_direction {
    XLSWRITER_PRINTED_LANDSCAPE,
    XLSWRITER_PRINTED_PORTRAIT,
};

typedef struct {
    lxlsx_workbook  *workbook;
    lxlsx_worksheet *worksheet;
    /* Auto-size tracking: per-column maximum estimated display width,
     * accumulated during writes (only while auto_size_enabled) so the widths
     * can be applied before the worksheet is packaged. Lazily allocated on
     * first tracked write; reset on sheet switch. */
    double        *auto_widths;
    size_t         auto_widths_n;
    int            auto_size_enabled;
    lxlsx_col_t      auto_size_first_col;
    lxlsx_col_t      auto_size_last_col;
} xls_resource_write_t;

/* Auto-size helpers (forward declarations — defined in kernel/write.c). */
double xls_estimate_cell_width(zval *value);
void   xls_track_auto_width(xls_resource_write_t *res, lxlsx_col_t col, double width);
void   xls_auto_widths_reset(xls_resource_write_t *res);
void   xls_auto_widths_apply(xls_resource_write_t *res, lxlsx_col_t first_col, lxlsx_col_t last_col);
void   xls_auto_widths_flush(xls_resource_write_t *res);

typedef struct {
    lxlsx_format  *format;
} xls_resource_format_t;

typedef struct {
    HashTable *maps;
} xls_resource_formats_cache_t;

typedef struct {
    lxlsx_data_validation *validation;
} xls_resource_validation_t;

typedef struct {
    lxlsx_chart *chart;
    lxlsx_chart_series *series;
} xls_resource_chart_t;

typedef struct {
    lxlsx_rich_string_tuple *tuple;
    zend_string *text;
} xls_resource_rich_string_t;

typedef struct _vtiful_xls_object {
    xls_resource_read_t          read_ptr;
    xls_resource_write_t         write_ptr;
    zend_long                    write_line;
    xls_resource_format_t        lxlsx_format_ptr;
    xls_resource_formats_cache_t formats_cache_ptr;
    lxlsx_row_col_options          *row_options;
    uint8_t                      compute_formula;  /* auto-evaluate insertFormula */
    zend_object                  zo;
} xls_object;

typedef struct _vtiful_format_object {
    xls_resource_format_t ptr;
    zend_object zo;
} lxlsx_format_object;

typedef struct _vtiful_chart_object {
    xls_resource_chart_t ptr;
    zend_object zo;
} lxlsx_chart_object;

typedef struct _vtiful_validation_object {
    xls_resource_validation_t ptr;
    zend_object zo;
} validation_object;

typedef struct _vtiful_rich_string_object {
    xls_resource_rich_string_t ptr;
    zend_object zo;
} rich_string_object;

/* Phase 2: ConditionalFormat builder object.
 * Owns the lxlsx_conditional_format struct itself (pointers inside reference
 * zend_string buffers held alongside, so we can free them on dtor). */
typedef struct {
    lxlsx_conditional_format *cf;
    /* Heap copies of any string fields the user supplied. Freed in the dtor. */
    char *value_string;
    char *min_value_string;
    char *mid_value_string;
    char *max_value_string;
    char *multi_range;
} xls_resource_cond_format_t;

typedef struct _vtiful_cond_format_object {
    xls_resource_cond_format_t ptr;
    zend_object zo;
} cond_format_object;

/* Phase 2: Excel Table builder object. */
typedef struct {
    lxlsx_table_options *opts;
    /* Owned string buffers for opts->name and any per-column header /
     * formula / total_string. Freed on dtor. */
    char *name;
    /* Owned columns array (NULL-terminated). columns[i] is calloc'd. */
    lxlsx_table_column **columns;
} xls_resource_table_t;

typedef struct _vtiful_table_object {
    xls_resource_table_t ptr;
    zend_object zo;
} table_object;

#define REGISTER_CLASS_CONST_LONG(class_name, const_name, value) \
    zend_declare_class_constant_long(class_name, const_name, sizeof(const_name)-1, (zend_long)value);

#define REGISTER_CLASS_PROPERTY_NULL(class_name, property_name, acc) \
    zend_declare_property_null(class_name, ZEND_STRL(property_name), acc);

#define Z_XLS_P(zv)         php_vtiful_xls_fetch_object(Z_OBJ_P(zv));
#define Z_CHART_P(zv)       php_vtiful_chart_fetch_object(Z_OBJ_P(zv));
#define Z_FORMAT_P(zv)      php_vtiful_format_fetch_object(Z_OBJ_P(zv));
#define Z_VALIDATION_P(zv)  php_vtiful_validation_fetch_object(Z_OBJ_P(zv));
#define Z_RICH_STR_P(zv)    php_vtiful_rich_string_fetch_object(Z_OBJ_P(zv));
#define Z_COND_FMT_P(zv)    php_vtiful_cond_format_fetch_object(Z_OBJ_P(zv));
#define Z_TABLE_P(zv)       php_vtiful_table_fetch_object(Z_OBJ_P(zv));

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
        if(xls_resource_write_t->worksheet->optimize && error == LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE) {              \
            zend_throw_exception(vtiful_exception_ce, "In const memory mode, you cannot modify the placed cells", 170); \
            return;                                                                                                     \
        }                                                                                                               \
    } while(0);

#define WORKSHEET_INDEX_OUT_OF_CHANGE_EXCEPTION(error)                                                    \
    do {                                                                                                  \
        if(error == LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE) {                                             \
            zend_throw_exception(vtiful_exception_ce, "Worksheet row or column index out of range", 180); \
            return;                                                                                       \
        }                                                                                                 \
    } while(0);

#define WORKSHEET_WRITER_EXCEPTION(error)                                                   \
    do {                                                                                    \
        if(error > LXLSX_NO_ERROR) {                                                          \
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
    lxlsx_name_to_row(range)

#define ROWS(range) \
    lxlsx_name_to_row(range), lxlsx_name_to_row_2(range)

#define SHEET_LINE_INIT(obj_p) \
    obj_p->write_line = 0;

#define SHEET_LINE_ADD(obj_p) \
    ++obj_p->write_line;

#define SHEET_LINE_SET(obj_p, current_line) \
    obj_p->write_line = current_line;

#define SHEET_CURRENT_LINE(obj_p) obj_p->write_line

#ifdef LXLSX_HAS_SNPRINTF
#define lxlsx_snprintf snprintf
#else
#define lxlsx_snprintf __builtin_snprintf
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

    return (xls_object *)((char *)(obj) - offsetof(xls_object, zo));
}

static inline lxlsx_format_object *php_vtiful_format_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (lxlsx_format_object *)((char *)(obj) - offsetof(lxlsx_format_object, zo));
}

static inline lxlsx_chart_object *php_vtiful_chart_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (lxlsx_chart_object *)((char *)(obj) - offsetof(lxlsx_chart_object, zo));
}

static inline validation_object *php_vtiful_validation_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (validation_object *)((char *)(obj) - offsetof(validation_object, zo));
}

static inline rich_string_object *php_vtiful_rich_string_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }

    return (rich_string_object *)((char *)(obj) - offsetof(rich_string_object, zo));
}

static inline cond_format_object *php_vtiful_cond_format_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }
    return (cond_format_object *)((char *)(obj) - offsetof(cond_format_object, zo));
}

static inline table_object *php_vtiful_table_fetch_object(zend_object *obj) {
    if (obj == NULL) {
        return NULL;
    }
    return (table_object *)((char *)(obj) - offsetof(table_object, zo));
}

static inline void php_vtiful_reset_reader_state(xls_resource_read_t *read_ptr) {
    if (read_ptr == NULL) {
        return;
    }

    if (Z_TYPE(read_ptr->pending_real_row) == IS_ARRAY) {
        zval_ptr_dtor(&read_ptr->pending_real_row);
    }

    ZVAL_NULL(&read_ptr->pending_real_row);
    read_ptr->cols               = 0;
    read_ptr->expected_row_nr    = 1;
    read_ptr->pending_synth_rows = 0;
}

static inline void php_vtiful_close_resource(zend_object *obj) {
    if (obj == NULL) {
        return;
    }

    xls_object *intern = php_vtiful_xls_fetch_object(obj);

    SHEET_LINE_INIT(intern);

    if (intern->write_ptr.workbook != NULL) {
        lxlsx_workbook_free(intern->write_ptr.workbook);
        intern->write_ptr.workbook = NULL;
        intern->write_ptr.worksheet = NULL;
    }

    xls_auto_widths_reset(&intern->write_ptr);

    if (intern->lxlsx_format_ptr.format != NULL) {
        intern->lxlsx_format_ptr.format = NULL;
    }

    if (intern->formats_cache_ptr.maps != NULL) {
        zend_hash_destroy(intern->formats_cache_ptr.maps);
        /* The HashTable struct was emalloc'd in vtiful_xls_objects_new;
         * zend_hash_destroy only releases the table's contents, the struct
         * itself must be efree'd separately. */
        efree(intern->formats_cache_ptr.maps);
        intern->formats_cache_ptr.maps = NULL;
    }

    if (intern->row_options != NULL) {
        efree(intern->row_options);
        intern->row_options = NULL;
    }

    if (intern->read_ptr.sheet_t != NULL) {
        lxlsx_reader_worksheet_close(intern->read_ptr.sheet_t);
        intern->read_ptr.sheet_t = NULL;
    }

    if (intern->read_ptr.file_t != NULL) {
        lxlsx_reader_workbook_close(intern->read_ptr.file_t);
        intern->read_ptr.file_t = NULL;
    }

    php_vtiful_reset_reader_state(&intern->read_ptr);

    intern->read_ptr.data_type_default = READ_TYPE_EMPTY;
}

lxlsx_format            * zval_get_format(zval *handle);
lxlsx_data_validation   * zval_get_validation(zval *resource);
lxlsx_rich_string_tuple * zval_get_rich_string(zval *resource);
xls_resource_write_t  * zval_get_resource(zval *handle);
xls_resource_chart_t  * zval_get_chart(zval *resource);

STATIC lxlsx_error _store_defined_name(lxlsx_workbook *self, const char *name, const char *app_name, const char *formula, int16_t index, uint8_t hidden);

STATIC int  _compare_defined_names(lxlsx_defined_name *a, lxlsx_defined_name *b);

STATIC void _prepare_drawings(lxlsx_workbook *self);
STATIC void _add_chart_cache_data(lxlsx_workbook *self);
STATIC void _prepare_vml(lxlsx_workbook *self);
STATIC void _prepare_defined_names(lxlsx_workbook *self);
STATIC void _prepare_tables(lxlsx_workbook *self);
STATIC void _populate_range(lxlsx_workbook *self, lxlsx_series_range *range);
STATIC void _populate_range_dimensions(lxlsx_workbook *self, lxlsx_series_range *range);

void comment_show(xls_resource_write_t *res);
void hide_worksheet(xls_resource_write_t *res);
void first_worksheet(xls_resource_write_t *res);
void zoom(xls_resource_write_t *res, zend_long zoom);
void paper(xls_resource_write_t *res, zend_long type);
void gridlines(xls_resource_write_t *res, zend_long option);
void printed_scale(xls_resource_write_t *res, zend_long scale);
void auto_filter(zend_string *range, xls_resource_write_t *res);
void protection(xls_resource_write_t *res, zend_string *password);
void lxlsx_format_copy(lxlsx_format *new_format, lxlsx_format *other_format);
void printed_direction(xls_resource_write_t *res, unsigned int direction);
void xls_file_path(zend_string *file_name, zval *dir_path, zval *file_path);
void freeze_panes(xls_resource_write_t *res, zend_long row, zend_long column);
void margins(xls_resource_write_t *res, double left, double right, double top, double bottom);
void outline_settings(xls_resource_write_t *res, uint8_t visible, uint8_t symbols_below, uint8_t symbols_right, uint8_t auto_style);
void set_row(zend_string *range, double height, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options);
void validation(xls_resource_write_t *res, zend_string *range, lxlsx_data_validation *validation);
void set_column(zend_string *range, double width, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options);
void merge_cells(zend_string *range, zval *value, xls_resource_write_t *res, lxlsx_format *format);
void comment_writer(zend_string *comment, zend_long row, zend_long columns, xls_resource_write_t *res);
void call_object_method(zval *object, const char *function_name, uint32_t param_count, zval *params, zval *ret_val);
void lxlsx_chart_writer(zend_long row, zend_long columns, xls_resource_chart_t *lxlsx_chart_resource, xls_resource_write_t *res);
lxlsx_error lxlsx_worksheet_set_rows(lxlsx_row_t start, lxlsx_row_t end, double height, xls_resource_write_t *res, lxlsx_format *format, lxlsx_row_col_options *user_options);
void image_writer(zval *value, zend_long row, zend_long columns, double width, double height, xls_resource_write_t *res);
void image_opt_writer(zval *value, zend_long row, zend_long columns, lxlsx_image_options *options, xls_resource_write_t *res);
void formula_writer(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format);
void formula_writer_calc(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format);
void formula_resolver(void *ctx, lxlsx_row_t row, lxlsx_col_t col, lxlsx_value *out);
void dynamic_formula_writer(zend_string *value, zend_long row, zend_long columns, xls_resource_write_t *res, lxlsx_format *format);
void dynamic_array_formula_writer(zend_string *value, zend_long first_row, zend_long first_col,
                                  zend_long last_row, zend_long last_col,
                                  xls_resource_write_t *res, lxlsx_format *format);
void type_writer(zval *value, zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *format, lxlsx_format *lxlsx_format_handle);
void rich_string_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zval *rich_strings, lxlsx_format *format);
void datetime_writer(lxlsx_datetime *datetime, zend_long row, zend_long columns, zend_string *format, xls_resource_write_t *res, lxlsx_format *lxlsx_format_handle);
void url_writer(zend_long row, zend_long columns, xls_resource_write_t *res, zend_string *url, zend_string *text, zend_string *tool_tip, lxlsx_format *format);

/* Phase 2 writer helpers */
void comment_opt_writer(zend_string *comment, zend_long row, zend_long columns,
                        lxlsx_comment_options *options, xls_resource_write_t *res);
void image_buffer_writer(zend_long row, zend_long columns,
                         const unsigned char *bytes, size_t size,
                         lxlsx_image_options *options,
                         xls_resource_write_t *res);
void header_writer(xls_resource_write_t *res, const char *value,
                   lxlsx_header_footer_options *options);
void footer_writer(xls_resource_write_t *res, const char *value,
                   lxlsx_header_footer_options *options);
void repeat_rows_writer   (xls_resource_write_t *res, zend_long first_row, zend_long last_row);
void repeat_columns_writer(xls_resource_write_t *res, zend_long first_col, zend_long last_col);
void print_area_writer    (xls_resource_write_t *res,
                           zend_long first_row, zend_long first_col,
                           zend_long last_row,  zend_long last_col);
void h_pagebreaks_writer  (xls_resource_write_t *res, lxlsx_row_t *breaks);
void v_pagebreaks_writer  (xls_resource_write_t *res, lxlsx_col_t *breaks);
void fit_to_pages_writer  (xls_resource_write_t *res, zend_long width, zend_long height);
void tab_color_writer     (xls_resource_write_t *res, zend_long rgb);
void background_image_writer(xls_resource_write_t *res, const char *path);
void background_image_buffer_writer(xls_resource_write_t *res,
                                    const unsigned char *bytes, size_t size);
void lxlsx_workbook_properties_writer(xls_resource_write_t *res, lxlsx_doc_properties *props);
void define_name_writer(xls_resource_write_t *res, const char *name, const char *formula);
void conditional_format_writer(xls_resource_write_t *res,
                               zend_long first_row, zend_long first_col,
                               zend_long last_row,  zend_long last_col,
                               lxlsx_conditional_format *cf);
void add_table_writer(xls_resource_write_t *res,
                      zend_long first_row, zend_long first_col,
                      zend_long last_row,  zend_long last_col,
                      lxlsx_table_options *opts);

lxlsx_error lxlsx_workbook_file(xls_resource_write_t *self);

/* Phase 3 — formula AST */
void formula_ast_parse(const char *src, size_t n, zval *return_value);

lxlsx_datetime timestamp_to_datetime(zend_long timestamp);
zend_string* char_join_to_zend_str(const char *left, const char *right);
zend_string* str_pick_up(zend_string *left, const char *right, size_t len);

lxlsx_format* object_format(xls_object *obj, zend_string *format, lxlsx_format *lxlsx_format_handle);

#endif
