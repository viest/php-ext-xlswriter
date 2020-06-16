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

#include "xlswriter.h"
#include "ext/date/php_date.h"

zend_class_entry *vtiful_xls_ce;

static zend_object_handlers vtiful_xls_handlers;

static zend_always_inline void *vtiful_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size + zend_object_properties_size(ce));
    memset(obj, 0, obj_size);
    return obj;
}

/* {{{ xls_objects_new
 */
PHP_VTIFUL_API zend_object *vtiful_xls_objects_new(zend_class_entry *ce)
{
    xls_object *intern = vtiful_object_alloc(sizeof(xls_object), ce);

    SHEET_LINE_INIT(intern)

    zend_object_std_init(&intern->zo, ce);
    object_properties_init(&intern->zo, ce);

    intern->zo.handlers = &vtiful_xls_handlers;

    intern->read_ptr.file_t   = NULL;
    intern->read_ptr.sheet_t  = NULL;
    intern->format_ptr.format = NULL;

    return &intern->zo;
}
/* }}} */

/* {{{ vtiful_xls_objects_free
 */
static void vtiful_xls_objects_free(zend_object *object)
{
    xls_object *intern = php_vtiful_xls_fetch_object(object);

    lxw_workbook_free(intern->write_ptr.workbook);

#ifdef ENABLE_READER
    xlsxioread_sheet_close(intern->read_ptr.sheet_t);
    xlsxioread_close(intern->read_ptr.file_t);
#endif

    if (intern->format_ptr.format != NULL) {
        intern->format_ptr.format = NULL;
    }

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(xls_construct_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, config)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, file_name)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_const_memory_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, file_name)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_add_sheet, 0, 0, 1)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_checkout_sheet, 0, 0, 1)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_activate_sheet, 0, 0, 1)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_header_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, header)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_data_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, data)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_output_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_get_handle_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_text_arginfo, 0, 0, 5)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, data)
                ZEND_ARG_INFO(0, format)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_date_arginfo, 0, 0, 5)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, timestamp)
                ZEND_ARG_INFO(0, format)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_url_arginfo, 0, 0, 4)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, url)
                ZEND_ARG_INFO(0, format)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_chart_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, chart_resource)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_image_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, image)
                ZEND_ARG_INFO(0, width)
                ZEND_ARG_INFO(0, height)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_formula_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, formula)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_comment_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, comment)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_show_comment_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_auto_filter_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, range)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_merge_cells_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, data)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_column_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, format_handle)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, width)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_row_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, format_handle)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, height)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_global_format, 0, 0, 1)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_open_file_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zs_file_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_open_sheet_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zs_sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_put_csv_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, fp)
                ZEND_ARG_INFO(0, delimiter_str)
                ZEND_ARG_INFO(0, enclosure_str)
                ZEND_ARG_INFO(0, escape_str)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_put_csv_callback_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, callback)
                ZEND_ARG_INFO(0, fp)
                ZEND_ARG_INFO(0, delimiter_str)
                ZEND_ARG_INFO(0, enclosure_str)
                ZEND_ARG_INFO(0, escape_str)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_sheet_list_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_get_sheet_data_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_next_row_arginfo, 0, 0, 0)
                ZEND_ARG_INFO(0, zv_type_t)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_type_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zv_type_t)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_skip_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zv_skip_t)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_next_cell_callback_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, fci)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_index_to_string, 0, 0, 1)
                ZEND_ARG_INFO(0, index)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_string_to_index, 0, 0, 1)
                ZEND_ARG_INFO(0, index)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_freeze_panes_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_sheet_gridline_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, option)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_sheet_zoom_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, scale)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::__construct(array $config)
 */
PHP_METHOD(vtiful_xls, __construct)
{
    zval *config = NULL, *c_path = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(config)
    ZEND_PARSE_PARAMETERS_END();

    if((c_path = zend_hash_str_find(Z_ARRVAL_P(config), ZEND_STRL(V_XLS_PAT))) == NULL)
    {
        zend_throw_exception(vtiful_exception_ce, "Lack of 'path' configuration", 110);
        return;
    }

    if(Z_TYPE_P(c_path) != IS_STRING)
    {
        zend_throw_exception(vtiful_exception_ce, "Configure 'path' must be a string type", 120);
        return;
    }

    add_property_zval_ex(getThis(), ZEND_STRL(V_XLS_COF), config);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::filename(string $fileName [, string $sheetName])
 */
PHP_METHOD(vtiful_xls, fileName)
{
    char *sheet_name = NULL;
    zval file_path, *dir_path = NULL;
    zend_string *zs_file_name = NULL, *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_STR(zs_file_name)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_xls_ce, return_value);

    if(directory_exists(ZSTR_VAL(Z_STR_P(dir_path))) == XLSWRITER_FALSE) {
        zend_throw_exception(vtiful_exception_ce, "Configure 'path' directory does not exist", 121);
        return;
    }

    xls_object *obj = Z_XLS_P(getThis());

    if(obj->write_ptr.workbook == NULL) {
        xls_file_path(zs_file_name, dir_path, &file_path);

        if(zs_sheet_name != NULL) {
            sheet_name = ZSTR_VAL(zs_sheet_name);
        }

        obj->write_ptr.workbook  = workbook_new(Z_STRVAL(file_path));
        obj->write_ptr.worksheet = workbook_add_worksheet(obj->write_ptr.workbook, sheet_name);

        add_property_zval(return_value, V_XLS_FIL, &file_path);

        zval_ptr_dtor(&file_path);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::addSheet(string $sheetName)
 */
PHP_METHOD(vtiful_xls, addSheet)
{
    char *sheet_name = NULL;
    zend_string *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);
    SHEET_LINE_INIT(obj)

    if(zs_sheet_name != NULL) {
        sheet_name = ZSTR_VAL(zs_sheet_name);
    }

    obj->write_ptr.worksheet = workbook_add_worksheet(obj->write_ptr.workbook, sheet_name);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::checkoutSheet(string $sheetName)
 */
PHP_METHOD(vtiful_xls, checkoutSheet)
{
    int line = 0;
    lxw_worksheet *sheet_t = NULL;
    zend_string *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if ((sheet_t = workbook_get_worksheet_by_name(obj->write_ptr.workbook, ZSTR_VAL(zs_sheet_name))) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Sheet not fund", 140);
        return;
    }

    line = sheet_t->table->cached_row_num + 1;

    SHEET_LINE_SET(obj, line);

    obj->write_ptr.worksheet = sheet_t;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::activateSheet(string $sheetName)
 */
PHP_METHOD(vtiful_xls, activateSheet)
{
    lxw_worksheet *sheet_t = NULL;
    zend_string *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if ((sheet_t = workbook_get_worksheet_by_name(obj->write_ptr.workbook, ZSTR_VAL(zs_sheet_name))) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Sheet not fund", 140);
        return;
    }

    worksheet_activate(sheet_t);

    RETURN_TRUE;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::constMemory(string $fileName [, string $sheetName])
 */
PHP_METHOD(vtiful_xls, constMemory)
{
    char *sheet_name = NULL;
    zval file_path, *dir_path = NULL;
    zend_string *zs_file_name = NULL, *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_STR(zs_file_name)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_xls_ce, return_value);

    xls_object *obj = Z_XLS_P(getThis());

    if(obj->write_ptr.workbook == NULL) {
        xls_file_path(zs_file_name, dir_path, &file_path);

        lxw_workbook_options options = {.constant_memory = LXW_TRUE, .tmpdir = NULL};

        if(zs_sheet_name != NULL) {
            sheet_name = ZSTR_VAL(zs_sheet_name);
        }

        obj->write_ptr.workbook  = workbook_new_opt(Z_STRVAL(file_path), &options);
        obj->write_ptr.worksheet = workbook_add_worksheet(obj->write_ptr.workbook, sheet_name);

        add_property_zval(return_value, V_XLS_FIL, &file_path);

        zval_ptr_dtor(&file_path);
    }
}
/* }}} */


/** {{{ \Vtiful\Kernel\Excel::header(array $header)
 */
PHP_METHOD(vtiful_xls, header)
{
    zend_long header_l_key;
    lxw_format *format_handle = NULL;
    zval *header = NULL, *header_value = NULL, *zv_format_handle = NULL;;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_ARRAY(header)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(zv_format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (zv_format_handle == NULL) {
        format_handle = obj->format_ptr.format;
    } else {
        format_handle = zval_get_format(zv_format_handle);
    }

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(header), header_l_key, header_value)
         type_writer(header_value, 0, header_l_key, &obj->write_ptr, NULL, format_handle);
    ZEND_HASH_FOREACH_END();

    SHEET_LINE_ADD(obj)
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_xls, data)
{
    zval *data = NULL, *data_r_value = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    ZEND_HASH_FOREACH_VAL(Z_ARRVAL_P(data), data_r_value)
        if(Z_TYPE_P(data_r_value) == IS_ARRAY) {
            ZEND_HASH_FOREACH_BUCKET(Z_ARRVAL_P(data_r_value), Bucket *bucket)
                type_writer(&bucket->val, SHEET_CURRENT_LINE(obj), bucket->h, &obj->write_ptr, NULL, obj->format_ptr.format);
            ZEND_HASH_FOREACH_END();

            SHEET_LINE_ADD(obj)
        }
    ZEND_HASH_FOREACH_END();
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_xls, output)
{
    zval rv, *file_path = NULL;

    file_path = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_FIL), 0, &rv TSRMLS_DC);

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    workbook_file(&obj->write_ptr);

    ZVAL_COPY(return_value, file_path);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::getHandle()
 */
PHP_METHOD(vtiful_xls, getHandle)
{
    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    RETURN_RES(zend_register_resource(&obj->write_ptr, le_xls_writer));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertText(int $row, int $column, string|int|double $data[, string $format, resource $formatHandle])
 */
PHP_METHOD(vtiful_xls, insertText)
{
    zend_long row = 0, column = 0;
    zend_string *format = NULL;
    zval *data = NULL, *format_handle = NULL;

    ZEND_PARSE_PARAMETERS_START(3, 5)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(data)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(format)
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    SHEET_LINE_SET(obj, row);

    if (format_handle) {
        type_writer(data, row, column, &obj->write_ptr, format, zval_get_format(format_handle));
    } else {
        type_writer(data, row, column, &obj->write_ptr, format, obj->format_ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertDate(int $row, int $column, int $timestamp[, string $format, resource $formatHandle])
 */
PHP_METHOD(vtiful_xls, insertDate)
{
    zval *data = NULL, *format_handle = NULL;
    zend_long row = 0, column = 0;
    zend_string *format = NULL;

    ZEND_PARSE_PARAMETERS_START(3, 5)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(data)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(format)
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);
    SHEET_LINE_SET(obj, row);

    if (Z_TYPE_P(data) != IS_LONG) {
        zend_throw_exception(vtiful_exception_ce, "timestamp is long", 160);
        return;
    }

    // Default datetime format
    if (format == NULL) {
        format = zend_string_init(ZEND_STRL("yyyy-mm-dd hh:mm:ss"), 0);
    }

    int yearLocal   = php_idate('Y', data->value.lval, 0);
    int monthLocal  = php_idate('m', data->value.lval, 0);
    int dayLocal    = php_idate('d', data->value.lval, 0);
    int hourLocal   = php_idate('H', data->value.lval, 0);
    int minuteLocal = php_idate('i', data->value.lval, 0);
    int secondLocal = php_idate('s', data->value.lval, 0);

    lxw_datetime datetime = {
            yearLocal, monthLocal, dayLocal, hourLocal, minuteLocal, secondLocal
    };

    if (format_handle) {
        datetime_writer(&datetime, row, column, format, &obj->write_ptr, zval_get_format(format_handle));
    } else {
        datetime_writer(&datetime, row, column, format, &obj->write_ptr, obj->format_ptr.format);
    }

    // Release default format
    if (ZEND_NUM_ARGS() == 3) {
        zend_string_release(format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertChart(int $row, int $column, resource $chartResource)
 */
PHP_METHOD(vtiful_xls, insertChart)
{
    zval *chart_resource = NULL;
    zend_long row = 0, column = 0;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(chart_resource)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    chart_writer(row, column, zval_get_chart(chart_resource), &obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertUrl(int $row, int $column, string $url)
 */
PHP_METHOD(vtiful_xls, insertUrl)
{
    zend_long row = 0, column = 0;
    zend_string *url = NULL;
    zval *format_handle = NULL;

    int argc = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(3, 4)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_STR(url)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 4) {
        url_writer(row, column, &obj->write_ptr, url, zval_get_format(format_handle));
    }

    if (argc == 3) {
        url_writer(row, column, &obj->write_ptr, url, obj->format_ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_xls, insertImage)
{
    zval *image = NULL;
    zend_long row = 0, column = 0;
    double width = 1, height = 1;

    ZEND_PARSE_PARAMETERS_START(3, 5)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(image)
            Z_PARAM_OPTIONAL
            Z_PARAM_DOUBLE(width)
            Z_PARAM_DOUBLE(height)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    image_writer(image, row, column, width, height,  &obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertFormula(int $row, int $column, string $formula)
 */
PHP_METHOD(vtiful_xls, insertFormula)
{
    zval *format_handle = NULL;
    zend_string *formula = NULL;
    zend_long row = 0, column = 0;

    int argc = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(3, 4)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_STR(formula)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3) {
        formula_writer(formula, row, column, &obj->write_ptr, obj->format_ptr.format);
    }

    if (argc == 4) {
        formula_writer(formula, row, column, &obj->write_ptr, zval_get_format(format_handle));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertComment(int $row, int $column, string $comment)
 */
PHP_METHOD(vtiful_xls, insertComment)
{
    zend_string *comment = NULL;
    zend_long row = 0, column = 0;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_STR(comment)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    comment_writer(comment, row, column, &obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::showComment()
 */
PHP_METHOD(vtiful_xls, showComment)
{
    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    comment_show(&obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::autoFilter(int $rowStart, int $rowEnd, int $columnStart, int $columnEnd)
 */
PHP_METHOD(vtiful_xls, autoFilter)
{
    zend_string *range = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(range)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    auto_filter(range, &obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::mergeCells(string $range, string $data, $formatHandle = NULL)
 */
PHP_METHOD(vtiful_xls, mergeCells)
{
    zend_string *range = NULL;
    zval *data = NULL, *format_handle = NULL;

    int argc = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(2, 3)
            Z_PARAM_STR(range)
            Z_PARAM_ZVAL(data)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 2) {
        merge_cells(range, data, &obj->write_ptr, obj->format_ptr.format);
    }

    if (argc == 3) {
        merge_cells(range, data, &obj->write_ptr, zval_get_format(format_handle));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setColumn(resource $format, string $range [, int $width])
 */
PHP_METHOD(vtiful_xls, setColumn)
{
    zval *format_handle = NULL;
    zend_string *range = NULL;

    double width = 0;
    int    argc  = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(2, 3)
            Z_PARAM_STR(range)
            Z_PARAM_DOUBLE(width)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3) {
        set_column(range, width, &obj->write_ptr, zval_get_format(format_handle));
    }

    if (argc == 2) {
        set_column(range, width, &obj->write_ptr, NULL);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setRow(resource $format, string $range [, int $heitght])
 */
PHP_METHOD(vtiful_xls, setRow)
{
    zval *format_handle = NULL;
    zend_string *range = NULL;

    double height = 0;
    int    argc  = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(2, 3)
            Z_PARAM_STR(range)
            Z_PARAM_DOUBLE(height)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3) {
        set_row(range, height, &obj->write_ptr, zval_get_format(format_handle));
    }

    if (argc == 2) {
        set_row(range, height, &obj->write_ptr, NULL);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::defaultFormat(resource $format)
 */
PHP_METHOD(vtiful_xls, defaultFormat)
{
    zval *format_handle = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_RESOURCE(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());
    obj->format_ptr.format = zval_get_format(format_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::freezePanes(int $row, int $column)
 */
PHP_METHOD(vtiful_xls, freezePanes)
{
    zend_long row = 0, column = 0;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    freeze_panes(&obj->write_ptr, row, column);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::columnIndexFromString(string $index)
 */
PHP_METHOD(vtiful_xls, columnIndexFromString)
{
    zend_string *index = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(index)
    ZEND_PARSE_PARAMETERS_END();

    RETURN_LONG(lxw_name_to_col(ZSTR_VAL(index)));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::stringFromColumnIndex(int $index)
 */
PHP_METHOD(vtiful_xls, stringFromColumnIndex)
{
    zend_long index = 0, current = 0;
    zend_string *result = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(index)
    ZEND_PARSE_PARAMETERS_END();

    if (index < 26) {
        current = index + 65;
        result  = zend_string_init((char *)(&current), 1, 0);
        RETURN_STR(result);
    }

    if (index < 702) {
        current = index / 26 + 64;
        result  = zend_string_init((char *)(&current), 1, 0);

        current = index % 26 + 65;
        RETURN_STR(str_pick_up(result, (char *)(&current)));
    }

    current = (index - 26) / 676 + 64;
    result  = zend_string_init((char *)(&current), 1, 0);

    current = ((index - 26) % 676) / 26 + 65;
    result  = str_pick_up(result, (char *)(&current));

    current = index % 26 + 65;
    RETURN_STR(str_pick_up(result, (char *)(&current)));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::timestampFromDateDouble(string $date)
 */
PHP_METHOD(vtiful_xls, timestampFromDateDouble)
{
    double date = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_DOUBLE(date)
    ZEND_PARSE_PARAMETERS_END();

    if (date <= 0) {
        RETURN_LONG(0);
    }

    RETURN_LONG(date_double_to_timestamp(date));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::gridline(int $option = \Vtiful\Kernel\Excel::GRIDLINES_SHOW_ALL)
 */
PHP_METHOD(vtiful_xls, gridline)
{
    zend_long option = LXW_SHOW_ALL_GRIDLINES;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(option)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    gridlines(&obj->write_ptr, option);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::zoom(int $scale)
 */
PHP_METHOD(vtiful_xls, zoom)
{
    zend_long scale = 100;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(scale)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    if (scale < 10) {
        scale = 10;
    }

    if (scale > 400) {
        scale = 400;
    }

    xls_object* obj = Z_XLS_P(getThis());

    zoom(&obj->write_ptr, scale);
}
/* }}} */

#ifdef ENABLE_READER

/** {{{ \Vtiful\Kernel\Excel::openFile()
 */
PHP_METHOD(vtiful_xls, openFile)
{
    zval *zv_config_path = NULL;
    zend_string *zs_file_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
        Z_PARAM_STR(zs_file_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(zv_config_path, vtiful_xls_ce, return_value);

    xls_object* obj = Z_XLS_P(getThis());

    obj->read_ptr.file_t = file_open(Z_STRVAL_P(zv_config_path), ZSTR_VAL(zs_file_name));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::openSheet()
 */
PHP_METHOD(vtiful_xls, openSheet)
{
    zend_long zl_flag = XLSXIOREAD_SKIP_NONE;
    zend_string *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 2)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(zs_sheet_name)
            Z_PARAM_LONG(zl_flag)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    if (obj->read_ptr.file_t == NULL) {
        RETURN_NULL();
    }

    if (obj->read_ptr.sheet_t != NULL) {
        xlsxioread_sheet_close(obj->read_ptr.sheet_t);
    }

    obj->read_ptr.sheet_t = sheet_open(obj->read_ptr.file_t, zs_sheet_name, zl_flag);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::sheetList()
 */
PHP_METHOD(vtiful_xls, sheetList)
{
    xls_object* obj = Z_XLS_P(getThis());

    if (obj->read_ptr.file_t == NULL) {
        RETURN_NULL();
    }

    sheet_list(obj->read_ptr.file_t, return_value);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setType(array $rowType)
 */
PHP_METHOD(vtiful_xls, setType)
{
    zval *zv_type_t = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(zv_type_t)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    add_property_zval_ex(getThis(), ZEND_STRL(V_XLS_TYPE), zv_type_t);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setSkipRows(int $skip)
 */
PHP_METHOD(vtiful_xls, setSkipRows)
{
    zend_long zl_skip = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_skip)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    skip_rows(obj->read_ptr.sheet_t, NULL, zl_skip);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::putCSV()
 */
PHP_METHOD(vtiful_xls, putCSV)
{
    zval *fp = NULL, *zv_type = NULL;
    char *delimiter_str = NULL, *enclosure_str = NULL, *escape_str = NULL;
    size_t delimiter_str_len = 0, enclosure_str_len = 0, escape_str_len = 0;

    ZEND_PARSE_PARAMETERS_START(1, 4)
            Z_PARAM_RESOURCE(fp)
            Z_PARAM_OPTIONAL
            Z_PARAM_STRING(delimiter_str, delimiter_str_len)
            Z_PARAM_STRING(enclosure_str, enclosure_str_len)
            Z_PARAM_STRING(escape_str,escape_str_len)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    zv_type = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (xlsx_to_csv(
            fp, delimiter_str, delimiter_str_len, enclosure_str, enclosure_str_len, escape_str, escape_str_len,
            obj->read_ptr.sheet_t, zv_type, READ_SKIP_ROW, NULL, NULL
            ) == XLSWRITER_TRUE) {
        RETURN_TRUE;
    }

    RETURN_FALSE;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::putCSVCallback()
 */
PHP_METHOD(vtiful_xls, putCSVCallback)
{
    zval *fp = NULL, *zv_type = NULL;
    zend_fcall_info fci = empty_fcall_info;
    zend_fcall_info_cache fci_cache = empty_fcall_info_cache;
    char *delimiter_str = NULL, *enclosure_str = NULL, *escape_str = NULL;
    size_t delimiter_str_len = 0, enclosure_str_len = 0, escape_str_len = 0;

    ZEND_PARSE_PARAMETERS_START(2, 5)
            Z_PARAM_FUNC(fci, fci_cache)
            Z_PARAM_RESOURCE(fp)
            Z_PARAM_OPTIONAL
            Z_PARAM_STRING(delimiter_str, delimiter_str_len)
            Z_PARAM_STRING(enclosure_str, enclosure_str_len)
            Z_PARAM_STRING(escape_str,escape_str_len)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    zv_type = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (xlsx_to_csv(
            fp, delimiter_str, delimiter_str_len, enclosure_str, enclosure_str_len, escape_str, escape_str_len,
            obj->read_ptr.sheet_t, zv_type, READ_SKIP_ROW, &fci, &fci_cache
            ) == XLSWRITER_TRUE) {
        RETURN_TRUE;
    }

    RETURN_FALSE;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::getSheetData()
 */
PHP_METHOD(vtiful_xls, getSheetData)
{
    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    zval *zv_type = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (zv_type != NULL && Z_TYPE_P(zv_type) == IS_ARRAY) {
        load_sheet_all_data(obj->read_ptr.sheet_t, zv_type, return_value);

        return;
    }

    load_sheet_all_data(obj->read_ptr.sheet_t, NULL, return_value);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::nextRow()
 */
PHP_METHOD(vtiful_xls, nextRow)
{
    zval *zv_type_t = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_ARRAY(zv_type_t)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    if (zv_type_t == NULL) {
        zv_type_t = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_TYPE), 0, NULL);
    }

    load_sheet_current_row_data(obj->read_ptr.sheet_t, return_value, zv_type_t, READ_ROW);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::nextCellCallback()
 */
PHP_METHOD(vtiful_xls, nextCellCallback)
{
    zend_string *zs_sheet_name = NULL;
    zend_fcall_info       fci       = empty_fcall_info;
    zend_fcall_info_cache fci_cache = empty_fcall_info_cache;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_FUNC(fci, fci_cache)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.file_t) {
        RETURN_FALSE;
    }

    xls_read_callback_data callback_data;

    callback_data.zv_type_t = zend_read_property(vtiful_xls_ce, getThis(), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    callback_data.fci = &fci;
    callback_data.fci_cache = &fci_cache;

    load_sheet_current_row_data_callback(zs_sheet_name, obj->read_ptr.file_t, &callback_data);
}
/* }}} */

#endif

/** {{{ xls_methods
*/
zend_function_entry xls_methods[] = {
        PHP_ME(vtiful_xls, __construct,   xls_construct_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, fileName,      xls_file_name_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, addSheet,      xls_file_add_sheet,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, checkoutSheet, xls_file_checkout_sheet,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, activateSheet, xls_file_activate_sheet,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, constMemory,   xls_const_memory_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, header,        xls_header_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, data,          xls_data_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, output,        xls_output_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, getHandle,     xls_get_handle_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, autoFilter,    xls_auto_filter_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertText,    xls_insert_text_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertDate,    xls_insert_date_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertChart,   xls_insert_chart_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertUrl,     xls_insert_url_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertImage,   xls_insert_image_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertFormula, xls_insert_formula_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertComment, xls_insert_comment_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, showComment,   xls_show_comment_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, mergeCells,    xls_merge_cells_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setColumn,     xls_set_column_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setRow,        xls_set_row_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, defaultFormat, xls_set_global_format,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, freezePanes,   xls_freeze_panes_arginfo,   ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, zoom,          xls_sheet_zoom_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, gridline,      xls_sheet_gridline_arginfo, ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, columnIndexFromString,   xls_index_to_string, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_ME(vtiful_xls, stringFromColumnIndex,   xls_string_to_index, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_ME(vtiful_xls, timestampFromDateDouble, xls_string_to_index, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)

#ifdef ENABLE_READER
        PHP_ME(vtiful_xls, openFile,         xls_open_file_arginfo,          ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, openSheet,        xls_open_sheet_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, putCSV,           xls_put_csv_arginfo,            ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, putCSVCallback,   xls_put_csv_callback_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, sheetList,        xls_sheet_list_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setType,          xls_set_type_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setSkipRows,      xls_set_skip_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, getSheetData,     xls_get_sheet_data_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, nextRow,          xls_next_row_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, nextCellCallback, xls_next_cell_callback_arginfo, ZEND_ACC_PUBLIC)
#endif

        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(excel) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Excel", xls_methods);
    ce.create_object = vtiful_xls_objects_new;
    vtiful_xls_ce = zend_register_internal_class(&ce);

    memcpy(&vtiful_xls_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    vtiful_xls_handlers.offset   = XtOffsetOf(xls_object, zo);
    vtiful_xls_handlers.free_obj = vtiful_xls_objects_free;

    REGISTER_CLASS_PROPERTY_NULL(vtiful_xls_ce, V_XLS_COF,  ZEND_ACC_PRIVATE);
    REGISTER_CLASS_PROPERTY_NULL(vtiful_xls_ce, V_XLS_FIL,  ZEND_ACC_PRIVATE);
    REGISTER_CLASS_PROPERTY_NULL(vtiful_xls_ce, V_XLS_TYPE, ZEND_ACC_PRIVATE);

#ifdef ENABLE_READER
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_NONE,        XLSXIOREAD_SKIP_NONE);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_EMPTY_ROW,   XLSXIOREAD_SKIP_EMPTY_ROWS);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_EMPTY_CELLS, XLSXIOREAD_SKIP_EMPTY_CELLS);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_EMPTY_VALUE, SKIP_EMPTY_VALUE);
#endif

    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_HIDE_ALL",    LXW_HIDE_ALL_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_ALL",    LXW_SHOW_ALL_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_PRINT",  LXW_SHOW_PRINT_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_SCREEN", LXW_SHOW_SCREEN_GRIDLINES)

    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_INT,      READ_TYPE_INT);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_DOUBLE,   READ_TYPE_DOUBLE);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_STRING,   READ_TYPE_STRING);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_DATETIME, READ_TYPE_DATETIME);

    return SUCCESS;
}
/* }}} */



