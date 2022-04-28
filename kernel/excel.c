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

    intern->format_ptr.format  = NULL;
    intern->write_ptr.workbook = NULL;

    intern->read_ptr.data_type_default = READ_TYPE_EMPTY;

    return &intern->zo;
}
/* }}} */

/* {{{ vtiful_xls_objects_free
 */
static void vtiful_xls_objects_free(zend_object *object)
{
    xls_object *intern = php_vtiful_xls_fetch_object(object);

    php_vtiful_close_resource(object);

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(xls_construct_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, config)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_close_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, file_name)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_const_memory_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, file_name)
                ZEND_ARG_INFO(0, sheet_name)
                ZEND_ARG_INFO(0, use_zip64)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_add_sheet, 0, 0, 1)
                ZEND_ARG_INFO(0, sheet_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_file_exist_sheet, 0, 0, 1)
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

ZEND_BEGIN_ARG_INFO_EX(xls_insert_text_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, data)
                ZEND_ARG_INFO(0, format)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_rtext_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, rich_strings)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_date_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, timestamp)
                ZEND_ARG_INFO(0, format)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_insert_url_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, url)
                ZEND_ARG_INFO(0, text)
                ZEND_ARG_INFO(0, tool_tip)
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

ZEND_BEGIN_ARG_INFO_EX(xls_set_curr_line_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, row)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_get_curr_line_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_paper_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, paper)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_margins_arginfo, 0, 0, 4)
                ZEND_ARG_INFO(0, left)
                ZEND_ARG_INFO(0, right)
                ZEND_ARG_INFO(0, top)
                ZEND_ARG_INFO(0, bottom)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_global_format, 0, 0, 1)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_open_file_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zs_file_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_open_sheet_arginfo, 0, 0, 0)
                ZEND_ARG_INFO(0, zs_sheet_name)
                ZEND_ARG_INFO(0, zl_flag)
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

ZEND_BEGIN_ARG_INFO_EX(xls_set_global_type_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zv_type_t)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_skip_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, zv_skip_t)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_next_cell_callback_arginfo, 0, 0, 1)
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

ZEND_BEGIN_ARG_INFO_EX(xls_protection_arginfo, 0, 0, 0)
                ZEND_ARG_INFO(0, password)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_validation_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, validation_resource)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_printed_portrait_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_set_printed_landscape_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_hide_sheet_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(xls_first_sheet_arginfo, 0, 0, 0)
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

/** {{{ \Vtiful\Kernel\Excel::close()
 */
PHP_METHOD(vtiful_xls, close)
{
    php_vtiful_close_resource(Z_OBJ_P(getThis()));

    ZVAL_COPY(return_value, getThis());
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
            Z_PARAM_STR_OR_NULL(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_xls_ce, PROP_OBJ(return_value));

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
            Z_PARAM_STR_OR_NULL(zs_sheet_name)
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

/** {{{ \Vtiful\Kernel\Excel::existSheet(string $sheetName)
 */
PHP_METHOD(vtiful_xls, existSheet)
{
    char *sheet_name = NULL;
    zend_string *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);
    SHEET_LINE_INIT(obj)

    sheet_name = ZSTR_VAL(zs_sheet_name);

    if (workbook_get_worksheet_by_name(obj->write_ptr.workbook, sheet_name)) {
        RETURN_TRUE;
    }

    RETURN_FALSE;
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

    // sheet not insert data
    if (sheet_t->table->cached_row_num > LXW_ROW_MAX) {
        line = 0;
    }

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
    zend_bool use_zip64 = LXW_TRUE;
    zval file_path, *dir_path = NULL;
    zend_string *zs_file_name = NULL, *zs_sheet_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 3)
            Z_PARAM_STR(zs_file_name)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR_OR_NULL(zs_sheet_name)
            Z_PARAM_BOOL_OR_NULL(use_zip64, _dummy)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_xls_ce, PROP_OBJ(return_value));

    xls_object *obj = Z_XLS_P(getThis());

    if(obj->write_ptr.workbook == NULL) {
        xls_file_path(zs_file_name, dir_path, &file_path);

        lxw_workbook_options options = {
            .constant_memory = LXW_TRUE,
            .tmpdir = NULL,
            .use_zip64 = use_zip64
        };

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

/** {{{ \Vtiful\Kernel\Excel::setCurrentLine(int $row)
 */
PHP_METHOD(vtiful_xls, setCurrentLine)
{
    zend_long row = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(row)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (row < SHEET_CURRENT_LINE(obj)) {
        zend_throw_exception(vtiful_exception_ce, "The row number is abnormal, the behavior will overwrite the previous data", 400);
        return;
    }

    SHEET_LINE_SET(obj, row);
}

/** {{{ \Vtiful\Kernel\Excel::getCurrentLine()
 */
PHP_METHOD(vtiful_xls, getCurrentLine)
{
    xls_object *obj = Z_XLS_P(getThis());

    RETURN_LONG(SHEET_CURRENT_LINE(obj));
}

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
            Z_PARAM_RESOURCE_OR_NULL(zv_format_handle)
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

    // When inserting the header for the first time, the row number is incremented by one,
    // and there is no need to increment by one again for subsequent calls.
    if (obj->write_line == 0) {
        SHEET_LINE_ADD(obj)
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_xls, data)
{
    zend_ulong column_index = 0;
    zval *data = NULL, *data_r_value = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    ZEND_HASH_FOREACH_VAL(Z_ARRVAL_P(data), data_r_value)
        if (Z_TYPE_P(data_r_value) == IS_REFERENCE) {
            data_r_value = Z_REFVAL_P(data_r_value);
        }

        if(Z_TYPE_P(data_r_value) != IS_ARRAY) {
            continue;
        }

        column_index = 0;

        ZEND_HASH_FOREACH_BUCKET(Z_ARRVAL_P(data_r_value), Bucket *bucket)
            // numeric index rewriting
            if (bucket->key == NULL) {
                column_index = bucket->h;
            }

            type_writer(&bucket->val, SHEET_CURRENT_LINE(obj), column_index, &obj->write_ptr, NULL, obj->format_ptr.format);

            // next number index
            ++column_index;
        ZEND_HASH_FOREACH_END();

        SHEET_LINE_ADD(obj)
    ZEND_HASH_FOREACH_END();
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_xls, output)
{
    zval rv, *file_path = NULL;

    file_path = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_FIL), 0, &rv TSRMLS_DC);

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
            Z_PARAM_STR_OR_NULL(format)
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    SHEET_LINE_SET(obj, row);

    if (format_handle != NULL) {
        type_writer(data, row, column, &obj->write_ptr, format, zval_get_format(format_handle));
    } else {
        type_writer(data, row, column, &obj->write_ptr, format, obj->format_ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertRichText(int $row, int $column, array $richString[, resource $formatHandle])
 */
PHP_METHOD(vtiful_xls, insertRichText)
{
    zend_long row = 0, column = 0;
    zval *rich_strings = NULL, *format_handle = NULL;

    ZEND_PARSE_PARAMETERS_START(3, 4)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ARRAY(rich_strings)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    SHEET_LINE_SET(obj, row);

    if (format_handle != NULL) {
        rich_string_writer(row, column, &obj->write_ptr, rich_strings, zval_get_format(format_handle));
    } else {
        rich_string_writer(row, column, &obj->write_ptr, rich_strings, obj->format_ptr.format);
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
            Z_PARAM_STR_OR_NULL(format)
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
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
    if (format == NULL || (format != NULL && ZSTR_LEN(format) == 0)) {
        format = zend_string_init(ZEND_STRL("yyyy-mm-dd hh:mm:ss"), 0);
    }

    lxw_datetime datetime = timestamp_to_datetime(data->value.lval);

    if (format_handle != NULL) {
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
    zval *format_handle = NULL;
    zend_string *url = NULL, *text = NULL, *tool_tip = NULL;

    int argc = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(3, 6)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_STR(url)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR_OR_NULL(text)
            Z_PARAM_STR_OR_NULL(tool_tip)
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (format_handle != NULL) {
        url_writer(row, column, &obj->write_ptr, url, text, tool_tip, zval_get_format(format_handle));
    } else {
        url_writer(row, column, &obj->write_ptr, url, text, tool_tip, obj->format_ptr.format);
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
            Z_PARAM_DOUBLE_OR_NULL(width, _dummy)
            Z_PARAM_DOUBLE_OR_NULL(height, _dummy)
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
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3) {
        formula_writer(formula, row, column, &obj->write_ptr, obj->format_ptr.format);
    }

    if (argc == 4 && format_handle != NULL) {
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
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 2) {
        merge_cells(range, data, &obj->write_ptr, obj->format_ptr.format);
    }

    if (argc == 3 && format_handle != NULL) {
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
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3 && format_handle != NULL) {
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
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    if (argc == 3 && format_handle != NULL) {
        set_row(range, height, &obj->write_ptr, zval_get_format(format_handle));
    }

    if (argc == 2) {
        set_row(range, height, &obj->write_ptr, NULL);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setPaper(int $paper)
 */
PHP_METHOD(vtiful_xls, setPaper)
{
    zend_long type = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(type)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    paper(&obj->write_ptr, type);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setMargins(double|null $left, double|null $right, double|null $top, double|null $bottom)
 */
PHP_METHOD(vtiful_xls, setMargins)
{
    double left = 0.7, right = 0.7, top = 0.75, bottom = 0.75;

    ZEND_PARSE_PARAMETERS_START(0, 4)
            Z_PARAM_OPTIONAL
            Z_PARAM_DOUBLE_OR_NULL(left, _dummy)
            Z_PARAM_DOUBLE_OR_NULL(right, _dummy)
            Z_PARAM_DOUBLE_OR_NULL(top, _dummy)
            Z_PARAM_DOUBLE_OR_NULL(bottom, _dummy)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    // units: inches to cm
    margins(&obj->write_ptr, left / 2.54, right / 2.54, top / 2.54, bottom / 2.54);
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

    char one[1];

    if (index < 26) {
        current = index + 65;
        one[0] = current;

        ZVAL_STRINGL(return_value, one, 1);
        return;
    }

    if (index < 702) {
        current = index / 26 + 64;
        one[0]  = current;
        result  = zend_string_init(one, 1, 0);

        current = index % 26 + 65;
        one[0]  = current;
        ZVAL_STR(return_value, str_pick_up(result, one, 1));
        return;
    }

    current = (index - 26) / 676 + 64;
    one[0]  = current;
    result  = zend_string_init(one, 1, 0);

    current = ((index - 26) % 676) / 26 + 65;
    one[0]  = current;
    result  = str_pick_up(result, one, 1);

    current = index % 26 + 65;
    one[0]  = current;
    ZVAL_STR(return_value, str_pick_up(result, one, 1));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::timestampFromDateDouble(string $date)
 */
PHP_METHOD(vtiful_xls, timestampFromDateDouble)
{
    double date = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_DOUBLE_OR_NULL(date, _dummy)
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

/** {{{ \Vtiful\Kernel\Excel::protection(string $password)
 */
PHP_METHOD(vtiful_xls, protection)
{
    zend_string *password = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR_OR_NULL(password)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    protection(&obj->write_ptr, password);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setPortrait()
 */
PHP_METHOD(vtiful_xls, setPortrait)
{
    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    printed_direction(&obj->write_ptr, XLSWRITER_PRINTED_PORTRAIT);
}
/* }}} */


/** {{{ \Vtiful\Kernel\Excel::setLandscape()
 */
PHP_METHOD(vtiful_xls, setLandscape)
{
    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    printed_direction(&obj->write_ptr, XLSWRITER_PRINTED_LANDSCAPE);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setCurrentSheetHide()
 */
PHP_METHOD(vtiful_xls, setCurrentSheetHide)
{
    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    hide_worksheet(&obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setCurrentSheetIsFirst()
 */
PHP_METHOD(vtiful_xls, setCurrentSheetIsFirst)
{
    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    first_worksheet(&obj->write_ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::validation()
 */
PHP_METHOD(vtiful_xls, validation)
{
    zend_string *range = NULL;
    zval *validation_handle = NULL;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_STR(range)
            Z_PARAM_RESOURCE(validation_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object *obj = Z_XLS_P(getThis());

    WORKBOOK_NOT_INITIALIZED(obj);

    validation(&obj->write_ptr, range, zval_get_validation(validation_handle));
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

    GET_CONFIG_PATH(zv_config_path, vtiful_xls_ce, PROP_OBJ(return_value));

    xls_object* obj = Z_XLS_P(getThis());

    if (obj->read_ptr.sheet_t != NULL) {
        xlsxioread_sheet_close(obj->read_ptr.sheet_t);
        obj->read_ptr.sheet_t = NULL;
    }

    if (obj->read_ptr.file_t != NULL) {
        xlsxioread_close(obj->read_ptr.file_t);
        obj->read_ptr.file_t = NULL;
    }

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
            Z_PARAM_STR_OR_NULL(zs_sheet_name)
            Z_PARAM_LONG_OR_NULL(zl_flag, _dummy)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(getThis());

    if (obj->read_ptr.file_t == NULL) {
        RETURN_NULL();
    }

    if (obj->read_ptr.sheet_t != NULL) {
        xlsxioread_sheet_close(obj->read_ptr.sheet_t);
    }

    obj->read_ptr.sheet_flag = zl_flag;
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

/** {{{ \Vtiful\Kernel\Excel::setGlobalType(int $rowType)
 */
PHP_METHOD(vtiful_xls, setGlobalType)
{
    zend_long zl_type = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_type)
    ZEND_PARSE_PARAMETERS_END();

    if (zl_type < READ_TYPE_STRING || zl_type > READ_TYPE_DATETIME) {
        zend_throw_exception(vtiful_exception_ce, "Invalid data type", 220);
        return;
    }

    if (zl_type != READ_TYPE_STRING && (zl_type % 2) != 0) {
        zend_throw_exception(vtiful_exception_ce, "Invalid data type", 220);
        return;
    }

    ZVAL_COPY(return_value, getThis());

    xls_object* obj = Z_XLS_P(return_value);
    obj->read_ptr.data_type_default = zl_type;
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

    skip_rows(obj->read_ptr.sheet_t, NULL, obj->read_ptr.data_type_default, zl_skip);
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
            Z_PARAM_STRING_OR_NULL(delimiter_str, delimiter_str_len)
            Z_PARAM_STRING_OR_NULL(enclosure_str, enclosure_str_len)
            Z_PARAM_STRING_OR_NULL(escape_str,escape_str_len)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    zv_type = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (xlsx_to_csv(
            fp, delimiter_str, delimiter_str_len, enclosure_str, enclosure_str_len, escape_str, escape_str_len,
            obj->read_ptr.sheet_t, zv_type, obj->read_ptr.data_type_default, READ_SKIP_ROW, NULL, NULL
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
            Z_PARAM_STRING_OR_NULL(delimiter_str, delimiter_str_len)
            Z_PARAM_STRING_OR_NULL(enclosure_str, enclosure_str_len)
            Z_PARAM_STRING_OR_NULL(escape_str,escape_str_len)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    zv_type = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (xlsx_to_csv(
            fp, delimiter_str, delimiter_str_len, enclosure_str, enclosure_str_len, escape_str, escape_str_len,
            obj->read_ptr.sheet_t, zv_type, obj->read_ptr.data_type_default, READ_SKIP_ROW, &fci, &fci_cache
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

    zval *zv_type = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    if (zv_type != NULL && Z_TYPE_P(zv_type) == IS_ARRAY) {
        load_sheet_all_data(obj->read_ptr.sheet_t, obj->read_ptr.sheet_flag, zv_type, obj->read_ptr.data_type_default, return_value);

        return;
    }

    load_sheet_all_data(obj->read_ptr.sheet_t, obj->read_ptr.sheet_flag, NULL, obj->read_ptr.data_type_default, return_value);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::nextRow()
 */
PHP_METHOD(vtiful_xls, nextRow)
{
    zval *zv_type_t = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_ARRAY_OR_NULL(zv_type_t)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.sheet_t) {
        RETURN_FALSE;
    }

    if (zv_type_t == NULL) {
        zv_type_t = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_TYPE), 0, NULL);
    }

    load_sheet_row_data(obj->read_ptr.sheet_t, obj->read_ptr.sheet_flag, zv_type_t, obj->read_ptr.data_type_default, return_value);
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
            Z_PARAM_STR_OR_NULL(zs_sheet_name)
    ZEND_PARSE_PARAMETERS_END();

    xls_object *obj = Z_XLS_P(getThis());

    if (!obj->read_ptr.file_t) {
        RETURN_FALSE;
    }

    xls_read_callback_data callback_data;

    callback_data.data_type_default = obj->read_ptr.data_type_default;
    callback_data.zv_type_t = zend_read_property(vtiful_xls_ce, PROP_OBJ(getThis()), ZEND_STRL(V_XLS_TYPE), 0, NULL);

    callback_data.fci = &fci;
    callback_data.fci_cache = &fci_cache;

    load_sheet_current_row_data_callback(zs_sheet_name, obj->read_ptr.file_t, &callback_data);
}
/* }}} */

#endif

/** {{{ xls_methods
*/
zend_function_entry xls_methods[] = {
        PHP_ME(vtiful_xls, __construct,    xls_construct_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, close,          xls_close_arginfo,          ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, fileName,       xls_file_name_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, addSheet,       xls_file_add_sheet,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, existSheet,     xls_file_exist_sheet,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, checkoutSheet,  xls_file_checkout_sheet,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, activateSheet,  xls_file_activate_sheet,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, constMemory,    xls_const_memory_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, header,         xls_header_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, data,           xls_data_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, output,         xls_output_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, getHandle,      xls_get_handle_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, autoFilter,     xls_auto_filter_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertText,     xls_insert_text_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertRichText, xls_insert_rtext_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertDate,     xls_insert_date_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertChart,    xls_insert_chart_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertUrl,      xls_insert_url_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertImage,    xls_insert_image_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertFormula,  xls_insert_formula_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, insertComment,  xls_insert_comment_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, showComment,    xls_show_comment_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, mergeCells,     xls_merge_cells_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setColumn,      xls_set_column_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setRow,         xls_set_row_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, getCurrentLine, xls_get_curr_line_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setCurrentLine, xls_set_curr_line_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, defaultFormat,  xls_set_global_format,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, freezePanes,    xls_freeze_panes_arginfo,   ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, protection,    xls_protection_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, validation,    xls_validation_arginfo,     ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, zoom,          xls_sheet_zoom_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, gridline,      xls_sheet_gridline_arginfo, ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, setPaper,     xls_set_paper_arginfo,             ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setMargins,   xls_set_margins_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setPortrait,  xls_set_printed_portrait_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setLandscape, xls_set_printed_landscape_arginfo, ZEND_ACC_PUBLIC)

        PHP_ME(vtiful_xls, setCurrentSheetHide,    xls_hide_sheet_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_xls, setCurrentSheetIsFirst, xls_first_sheet_arginfo, ZEND_ACC_PUBLIC)

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
        PHP_ME(vtiful_xls, setGlobalType,    xls_set_global_type_arginfo,    ZEND_ACC_PUBLIC)
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
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_HIDDEN_ROW,  XLSXIOREAD_SKIP_HIDDEN_ROWS);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_EMPTY_CELLS, XLSXIOREAD_SKIP_EMPTY_CELLS);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_SKIP_EMPTY_VALUE, SKIP_EMPTY_VALUE);
#endif

    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_HIDE_ALL",    LXW_HIDE_ALL_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_ALL",    LXW_SHOW_ALL_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_PRINT",  LXW_SHOW_PRINT_GRIDLINES)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "GRIDLINES_SHOW_SCREEN", LXW_SHOW_SCREEN_GRIDLINES)

    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_DEFAULT",      0)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_LETTER",       1)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_LETTER_SMALL", 2)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_TABLOID",      3)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_LEDGER",       4)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_LEGAL",        5)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_STATEMENT",    6)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_EXECUTIVE",    7)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_A3",           8)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_A4",           9)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_A4_SMALL",     10)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_A5",           11)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_B4",           12)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_B5",           13)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_FOLIO",        14)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_QUARTO",       15)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_NOTE",         18)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_9",   19)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_10",  20)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_11",  21)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_12",  22)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_14",  23)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_C_SIZE_SHEET", 24)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_D_SIZE_SHEET", 25)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_E_SIZE_SHEET", 26)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_DL",  27)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_C3",  28)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_C4",  29)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_C5",  30)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_C6",  31)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_C65", 32)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_B4",  33)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_B5",  34)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_B6",  35)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_1",   36)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_MONARCH",      37)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_ENVELOPE_2",   38)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_FANFOLD",      39)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_GERMAN_STD_FANFOLD",   40)
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, "PAPER_GERMAN_LEGAL_FANFOLD", 41)

    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_INT,      READ_TYPE_INT);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_DOUBLE,   READ_TYPE_DOUBLE);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_STRING,   READ_TYPE_STRING);
    REGISTER_CLASS_CONST_LONG(vtiful_xls_ce, V_XLS_CONST_READ_TYPE_DATETIME, READ_TYPE_DATETIME);

    return SUCCESS;
}
/* }}} */



