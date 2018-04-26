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

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "include.h"

zend_class_entry *vtiful_excel_ce;

static zend_object_handlers vtiful_excel_handlers;

static zend_always_inline void *vtiful_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size + zend_object_properties_size(ce));
    memset(obj, 0, obj_size);
    return obj;
}

/* {{{ excel_objects_new
 */
PHP_VTIFUL_API zend_object *excel_objects_new(zend_class_entry *ce)
{
    excel_object *intern = vtiful_object_alloc(sizeof(excel_object), ce);

    zend_object_std_init(&intern->zo, ce);
    object_properties_init(&intern->zo, ce);

    intern->zo.handlers = &vtiful_excel_handlers;

    return &intern->zo;
}
/* }}} */

/* {{{ vtiful_excel_objects_free
 */
static void vtiful_excel_objects_free(zend_object *object)
{
    excel_object *intern = php_vtiful_excel_fetch_object(object);

    lxw_workbook_free(intern->ptr.workbook);

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(excel_construct_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, config)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_file_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, file_name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_header_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, header)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_data_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, data)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_insert_text_arginfo, 0, 0, 4)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, data)
                ZEND_ARG_INFO(0, format)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_insert_image_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, image)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_insert_formula_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, formula)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_auto_filter_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, range)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_merge_cells_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, data)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_set_column_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, format_handle)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, width)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(excel_set_row_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, format_handle)
                ZEND_ARG_INFO(0, range)
                ZEND_ARG_INFO(0, height)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::__construct(array $config)
 */
PHP_METHOD(vtiful_excel, __construct)
{
    zval *config, *c_path;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(config)
    ZEND_PARSE_PARAMETERS_END();

    if((c_path = zend_hash_str_find(Z_ARRVAL_P(config), ZEND_STRL(V_EXCEL_PAT))) == NULL)
    {
        zend_throw_exception(vtiful_exception_ce, "Lack of 'path' configuration", 110);
        return;
    }

    if(Z_TYPE_P(c_path) != IS_STRING)
    {
        zend_throw_exception(vtiful_exception_ce, "Configure 'path' must be a string type", 120);
        return;
    }

    add_property_zval(getThis(), V_EXCEL_COF, config);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::filename(string $fileName)
 */
PHP_METHOD(vtiful_excel, fileName)
{
    zval file_path, *dir_path;
    zend_string *file_name;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(file_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_excel_ce, return_value);

    excel_object *obj = Z_EXCEL_P(getThis());

    if(obj->ptr.workbook == NULL) {
        excel_file_path(file_name, dir_path, &file_path);

        obj->ptr.workbook  = workbook_new(Z_STRVAL(file_path));
        obj->ptr.worksheet = workbook_add_worksheet(obj->ptr.workbook, NULL);

        add_property_zval(return_value, V_EXCEL_FIL, &file_path);

        zval_ptr_dtor(&file_path);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::constMemory(string $fileName)
 */
PHP_METHOD(vtiful_excel, constMemory)
{
    zval file_path, *dir_path;
    zend_string *file_name;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(file_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    GET_CONFIG_PATH(dir_path, vtiful_excel_ce, return_value);

    excel_object *obj = Z_EXCEL_P(getThis());

    if(obj->ptr.workbook == NULL) {
        excel_file_path(file_name, dir_path, &file_path);

        lxw_workbook_options options = {.constant_memory = LXW_TRUE, .tmpdir = NULL};

        obj->ptr.workbook  = workbook_new_opt(Z_STRVAL(file_path), &options);
        obj->ptr.worksheet = workbook_add_worksheet(obj->ptr.workbook, NULL);

        add_property_zval(return_value, V_EXCEL_FIL, &file_path);

        zval_ptr_dtor(&file_path);
    }
}
/* }}} */


/** {{{ \Vtiful\Kernel\Excel::header(array $header)
 */
PHP_METHOD(vtiful_excel, header)
{
    zval *header, *header_value;
    zend_long header_l_key;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(header)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(header), header_l_key, header_value) {
         type_writer(header_value, 0, header_l_key, &obj->ptr, NULL);
         zval_ptr_dtor(header_value);
    } ZEND_HASH_FOREACH_END();
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_excel, data)
{
    zval *data, *data_r_value, *data_l_value;
    zend_long data_r_key, data_l_key;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data), data_r_key, data_r_value) {
        if(Z_TYPE_P(data_r_value) == IS_ARRAY) {
            ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data_r_value), data_l_key, data_l_value) {
                type_writer(data_l_value, data_r_key+1, data_l_key, &obj->ptr, NULL);
                zval_ptr_dtor(data_l_value);
            } ZEND_HASH_FOREACH_END();
        }
    } ZEND_HASH_FOREACH_END();
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_excel, output)
{
    zval rv, *file_path;

    file_path = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_FIL), 0, &rv TSRMLS_DC);

    excel_object *obj = Z_EXCEL_P(getThis());

    workbook_file(&obj->ptr);

    add_property_null(getThis(), V_EXCEL_HANDLE);
    add_property_null(getThis(), V_EXCEL_PAT);

    ZVAL_COPY(return_value, file_path);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::getHandle()
 */
PHP_METHOD(vtiful_excel, getHandle)
{
    excel_object *obj = Z_EXCEL_P(getThis());

    RETURN_RES(zend_register_resource(&obj->ptr, le_excel_writer));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertText(int $row, int $column, string|int|double $data)
 */
PHP_METHOD(vtiful_excel, insertText)
{
    zval *data;
    zend_long row, column;
    zend_string *format = NULL;

    ZEND_PARSE_PARAMETERS_START(3, 4)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(data)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(format)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    type_writer(data, row, column, &obj->ptr, format);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertImage)
{
    zval *image;
    zend_long row, column;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(image)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    image_writer(image, row, column, &obj->ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertFormula)
{
    zval *formula;
    zend_long row, column;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(formula)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    formula_writer(formula, row, column, &obj->ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::autoFilter(int $rowStart, int $rowEnd, int $columnStart, int $columnEnd)
 */
PHP_METHOD(vtiful_excel, autoFilter)
{
    zend_string *range;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(range)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    auto_filter(range, &obj->ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::mergeCells(string $range, string $data)
 */
PHP_METHOD(vtiful_excel, mergeCells)
{
    zval res_handle;
    zend_string *range, *data;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_STR(range)
            Z_PARAM_STR(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    excel_object *obj = Z_EXCEL_P(getThis());

    merge_cells(range, data, &obj->ptr);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setColumn(resource $format, string $range [, int $width])
 */
PHP_METHOD(vtiful_excel, setColumn)
{
    zval *format_handle, res_handle;
    zend_string *range;

    double width = 0;
    int    argc  = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(2, 3)
            Z_PARAM_RESOURCE(format_handle)
            Z_PARAM_STR(range)
            Z_PARAM_OPTIONAL
            Z_PARAM_DOUBLE(width)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    if (argc == 2) {
        width = 10;
    }

    excel_object *obj = Z_EXCEL_P(getThis());

    set_column(range, width, &obj->ptr, zval_get_format(format_handle));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::setRow(resource $format, string $range [, int $heitght])
 */
PHP_METHOD(vtiful_excel, setRow)
{
    zval *format_handle, res_handle;
    zend_string *range;

    double height = 0;
    int    argc  = ZEND_NUM_ARGS();

    ZEND_PARSE_PARAMETERS_START(2, 3)
            Z_PARAM_RESOURCE(format_handle)
            Z_PARAM_STR(range)
            Z_PARAM_OPTIONAL
            Z_PARAM_DOUBLE(height)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    if (argc == 2) {
        height = 18;
    }

    excel_object *obj = Z_EXCEL_P(getThis());

    set_row(range, height, &obj->ptr, zval_get_format(format_handle));
}
/* }}} */

/** {{{ excel_methods
*/
zend_function_entry excel_methods[] = {
        PHP_ME(vtiful_excel, __construct,   excel_construct_arginfo,      ZEND_ACC_PUBLIC|ZEND_ACC_CTOR)
        PHP_ME(vtiful_excel, fileName,      excel_file_name_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, constMemory,   excel_file_name_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, header,        excel_header_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, data,          excel_data_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, output,        NULL,                         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, getHandle,     NULL,                         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, autoFilter,    excel_auto_filter_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertText,    excel_insert_text_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertImage,   excel_insert_image_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertFormula, excel_insert_formula_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, mergeCells,    excel_merge_cells_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, setColumn,     excel_set_column_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, setRow,        excel_set_row_arginfo,        ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(excel) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Excel", excel_methods);
    ce.create_object = excel_objects_new;
    vtiful_excel_ce = zend_register_internal_class(&ce);

    memcpy(&vtiful_excel_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    vtiful_excel_handlers.offset   = XtOffsetOf(excel_object, zo);
    vtiful_excel_handlers.free_obj = vtiful_excel_objects_free;

    REGISTER_CLASS_PROPERTY_NULL(vtiful_excel_ce, V_EXCEL_COF, ZEND_ACC_PRIVATE);
    REGISTER_CLASS_PROPERTY_NULL(vtiful_excel_ce, V_EXCEL_FIL, ZEND_ACC_PRIVATE);

    return SUCCESS;
}
/* }}} */



