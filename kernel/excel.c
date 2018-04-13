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

    if((c_path = zend_hash_str_find(Z_ARRVAL_P(config), V_EXCEL_PAT, sizeof(V_EXCEL_PAT)-1)) == NULL)
    {
        zend_throw_exception(vtiful_exception_ce, "Lack of 'path' configuration", 110);
        return;
    }

    if(Z_TYPE_P(c_path) != IS_STRING)
    {
        zend_throw_exception(vtiful_exception_ce, "Configure 'path' must be a string type", 120);
        return;
    }

    zend_update_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_COF), config);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::filename(string $fileName)
 */
PHP_METHOD(vtiful_excel, fileName)
{
    zval rv, file_path, handle, *config, *tmp_path;
    zend_string *file_name, *full_path, *zstr_path;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(file_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    config   = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_COF), 0, &rv TSRMLS_DC);
    tmp_path = zend_hash_str_find(Z_ARRVAL_P(config), V_EXCEL_PAT, sizeof(V_EXCEL_PAT)-1);

    zstr_path = zval_get_string(tmp_path);

    if (Z_STRVAL_P(tmp_path)[Z_STRLEN_P(tmp_path)-1] == '/') {
        full_path = zend_string_extend(zstr_path, ZSTR_LEN(zstr_path) + ZSTR_LEN(file_name), 0);
        memcpy(ZSTR_VAL(full_path)+ZSTR_LEN(zstr_path), ZSTR_VAL(file_name), ZSTR_LEN(file_name)+1);
    } else {
        full_path = zend_string_extend(zstr_path, ZSTR_LEN(zstr_path) + ZSTR_LEN(file_name) + 1, 0);
        ZSTR_VAL(full_path)[ZSTR_LEN(zstr_path)] ='/';
        memcpy(ZSTR_VAL(full_path)+ZSTR_LEN(zstr_path)+1, ZSTR_VAL(file_name), ZSTR_LEN(file_name)+1);
    }

    ZVAL_STR(&file_path, full_path);

    excel_res->workbook  = workbook_new(Z_STRVAL(file_path));
    excel_res->worksheet = workbook_add_worksheet(excel_res->workbook, NULL);

    ZVAL_RES(&handle, zend_register_resource(excel_res, le_excel_writer));

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_FIL), &file_path);
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &handle);

    zval_ptr_dtor(&file_path);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::header(array $header)
 */
PHP_METHOD(vtiful_excel, header)
{
    zval res_handle, *header, *header_value;
    zend_long header_l_key;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(header)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(header), header_l_key, header_value) {
         type_writer(header_value, 0, header_l_key, excel_res, NULL);
         zval_ptr_dtor(header_value);
    } ZEND_HASH_FOREACH_END();

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_excel, data)
{
    zval *data, res_handle, *data_r_value, *data_l_value;
    zend_long data_r_key, data_l_key;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data), data_r_key, data_r_value) {
        if(Z_TYPE_P(data_r_value) == IS_ARRAY) {
            ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data_r_value), data_l_key, data_l_value) {
                type_writer(data_l_value, data_r_key+1, data_l_key, excel_res, NULL);
                zval_ptr_dtor(data_l_value);
            } ZEND_HASH_FOREACH_END();
        }
    } ZEND_HASH_FOREACH_END();

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_excel, output)
{
    zval rv, null_handle, *handle, *file_path;

    handle    = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    file_path = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_FIL), 0, &rv TSRMLS_DC);

    workbook_file(excel_res, handle);

    free(excel_res);

    ZVAL_NULL(&null_handle);
    zend_update_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), &null_handle);

    ZVAL_COPY(return_value, file_path);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::getHandle()
 */
PHP_METHOD(vtiful_excel, getHandle)
{
    zval rv;
    zval *handle;

    handle = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);

    ZVAL_COPY(return_value, handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertText(int $row, int $column, string|int|double $data)
 */
PHP_METHOD(vtiful_excel, insertText)
{
    zval res_handle;
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

    type_writer(data, row, column, excel_res, format);

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertImage)
{
    zval res_handle;
    zval *image;
    zend_long row, column;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(image)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    image_writer(image, row, column, excel_res);

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertFormula)
{
    zval res_handle;
    zval *formula;
    zend_long row, column;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(formula)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    formula_writer(formula, row, column, excel_res);

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Excel::autoFilter(int $rowStart, int $rowEnd, int $columnStart, int $columnEnd)
 */
PHP_METHOD(vtiful_excel, autoFilter)
{
    zval res_handle;
    zend_string *range;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(range)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    auto_filter(range, excel_res);

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
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

    merge_cells(range, data, excel_res);

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
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

    set_column(range, width, excel_res, zval_get_format(format_handle));

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
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

    set_row(range, height, excel_res, zval_get_format(format_handle));

    ZVAL_RES(&res_handle, zend_register_resource(excel_res, le_excel_writer));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/** {{{ excel_methods
*/
zend_function_entry excel_methods[] = {
        PHP_ME(vtiful_excel, __construct,   excel_construct_arginfo,      ZEND_ACC_PUBLIC|ZEND_ACC_CTOR)
        PHP_ME(vtiful_excel, fileName,      excel_file_name_arginfo,      ZEND_ACC_PUBLIC)
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

    vtiful_excel_ce = zend_register_internal_class(&ce);

    REGISTER_CLASS_PROPERTY_NULL(vtiful_excel_ce, V_EXCEL_COF,    ZEND_ACC_PRIVATE);
    REGISTER_CLASS_PROPERTY_NULL(vtiful_excel_ce, V_EXCEL_FIL,    ZEND_ACC_PRIVATE);
    REGISTER_CLASS_PROPERTY_NULL(vtiful_excel_ce, V_EXCEL_HANDLE, ZEND_ACC_PRIVATE);

    return SUCCESS;
}
/* }}} */



