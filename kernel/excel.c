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

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "zend.h"
#include "zend_API.h"
#include "zend_exceptions.h"

#include "php.h"

#include "excel.h"
#include "exception.h"
#include "write.h"

zend_class_entry *vtiful_excel_ce;

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

ZEND_BEGIN_ARG_INFO_EX(excel_insert_text_arginfo, 0, 0, 3)
                ZEND_ARG_INFO(0, row)
                ZEND_ARG_INFO(0, column)
                ZEND_ARG_INFO(0, data)
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

excel_resource_t * zval_get_resource(zval *handle)
{
    excel_resource_t *res;

    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    return res;
}

/* {{{ \Vtiful\Kernel\Excel::__construct(array $config)
 */
PHP_METHOD(vtiful_excel, __construct)
{
    zval *config;
    zend_string *key;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(config)
    ZEND_PARSE_PARAMETERS_END();

    key = zend_string_init(V_EXCEL_PAT, sizeof(V_EXCEL_PAT)-1, 0);

    if(zend_hash_find(Z_ARRVAL_P(config), key) == NULL)
    {
        zend_throw_exception(vtiful_exception_ce, "Lack of 'path' configuration", 110);
    }

    zend_update_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_COF), config);

    zend_string_release(key);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::filename(string $fileName)
 */
PHP_METHOD(vtiful_excel, fileName)
{
    zval rv, tmp_file_name, file_path, handle, *config, *tmp_path;
    zend_string *file_name, *key;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(file_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    key      = zend_string_init(V_EXCEL_PAT, sizeof(V_EXCEL_PAT)-1, 0);
    config   = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_COF), 0, &rv TSRMLS_DC);
    tmp_path = zend_hash_find(Z_ARRVAL_P(config), key);

    zend_string_release(key);

    if(!tmp_path && Z_TYPE_P(tmp_path) != IS_STRING)
    {
        zend_throw_exception(vtiful_exception_ce, "Configure 'path' must be a string type", 120);
    }

    ZVAL_STR(&tmp_file_name, file_name);
    concat_function(&file_path, tmp_path, &tmp_file_name);

    res = emalloc(sizeof(excel_resource_t));
    res->workbook  = workbook_new(ZSTR_VAL(zval_get_string(&file_path)));
    res->worksheet = workbook_add_worksheet(res->workbook, NULL);
    zval_ptr_dtor(&file_path);

    ZVAL_RES(&handle, zend_register_resource(res, le_vtiful));

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_FIL), &file_path);
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &handle);
    zval_ptr_dtor(&file_path);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::header(array $header)
 */
PHP_METHOD(vtiful_excel, header)
{
    zval rv, res_handle, *attr_handle, *header, *header_value;
    zend_long header_l_key;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(header)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(attr_handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(header), header_l_key, header_value) {
         type_writer(header_value, 0, header_l_key, res);
         zval_ptr_dtor(header_value);
    } ZEND_HASH_FOREACH_END();

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_excel, data)
{
    zval rv, *data, *attr_handle, res_handle, *data_r_value, *data_l_value;
    zend_long data_r_key, data_l_key;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(attr_handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data), data_r_key, data_r_value) {
        if(Z_TYPE_P(data_r_value) == IS_ARRAY) {
            ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data_r_value), data_l_key, data_l_value) {
                type_writer(data_l_value, data_r_key+1, data_l_key, res);
                zval_ptr_dtor(data_l_value);
            } ZEND_HASH_FOREACH_END();
        }
    } ZEND_HASH_FOREACH_END();

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_excel, output)
{
    zval rv, *handle, null_handle;
    excel_resource_t *res;

    handle = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);

    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    workbook_file(res, handle);

    efree(res);

    ZVAL_NULL(&null_handle);
    zend_update_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), &null_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::getHandle()
 */
PHP_METHOD(vtiful_excel, getHandle)
{
    zval rv;
    zval *handle;

    handle = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);

    ZVAL_COPY(return_value, handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::insertText(int $row, int $column, string|int|double $data)
 */
PHP_METHOD(vtiful_excel, insertText)
{
    zval rv, res_handle;
    zval *attr_handle, *data;
    zend_long row, column;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(attr_handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    type_writer(data, row, column, res);

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertImage)
{
    zval rv, res_handle;
    zval *attr_handle, *image;
    zend_long row, column;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(image)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(attr_handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    image_writer(image, row, column, res);

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::insertImage(int $row, int $column, string $imagePath)
 */
PHP_METHOD(vtiful_excel, insertFormula)
{
    zval rv, res_handle;
    zval *attr_handle, *formula;
    zend_long row, column;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(3, 3)
            Z_PARAM_LONG(row)
            Z_PARAM_LONG(column)
            Z_PARAM_ZVAL(formula)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(attr_handle), VTIFUL_RESOURCE_NAME, le_vtiful)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    formula_writer(formula, row, column, res);

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::autoFilter(int $rowStart, int $rowEnd, int $columnStart, int $columnEnd)
 */
PHP_METHOD(vtiful_excel, autoFilter)
{
    zval rv, res_handle;
    zval *attr_handle;
    zend_string *range;
    excel_resource_t *res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(range)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    attr_handle = zend_read_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), 0, &rv TSRMLS_DC);
    res = zval_get_resource(attr_handle);

    auto_filter(range, res);

    ZVAL_RES(&res_handle, zend_register_resource(res, le_vtiful));
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &res_handle);
}
/* }}} */

zend_function_entry excel_methods[] = {
        PHP_ME(vtiful_excel, __construct, excel_construct_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_CTOR)
        PHP_ME(vtiful_excel, fileName,    excel_file_name_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, header,      excel_header_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, data,        excel_data_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, output,      NULL,                    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, getHandle,   NULL,                    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, autoFilter,    excel_auto_filter_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertText,    excel_insert_text_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertImage,   excel_insert_image_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, insertFormula, excel_insert_formula_arginfo, ZEND_ACC_PUBLIC)
        PHP_FE_END
};

VTIFUL_STARTUP_FUNCTION(excel) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Excel", excel_methods);

    vtiful_excel_ce = zend_register_internal_class(&ce);

    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_COF), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_FIL), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_HANDLE), ZEND_ACC_PRIVATE);

    return SUCCESS;
}



