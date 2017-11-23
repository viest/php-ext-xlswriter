#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "zend.h"
#include "zend_API.h"
#include "zend_exceptions.h"

#include "php.h"

#include "xlsxwriter.h"
#include "php_vtiful.h"
#include "excel.h"
#include "exception.h"

#include "ext/standard/php_var.h"

typedef struct {
    lxw_workbook *workbook;
    lxw_worksheet *worksheet;
} excel_resource_t;

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

    res = malloc(sizeof(excel_resource_t));
    res->workbook  = workbook_new(ZSTR_VAL(zval_get_string(&file_path)));
    res->worksheet = workbook_add_worksheet(res->workbook, NULL);

    ZVAL_RES(&handle, zend_register_resource(res, le_vtiful));

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_FIL), &file_path);
    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HANDLE), &handle);

    zval_ptr_dtor(&file_path);
    zval_ptr_dtor(&file_path);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::header(array $header)
 */
PHP_METHOD(vtiful_excel, header)
{
    zval *header;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(header)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_HEADER), header);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::data(array $data)
 */
PHP_METHOD(vtiful_excel, data)
{
    zval *data;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(data)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_DATA), data);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::output()
 */
PHP_METHOD(vtiful_excel, output)
{
    zval rv1, rv2, rv3;
    zval *file_name, *header, *data, *value, *data_r_value, *data_l_value;
    zend_long header_l_key, data_r_key, data_l_key;
    excel_resource_t *res;

    file_name = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_FIL), 0, &rv1 TSRMLS_DC);
    header    = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_HEADER), 0, &rv2 TSRMLS_DC);
    data      = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_DATA), 0, &rv3 TSRMLS_DC);

    res = malloc(sizeof(excel_resource_t));

    res->workbook  = workbook_new(ZSTR_VAL(zval_get_string(file_name)));
    res->worksheet = workbook_add_worksheet(res->workbook, NULL);

    zval_ptr_dtor(file_name);

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(header), header_l_key, value) {
        worksheet_write_string(res->worksheet, 0, header_l_key, ZSTR_VAL(zval_get_string(value)), NULL);
        zval_ptr_dtor(value);
    } ZEND_HASH_FOREACH_END();

    ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data), data_r_key, data_r_value) {
        ZEND_HASH_FOREACH_NUM_KEY_VAL(Z_ARRVAL_P(data_r_value), data_l_key, data_l_value) {
            switch (Z_TYPE_P(data_l_value)) {
                case IS_STRING:
                    worksheet_write_string(res->worksheet, data_r_key+1, data_l_key, ZSTR_VAL(zval_get_string(data_l_value)), NULL);
                    zval_ptr_dtor(data_l_value);
                    break;
                case IS_LONG:
                    worksheet_write_number(res->worksheet, data_r_key+1, data_l_key, zval_get_long(data_l_value), NULL);
                    break;
            }
        } ZEND_HASH_FOREACH_END();
    } ZEND_HASH_FOREACH_END();

    workbook_close(res->workbook);
}
/* }}} */

/* {{{ \Vtiful\Kernel\Excel::getHandle()
 */
PHP_METHOD(vtiful_excel, getHandle)
{
    zval rv;
    zval *file_name;
    excel_resource_t *res;

    file_name = zend_read_property(vtiful_excel_ce, getThis(), ZEND_STRL(V_EXCEL_FIL), 0, &rv TSRMLS_DC);

    res = malloc(sizeof(excel_resource_t));

    res->workbook  = workbook_new(ZSTR_VAL(zval_get_string(file_name)));
    res->worksheet = workbook_add_worksheet(res->workbook, NULL);

    zval_ptr_dtor(file_name);

    RETURN_RES(zend_register_resource(res, le_vtiful));
}
/* }}} */

zend_function_entry excel_methods[] = {
        PHP_ME(vtiful_excel, __construct, excel_construct_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_CTOR)
        PHP_ME(vtiful_excel, fileName,    excel_file_name_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, header,      excel_header_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, data,        excel_data_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, output,      NULL,                    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_excel, getHandle,   NULL,                    ZEND_ACC_PUBLIC)
        PHP_FE_END
};

VTIFUL_STARTUP_FUNCTION(excel) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Excel", excel_methods);

    vtiful_excel_ce = zend_register_internal_class(&ce);

    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_COF), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_FIL), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_DATA), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_HEADER), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_HANDLE), ZEND_ACC_PRIVATE);

    return SUCCESS;
}



