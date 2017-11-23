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
    zval rv, tmp_file_name, *config, *tmp_path, file_path;
    zend_string *file_name, *key;

    return_value = getThis();

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(file_name)
    ZEND_PARSE_PARAMETERS_END();

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

    zend_update_property(vtiful_excel_ce, return_value, ZEND_STRL(V_EXCEL_FIL), &file_path);
    zval_ptr_dtor(&file_path);
}
/* }}} */

zend_function_entry excel_methods[] = {
        PHP_ME(vtiful_excel, __construct, excel_construct_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_CTOR)
        PHP_ME(vtiful_excel, fileName,    excel_file_name_arginfo, ZEND_ACC_PUBLIC)
        PHP_FE_END
};

VTIFUL_STARTUP_FUNCTION(excel) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Excel", excel_methods);

    vtiful_excel_ce = zend_register_internal_class(&ce);

    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_FIL), ZEND_ACC_PRIVATE);
    zend_declare_property_null(vtiful_excel_ce, ZEND_STRL(V_EXCEL_COF), ZEND_ACC_PRIVATE);

    return SUCCESS;
}



