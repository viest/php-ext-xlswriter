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

zend_class_entry *vtiful_format_ce;

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(format_style_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, handle)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Format::boldStype()
 */
PHP_METHOD(vtiful_format, boldStype)
{
    zval *handle;
    lxw_format *bold_format;
    excel_resource_t *excel_res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_RESOURCE(handle)
    ZEND_PARSE_PARAMETERS_END();

    excel_res   = zval_get_resource(handle);
    bold_format = workbook_add_format(excel_res->workbook);

    format_set_bold(bold_format);

    RETURN_RES(zend_register_resource(bold_format, le_excel_writer));
}
/* }}} */

/** {{{ excel_methods
*/
zend_function_entry formac_methods[] = {
        PHP_ME(vtiful_format, boldStype, format_style_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(format) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Format", formac_methods);

    vtiful_format_ce = zend_register_internal_class(&ce);

    return SUCCESS;
}
/* }}} */
