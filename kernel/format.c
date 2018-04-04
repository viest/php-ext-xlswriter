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

ZEND_BEGIN_ARG_INFO_EX(format_underline_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, handle)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Format::bold()
 */
PHP_METHOD(vtiful_format, bold)
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

/** {{{ \Vtiful\Kernel\Format::italic()
 */
PHP_METHOD(vtiful_format, italic)
{
    zval *handle;
    lxw_format *italic_format;
    excel_resource_t *excel_res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_RESOURCE(handle)
    ZEND_PARSE_PARAMETERS_END();

    excel_res   = zval_get_resource(handle);
    italic_format = workbook_add_format(excel_res->workbook);

    format_set_italic(italic_format);

    RETURN_RES(zend_register_resource(italic_format, le_excel_writer));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::underline()
 */
PHP_METHOD(vtiful_format, underline)
{
    zval *handle;
    zend_long style;
    lxw_format *underline_format;
    excel_resource_t *excel_res;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_RESOURCE(handle)
            Z_PARAM_LONG(style)
    ZEND_PARSE_PARAMETERS_END();

    excel_res = zval_get_resource(handle);
    underline_format = workbook_add_format(excel_res->workbook);

    format_set_underline(underline_format, style);

    RETURN_RES(zend_register_resource(underline_format, le_excel_writer));
}
/* }}} */


/** {{{ excel_methods
*/
zend_function_entry format_methods[] = {
        PHP_ME(vtiful_format, bold,      format_style_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_ME(vtiful_format, italic,    format_style_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_ME(vtiful_format, underline, format_style_arginfo, ZEND_ACC_PUBLIC|ZEND_ACC_STATIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(format) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Format", format_methods);

//    zend_declare_class_constant_long(&ce, "UNDERLINE_SINGLE", sizeof("UNDERLINE_SINGLE")-1, (zend_long)LXW_UNDERLINE_SINGLE);

    vtiful_format_ce = zend_register_internal_class(&ce);

    return SUCCESS;
}
/* }}} */
