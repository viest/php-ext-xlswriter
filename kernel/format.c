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

zend_class_entry *vtiful_format_ce;

/* {{{ format_objects_new
 */
static zend_object_handlers format_handlers;

static zend_always_inline void *vtiful_format_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *format_objects_new(zend_class_entry *ce)
{
    format_object *format = vtiful_format_object_alloc(sizeof(format_object), ce);

    zend_object_std_init(&format->zo, ce);
    object_properties_init(&format->zo, ce);

    format->ptr.format  = NULL;
    format->zo.handlers = &format_handlers;

    return &format->zo;
}
/* }}} */

/* {{{ format_objects_free
 */
static void format_objects_free(zend_object *object)
{
    format_object *intern = php_vtiful_format_fetch_object(object);

    if (intern->ptr.format != NULL) {
        // free by workbook
        intern->ptr.format = NULL;
    }

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(format_construct_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, handle)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_underline_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_align_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_color_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, color)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Format::__construct()
 */
PHP_METHOD(vtiful_format, __construct)
{
    zval *handle;
    format_object *obj;
    xls_resource_t *xls_res;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_RESOURCE(handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_res = zval_get_resource(handle);
    obj     = Z_FORMAT_P(getThis());

    if (obj->ptr.format == NULL) {
        obj->ptr.format = workbook_add_format(xls_res->workbook);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::bold()
 */
PHP_METHOD(vtiful_format, bold)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());
    format_set_bold(obj->ptr.format);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::italic()
 */
PHP_METHOD(vtiful_format, italic)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());
    format_set_italic(obj->ptr.format);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::underline()
 */
PHP_METHOD(vtiful_format, underline)
{
    zend_long style;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(style)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());
    format_set_underline(obj->ptr.format, style);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::align()
 */
PHP_METHOD(vtiful_format, align)
{
    zval *args = NULL;
    int argc, i;

    ZEND_PARSE_PARAMETERS_START(1, -1)
            Z_PARAM_VARIADIC('+', args, argc)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    for (i = 0; i < argc; ++i) {
        zval *arg = args + i;

        if (Z_TYPE_P(arg) != IS_LONG) {
            zend_throw_exception(vtiful_exception_ce, "Format exception, please view the manual", 150);
        }

        format_set_align(obj->ptr.format, Z_LVAL_P(arg));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::color(int $color)
 */
PHP_METHOD(vtiful_format, color)
{
    zend_long color;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(color)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    format_set_font_color(obj->ptr.format, color);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::toResource()
 */
PHP_METHOD(vtiful_format, toResource)
{
    format_object *obj = Z_FORMAT_P(getThis());

    RETURN_RES(zend_register_resource(obj->ptr.format, le_xls_writer));
}
/* }}} */


/** {{{ format_methods
*/
zend_function_entry format_methods[] = {
        PHP_ME(vtiful_format, __construct, format_construct_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, bold,        NULL,                     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, italic,      NULL,                     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, underline,   format_underline_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, align,       format_align_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, color,       format_color_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, toResource,  NULL,                     ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(format) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Format", format_methods);
    ce.create_object = format_objects_new;
    vtiful_format_ce = zend_register_internal_class(&ce);

    memcpy(&format_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    format_handlers.offset   = XtOffsetOf(format_object, zo);
    format_handlers.free_obj = format_objects_free;

    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "UNDERLINE_SINGLE",            LXW_UNDERLINE_SINGLE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "UNDERLINE_DOUBLE ",           LXW_UNDERLINE_DOUBLE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "UNDERLINE_SINGLE_ACCOUNTING", LXW_UNDERLINE_SINGLE_ACCOUNTING)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "UNDERLINE_DOUBLE_ACCOUNTING", LXW_UNDERLINE_DOUBLE_ACCOUNTING)

    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_LEFT",                 LXW_ALIGN_LEFT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_CENTER",               LXW_ALIGN_CENTER)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_RIGHT",                LXW_ALIGN_RIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_FILL",                 LXW_ALIGN_FILL)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_JUSTIFY",              LXW_ALIGN_JUSTIFY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_CENTER_ACROSS",        LXW_ALIGN_CENTER_ACROSS)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_DISTRIBUTED",          LXW_ALIGN_DISTRIBUTED)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_VERTICAL_TOP",         LXW_ALIGN_VERTICAL_TOP)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_VERTICAL_BOTTOM",      LXW_ALIGN_VERTICAL_BOTTOM)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_VERTICAL_CENTER",      LXW_ALIGN_VERTICAL_CENTER)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_VERTICAL_JUSTIFY",     LXW_ALIGN_VERTICAL_JUSTIFY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "FORMAT_ALIGN_VERTICAL_DISTRIBUTED", LXW_ALIGN_VERTICAL_DISTRIBUTED)

    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_BLACK",   LXW_COLOR_BLACK)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_BLUE",    LXW_COLOR_BLUE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_BROWN",   LXW_COLOR_BROWN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_CYAN",    LXW_COLOR_CYAN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_GRAY",    LXW_COLOR_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_GREEN",   LXW_COLOR_GREEN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_LIME",    LXW_COLOR_LIME)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_MAGENTA", LXW_COLOR_MAGENTA)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_NAVY",    LXW_COLOR_NAVY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_ORANGE",  LXW_COLOR_ORANGE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_PINK",    LXW_COLOR_PINK)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_PURPLE",  LXW_COLOR_PURPLE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_RED",     LXW_COLOR_RED)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_SILVER",  LXW_COLOR_SILVER)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_WHITE",   LXW_COLOR_WHITE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "COLOR_YELLOW",  LXW_COLOR_YELLOW)

    return SUCCESS;
}
/* }}} */
