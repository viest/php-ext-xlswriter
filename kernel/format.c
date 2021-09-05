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

ZEND_BEGIN_ARG_INFO_EX(format_wrap_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_bold_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_italic_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_underline_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_unlocked_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_align_arginfo, 0, 0, 1)
                ZEND_ARG_VARIADIC_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_color_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, color)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_size_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, size)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_strikeout_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_number_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, format)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_background_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, pattern)
                ZEND_ARG_INFO(0, color)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_border_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_font_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, font)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(format_to_resource_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Format::__construct()
 */
PHP_METHOD(vtiful_format, __construct)
{
    zval *handle;
    format_object *obj;
    xls_resource_write_t *xls_res;

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

    if (obj->ptr.format) {
        format_set_bold(obj->ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::italic()
 */
PHP_METHOD(vtiful_format, italic)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_italic(obj->ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::underline()
 */
PHP_METHOD(vtiful_format, underline)
{
    zend_long style = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(style)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_underline(obj->ptr.format, style);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::unlocked()
 */
PHP_METHOD(vtiful_format, unlocked)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_unlocked(obj->ptr.format);
    }
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

        if (obj->ptr.format) {
            format_set_align(obj->ptr.format, Z_LVAL_P(arg));
        }
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::fontColor(int $color)
 */
PHP_METHOD(vtiful_format, fontColor)
{
    zend_long color = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(color)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_font_color(obj->ptr.format, color);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::number(string $format)
 */
PHP_METHOD(vtiful_format, number)
{
    zend_string *format;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(format)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_num_format(obj->ptr.format, ZSTR_VAL(format));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::background(int $color [, int $pattern = \Vtiful\Kernel\Format::PATTERN_SOLID])
 */
PHP_METHOD(vtiful_format, background)
{
    zend_long pattern = LXW_PATTERN_SOLID, color = 0;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_LONG(color)
            Z_PARAM_OPTIONAL
            Z_PARAM_LONG(pattern)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_pattern(obj->ptr.format, pattern);
        format_set_bg_color(obj->ptr.format, color);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::fontSize(double $size)
 */
PHP_METHOD(vtiful_format, fontSize)
{
    double size;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_DOUBLE(size)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_font_size(obj->ptr.format, size);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::font(string $fontName)
 */
PHP_METHOD(vtiful_format, font)
{
    zend_string *font_name = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(font_name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_font_name(obj->ptr.format, ZSTR_VAL(font_name));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::strikeout()
 */
PHP_METHOD(vtiful_format, strikeout)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_font_strikeout(obj->ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::wrap()
 */
PHP_METHOD(vtiful_format, wrap)
{
    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_text_wrap(obj->ptr.format);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Format::toResource()
 */
PHP_METHOD(vtiful_format, border)
{
    zend_long style = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(style)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    format_object *obj = Z_FORMAT_P(getThis());

    if (obj->ptr.format) {
        format_set_border(obj->ptr.format, style);
    }
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
        PHP_ME(vtiful_format, __construct,   format_construct_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, wrap,          format_wrap_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, bold,          format_bold_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, italic,        format_italic_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, border,        format_border_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, align,         format_align_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, number,        format_number_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, fontColor,     format_color_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, font,          format_font_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, fontSize,      format_size_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, strikeout,     format_strikeout_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, underline,     format_underline_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, unlocked,      format_unlocked_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, toResource,    format_to_resource_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_format, background,    format_background_arginfo,  ZEND_ACC_PUBLIC)
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

    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_NONE",             LXW_PATTERN_NONE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_SOLID",            LXW_PATTERN_SOLID)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_MEDIUM_GRAY",      LXW_PATTERN_MEDIUM_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_GRAY",        LXW_PATTERN_DARK_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_GRAY",       LXW_PATTERN_LIGHT_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_HORIZONTAL",  LXW_PATTERN_DARK_HORIZONTAL)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_VERTICAL",    LXW_PATTERN_DARK_VERTICAL)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_DOWN",        LXW_PATTERN_DARK_DOWN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_UP",          LXW_PATTERN_DARK_UP)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_GRID",        LXW_PATTERN_DARK_GRID)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_DARK_TRELLIS",     LXW_PATTERN_DARK_TRELLIS)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_HORIZONTAL", LXW_PATTERN_LIGHT_HORIZONTAL)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_VERTICAL",   LXW_PATTERN_LIGHT_VERTICAL)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_DOWN",       LXW_PATTERN_LIGHT_DOWN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_UP",         LXW_PATTERN_LIGHT_UP)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_GRID",       LXW_PATTERN_LIGHT_GRID)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_LIGHT_TRELLIS",    LXW_PATTERN_LIGHT_TRELLIS)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_GRAY_125",         LXW_PATTERN_GRAY_125)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "PATTERN_GRAY_0625",        LXW_PATTERN_GRAY_0625)

    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_NONE",                 LXW_BORDER_NONE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_THIN",                 LXW_BORDER_THIN)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_MEDIUM",               LXW_BORDER_MEDIUM)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_DASHED",               LXW_BORDER_DASHED)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_DOTTED",               LXW_BORDER_DOTTED)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_THICK",                LXW_BORDER_THICK)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_DOUBLE",               LXW_BORDER_DOUBLE)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_HAIR",                 LXW_BORDER_HAIR)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_MEDIUM_DASHED",        LXW_BORDER_MEDIUM_DASHED)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_DASH_DOT",             LXW_BORDER_DASH_DOT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_MEDIUM_DASH_DOT",      LXW_BORDER_MEDIUM_DASH_DOT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_DASH_DOT_DOT",         LXW_BORDER_DASH_DOT_DOT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_MEDIUM_DASH_DOT_DOT",  LXW_BORDER_MEDIUM_DASH_DOT_DOT)
    REGISTER_CLASS_CONST_LONG(vtiful_format_ce, "BORDER_SLANT_DASH_DOT",       LXW_BORDER_SLANT_DASH_DOT)

    return SUCCESS;
}
/* }}} */
