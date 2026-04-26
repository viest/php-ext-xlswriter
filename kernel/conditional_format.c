/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension — ConditionalFormat builder.                      |
  +----------------------------------------------------------------------+
  | A thin PHP veneer over libxlsxwriter's lxw_conditional_format struct. |
  | One ConditionalFormat instance per rule; pass it to                   |
  | Excel::conditionalFormatCell()/Range() which copies what it needs and |
  | hands the underlying struct to libxlsxwriter.                         |
  +----------------------------------------------------------------------+
*/

#include "xlswriter.h"

zend_class_entry *vtiful_cond_format_ce;

/* {{{ object lifecycle */

static zend_object_handlers cond_format_handlers;

static zend_always_inline void *vtiful_cond_format_object_alloc(size_t obj_size, zend_class_entry *ce)
{
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *cond_format_objects_new(zend_class_entry *ce)
{
    cond_format_object *intern = vtiful_cond_format_object_alloc(
        sizeof(cond_format_object), ce);

    zend_object_std_init(&intern->zo, ce);
    object_properties_init(&intern->zo, ce);

    intern->ptr.cf = NULL;
    intern->ptr.value_string     = NULL;
    intern->ptr.min_value_string = NULL;
    intern->ptr.mid_value_string = NULL;
    intern->ptr.max_value_string = NULL;
    intern->ptr.multi_range      = NULL;

    intern->zo.handlers = &cond_format_handlers;
    return &intern->zo;
}

static void cond_format_objects_free(zend_object *object)
{
    cond_format_object *intern = php_vtiful_cond_format_fetch_object(object);

    if (intern->ptr.cf) {
        efree(intern->ptr.cf);
        intern->ptr.cf = NULL;
    }
    if (intern->ptr.value_string)     { efree(intern->ptr.value_string);     intern->ptr.value_string     = NULL; }
    if (intern->ptr.min_value_string) { efree(intern->ptr.min_value_string); intern->ptr.min_value_string = NULL; }
    if (intern->ptr.mid_value_string) { efree(intern->ptr.mid_value_string); intern->ptr.mid_value_string = NULL; }
    if (intern->ptr.max_value_string) { efree(intern->ptr.max_value_string); intern->ptr.max_value_string = NULL; }
    if (intern->ptr.multi_range)      { efree(intern->ptr.multi_range);      intern->ptr.multi_range      = NULL; }

    zend_object_std_dtor(&intern->zo);
}

/* }}} */

/* Helper: store/replace an internalized string buffer and update the cf
 * field to point at it (libxlsxwriter expects const char *, lifetime
 * until conditional_format_cell/range is called). */
static void cf_set_str(char **slot, const char **field, const char *src)
{
    if (*slot) efree(*slot);
    *slot = src ? estrdup(src) : NULL;
    *field = *slot;
}

/* {{{ ARG_INFO */

ZEND_BEGIN_ARG_INFO_EX(cf_construct_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(cf_long_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(cf_double_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(cf_string_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(cf_format_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, format)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(cf_bool_arginfo, 0, 0, 0)
    ZEND_ARG_INFO(0, on)
ZEND_END_ARG_INFO()
/* }}} */

/* {{{ \Vtiful\Kernel\ConditionalFormat::__construct() */
PHP_METHOD(vtiful_cond_format, __construct)
{
    cond_format_object *obj;

    ZVAL_COPY(return_value, getThis());
    obj = Z_COND_FMT_P(getThis());
    if (!obj->ptr.cf) {
        obj->ptr.cf = ecalloc(1, sizeof(lxw_conditional_format));
    }
}
/* }}} */

/* Common helpers used by the setters. */
#define CF_THIS()  cond_format_object *_obj = Z_COND_FMT_P(getThis()); \
                    if (!_obj->ptr.cf) { RETURN_NULL(); } \
                    ZVAL_COPY(return_value, getThis())

#define CF_FIELD()  _obj->ptr.cf

/* type / criteria / rule selectors */
PHP_METHOD(vtiful_cond_format, type)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->type = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, criteria)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->criteria = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, value)
{
    double v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_DOUBLE(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->value = v;
}

PHP_METHOD(vtiful_cond_format, valueString)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    cf_set_str(&_obj->ptr.value_string,
               &CF_FIELD()->value_string, ZSTR_VAL(v));
}

PHP_METHOD(vtiful_cond_format, minimum)
{
    double v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_DOUBLE(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->min_value = v;
}

PHP_METHOD(vtiful_cond_format, minimumString)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    cf_set_str(&_obj->ptr.min_value_string,
               &CF_FIELD()->min_value_string, ZSTR_VAL(v));
}

PHP_METHOD(vtiful_cond_format, minimumRule)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->min_rule_type = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, minimumColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->min_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, middle)
{
    double v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_DOUBLE(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->mid_value = v;
}

PHP_METHOD(vtiful_cond_format, middleString)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    cf_set_str(&_obj->ptr.mid_value_string,
               &CF_FIELD()->mid_value_string, ZSTR_VAL(v));
}

PHP_METHOD(vtiful_cond_format, middleRule)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->mid_rule_type = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, middleColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->mid_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, maximum)
{
    double v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_DOUBLE(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->max_value = v;
}

PHP_METHOD(vtiful_cond_format, maximumString)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    cf_set_str(&_obj->ptr.max_value_string,
               &CF_FIELD()->max_value_string, ZSTR_VAL(v));
}

PHP_METHOD(vtiful_cond_format, maximumRule)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->max_rule_type = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, maximumColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->max_color = (lxw_color_t)v;
}

/* Format: takes either a Format object, a Format resource, or a name
 * (delegates to existing object_format helper). For simplicity we accept a
 * Format resource (toResource()) and stash the lxw_format pointer. */
PHP_METHOD(vtiful_cond_format, format)
{
    zval *handle;
    lxw_format *fmt = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
        Z_PARAM_ZVAL(handle)
    ZEND_PARSE_PARAMETERS_END();

    CF_THIS();

    if (Z_TYPE_P(handle) == IS_RESOURCE) {
        fmt = (lxw_format *)zend_fetch_resource(
            Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_xls_writer);
    } else if (Z_TYPE_P(handle) == IS_OBJECT &&
               instanceof_function(Z_OBJCE_P(handle), vtiful_format_ce)) {
        format_object *fo = Z_FORMAT_P(handle);
        fmt = fo->ptr.format;
    }

    CF_FIELD()->format = fmt;
}

/* Data bar / icon set parameters */
PHP_METHOD(vtiful_cond_format, barColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, barOnly)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_only = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, dataBar2010)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->data_bar_2010 = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, barSolid)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_solid = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, barNegativeColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_negative_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, barBorderColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_border_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, barNegativeBorderColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_negative_border_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, barNoBorder)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_no_border = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, barDirection)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_direction = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, barAxisPosition)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_axis_position = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, barAxisColor)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->bar_axis_color = (lxw_color_t)v;
}

PHP_METHOD(vtiful_cond_format, iconStyle)
{
    zend_long v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_LONG(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->icon_style = (uint8_t)v;
}

PHP_METHOD(vtiful_cond_format, reverseIcons)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->reverse_icons = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, iconsOnly)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->icons_only = on ? 1 : 0;
}

PHP_METHOD(vtiful_cond_format, multiRange)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    cf_set_str(&_obj->ptr.multi_range, &CF_FIELD()->multi_range, ZSTR_VAL(v));
}

PHP_METHOD(vtiful_cond_format, stopIfTrue)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    CF_THIS();
    CF_FIELD()->stop_if_true = on ? 1 : 0;
}

/* {{{ method table */

zend_function_entry cond_format_methods[] = {
    PHP_ME(vtiful_cond_format, __construct,            cf_construct_arginfo, ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, type,                   cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, criteria,               cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, value,                  cf_double_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, valueString,            cf_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, minimum,                cf_double_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, minimumString,          cf_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, minimumRule,            cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, minimumColor,           cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, middle,                 cf_double_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, middleString,           cf_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, middleRule,             cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, middleColor,            cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, maximum,                cf_double_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, maximumString,          cf_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, maximumRule,            cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, maximumColor,           cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, format,                 cf_format_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barColor,               cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barOnly,                cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, dataBar2010,            cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barSolid,               cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barNegativeColor,       cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barBorderColor,         cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barNegativeBorderColor, cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barNoBorder,            cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barDirection,           cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barAxisPosition,        cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, barAxisColor,           cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, iconStyle,              cf_long_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, reverseIcons,           cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, iconsOnly,              cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, multiRange,             cf_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_cond_format, stopIfTrue,             cf_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_FE_END
};
/* }}} */

VTIFUL_STARTUP_FUNCTION(conditional_format)
{
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "ConditionalFormat", cond_format_methods);
    ce.create_object       = cond_format_objects_new;
    vtiful_cond_format_ce  = zend_register_internal_class(&ce);

    memcpy(&cond_format_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    cond_format_handlers.offset   = XtOffsetOf(cond_format_object, zo);
    cond_format_handlers.free_obj = cond_format_objects_free;

    /* Type constants */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_CELL",          LXW_CONDITIONAL_TYPE_CELL)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_TEXT",          LXW_CONDITIONAL_TYPE_TEXT)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_TIME_PERIOD",   LXW_CONDITIONAL_TYPE_TIME_PERIOD)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_AVERAGE",       LXW_CONDITIONAL_TYPE_AVERAGE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_DUPLICATE",     LXW_CONDITIONAL_TYPE_DUPLICATE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_UNIQUE",        LXW_CONDITIONAL_TYPE_UNIQUE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_TOP",           LXW_CONDITIONAL_TYPE_TOP)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_BOTTOM",        LXW_CONDITIONAL_TYPE_BOTTOM)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_BLANKS",        LXW_CONDITIONAL_TYPE_BLANKS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_NO_BLANKS",     LXW_CONDITIONAL_TYPE_NO_BLANKS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_ERRORS",        LXW_CONDITIONAL_TYPE_ERRORS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_NO_ERRORS",     LXW_CONDITIONAL_TYPE_NO_ERRORS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_FORMULA",       LXW_CONDITIONAL_TYPE_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_2_COLOR_SCALE", LXW_CONDITIONAL_2_COLOR_SCALE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_3_COLOR_SCALE", LXW_CONDITIONAL_3_COLOR_SCALE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_DATA_BAR",      LXW_CONDITIONAL_DATA_BAR)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "TYPE_ICON_SETS",     LXW_CONDITIONAL_TYPE_ICON_SETS)

    /* Criteria constants */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_EQUAL_TO",                 LXW_CONDITIONAL_CRITERIA_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_NOT_EQUAL_TO",             LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_GREATER_THAN",             LXW_CONDITIONAL_CRITERIA_GREATER_THAN)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_LESS_THAN",                LXW_CONDITIONAL_CRITERIA_LESS_THAN)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_GREATER_THAN_OR_EQUAL_TO", LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_LESS_THAN_OR_EQUAL_TO",    LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_BETWEEN",                  LXW_CONDITIONAL_CRITERIA_BETWEEN)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_NOT_BETWEEN",              LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TEXT_CONTAINING",          LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TEXT_NOT_CONTAINING",      LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TEXT_BEGINS_WITH",         LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TEXT_ENDS_WITH",           LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_YESTERDAY",    LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_TODAY",        LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_TOMORROW",     LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_LAST_7_DAYS",  LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_LAST_WEEK",    LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_THIS_WEEK",    LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_NEXT_WEEK",    LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_LAST_MONTH",   LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_THIS_MONTH",   LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TIME_PERIOD_NEXT_MONTH",   LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_AVERAGE_ABOVE",            LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_AVERAGE_BELOW",            LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "CRITERIA_TOP_OR_BOTTOM_PERCENT",    LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT)

    /* Rule types for color scales / data bars */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_MINIMUM",     LXW_CONDITIONAL_RULE_TYPE_MINIMUM)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_NUMBER",      LXW_CONDITIONAL_RULE_TYPE_NUMBER)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_PERCENT",     LXW_CONDITIONAL_RULE_TYPE_PERCENT)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_PERCENTILE",  LXW_CONDITIONAL_RULE_TYPE_PERCENTILE)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_FORMULA",     LXW_CONDITIONAL_RULE_TYPE_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "RULE_MAXIMUM",     LXW_CONDITIONAL_RULE_TYPE_MAXIMUM)

    /* Bar direction */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_DIRECTION_CONTEXT",         LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_DIRECTION_LEFT_TO_RIGHT",   LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_DIRECTION_RIGHT_TO_LEFT",   LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT)

    /* Bar axis position */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_AXIS_AUTOMATIC", LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_AXIS_MIDPOINT",  LXW_CONDITIONAL_BAR_AXIS_MIDPOINT)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "BAR_AXIS_NONE",      LXW_CONDITIONAL_BAR_AXIS_NONE)

    /* Icon styles */
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_ARROWS_COLORED", LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_ARROWS_GRAY",    LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_FLAGS",          LXW_CONDITIONAL_ICONS_3_FLAGS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_TRAFFIC_LIGHTS_UNRIMMED", LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_TRAFFIC_LIGHTS_RIMMED",   LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_SIGNS",                   LXW_CONDITIONAL_ICONS_3_SIGNS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_SYMBOLS_CIRCLED",         LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_3_SYMBOLS_UNCIRCLED",       LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_4_ARROWS_COLORED",          LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_4_ARROWS_GRAY",             LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_4_RED_TO_BLACK",            LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_4_RATINGS",                 LXW_CONDITIONAL_ICONS_4_RATINGS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_4_TRAFFIC_LIGHTS",          LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_5_ARROWS_COLORED",          LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_5_ARROWS_GRAY",             LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_5_RATINGS",                 LXW_CONDITIONAL_ICONS_5_RATINGS)
    REGISTER_CLASS_CONST_LONG(vtiful_cond_format_ce, "ICONS_5_QUARTERS",                LXW_CONDITIONAL_ICONS_5_QUARTERS)

    return SUCCESS;
}
