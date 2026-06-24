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

zend_class_entry *vtiful_chart_ce;

/* {{{ lxlsx_format_objects_new
 */
static zend_object_handlers lxlsx_chart_handlers;

static zend_always_inline void *vtiful_char_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *lxlsx_chart_objects_new(zend_class_entry *ce)
{
    lxlsx_chart_object *format = (lxlsx_chart_object *)vtiful_char_object_alloc(sizeof(lxlsx_chart_object), ce);

    zend_object_std_init(&format->zo, ce);
    object_properties_init(&format->zo, ce);

    format->ptr.chart   = NULL;
    format->ptr.series  = NULL;
    format->zo.handlers = &lxlsx_chart_handlers;

    return &format->zo;
}
/* }}} */

/* {{{ lxlsx_chart_objects_free
 */
static void lxlsx_chart_objects_free(zend_object *object)
{
    lxlsx_chart_object *intern = php_vtiful_chart_fetch_object(object);

    if (intern->ptr.series != NULL) {
        // free by workbook
        intern->ptr.series = NULL;
    }

    if (intern->ptr.chart != NULL) {
        // free by workbook
        intern->ptr.chart = NULL;
    }

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_construct_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, handle)
                ZEND_ARG_INFO(0, type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_series_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value)
                ZEND_ARG_INFO(0, categories)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_series_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_style_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, style)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_axis_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, name)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_title_name_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, title)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_legend_set_position_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(lxlsx_chart_to_resource_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::__construct(resource $handle, int $type)
 */
PHP_METHOD(vtiful_chart, __construct)
{
    zval *handle;
    lxlsx_chart_object *obj;
    zend_long type = 0;
    xls_resource_write_t *xls_res;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_RESOURCE(handle)
            Z_PARAM_LONG(type)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_res = zval_get_resource(handle);

    if (xls_res == NULL) {
        return;
    }

    obj = Z_CHART_P(getThis());

    U8_RANGE_EXCEPTION(type);

    if (obj->ptr.chart == NULL) {
        obj->ptr.chart = lxlsx_workbook_add_chart(xls_res->workbook, (uint8_t)type);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::series(string $value, string $categories)
 */
PHP_METHOD(vtiful_chart, series)
{
    lxlsx_chart_object *obj;
    zend_string *values, *categories = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_STR(values)
            Z_PARAM_OPTIONAL
            Z_PARAM_STR(categories)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    if (ZEND_NUM_ARGS() == 2) {
        obj->ptr.series = lxlsx_chart_add_series(obj->ptr.chart, ZSTR_VAL(categories), ZSTR_VAL(values));
    }

    if (ZEND_NUM_ARGS() == 1) {
        obj->ptr.series = lxlsx_chart_add_series(obj->ptr.chart, NULL, ZSTR_VAL(values));
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::seriesName(string $value)
 */
PHP_METHOD(vtiful_chart, seriesName)
{
    lxlsx_chart_object *obj;
    zend_string *values;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(values)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    lxlsx_chart_series_set_name(obj->ptr.series, ZSTR_VAL(values));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::style(int $style)
 */
PHP_METHOD(vtiful_chart, style)
{
    lxlsx_chart_object *obj;
    zend_long style = 0;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(style)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    U8_RANGE_EXCEPTION(style);

    lxlsx_chart_set_style(obj->ptr.chart, (uint8_t)style);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::axisNameX(string $name)
 */
PHP_METHOD(vtiful_chart, axisNameX)
{
    lxlsx_chart_object *obj;
    zend_string *name;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    lxlsx_chart_axis_set_name(obj->ptr.chart->x_axis, ZSTR_VAL(name));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::axisNameY(string $name)
 */
PHP_METHOD(vtiful_chart, axisNameY)
{
    lxlsx_chart_object *obj;
    zend_string *name;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(name)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    lxlsx_chart_axis_set_name(obj->ptr.chart->y_axis, ZSTR_VAL(name));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::title(string $title)
 */
PHP_METHOD(vtiful_chart, title)
{
    lxlsx_chart_object *obj;
    zend_string *title;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(title)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    lxlsx_chart_title_set_name(obj->ptr.chart, ZSTR_VAL(title));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::legendSetPosition(int $type)
 */
PHP_METHOD(vtiful_chart, legendSetPosition)
{
    zend_long type = 0;
    lxlsx_chart_object *obj;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(type)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    U8_RANGE_EXCEPTION(type);

    lxlsx_chart_legend_set_position(obj->ptr.chart, (uint8_t)type);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::toResource()
 */
PHP_METHOD(vtiful_chart, toResource)
{
    lxlsx_chart_object *obj = Z_CHART_P(getThis());

    RETURN_RES(zend_register_resource(&obj->ptr, le_xls_writer));
}
/* }}} */

/** {{{ lxlsx_chart_methods
*/
zend_function_entry lxlsx_chart_methods[] = {
        PHP_ME(vtiful_chart, __construct,       lxlsx_chart_construct_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, series,            lxlsx_chart_series_arginfo,              ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, seriesName,        lxlsx_chart_series_name_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, style,             lxlsx_chart_style_arginfo,               ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, axisNameY,         lxlsx_chart_axis_name_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, axisNameX,         lxlsx_chart_axis_name_arginfo,           ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, title,             lxlsx_chart_title_name_arginfo,          ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, legendSetPosition, lxlsx_chart_legend_set_position_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, toResource,        lxlsx_chart_to_resource_arginfo,         ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

/* {{{ */
VTIFUL_STARTUP_FUNCTION(chart)
{
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Chart", lxlsx_chart_methods);
    ce.create_object = lxlsx_chart_objects_new;
    vtiful_chart_ce  = zend_register_internal_class(&ce);

    memcpy(&lxlsx_chart_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    lxlsx_chart_handlers.offset   = offsetof(lxlsx_chart_object, zo);
    lxlsx_chart_handlers.free_obj = lxlsx_chart_objects_free;

    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_BAR",                           LXLSX_CHART_BAR)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_BAR_STACKED",                   LXLSX_CHART_BAR_STACKED)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_BAR_STACKED_PERCENT",           LXLSX_CHART_BAR_STACKED_PERCENT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_AREA",                          LXLSX_CHART_AREA)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_AREA_STACKED",                  LXLSX_CHART_AREA_STACKED)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_AREA_STACKED_PERCENT",          LXLSX_CHART_AREA_STACKED_PERCENT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LINE",                          LXLSX_CHART_LINE)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_COLUMN",                        LXLSX_CHART_COLUMN)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_COLUMN_STACKED",                LXLSX_CHART_COLUMN_STACKED)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_COLUMN_STACKED_PERCENT",        LXLSX_CHART_COLUMN_STACKED_PERCENT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_DOUGHNUT",                      LXLSX_CHART_DOUGHNUT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_PIE",                           LXLSX_CHART_PIE)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_SCATTER",                       LXLSX_CHART_SCATTER)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_SCATTER_STRAIGHT",              LXLSX_CHART_SCATTER_STRAIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_SCATTER_STRAIGHT_WITH_MARKERS", LXLSX_CHART_SCATTER_STRAIGHT_WITH_MARKERS)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_SCATTER_SMOOTH",                LXLSX_CHART_SCATTER_SMOOTH)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_SCATTER_SMOOTH_WITH_MARKERS",   LXLSX_CHART_SCATTER_SMOOTH_WITH_MARKERS)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_RADAR",                         LXLSX_CHART_RADAR)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_RADAR_WITH_MARKERS",            LXLSX_CHART_RADAR_WITH_MARKERS)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_RADAR_FILLED",                  LXLSX_CHART_RADAR_FILLED)

    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_NONE",                   LXLSX_CHART_LEGEND_NONE)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_RIGHT",                  LXLSX_CHART_LEGEND_RIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_LEFT",                   LXLSX_CHART_LEGEND_LEFT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_TOP",                    LXLSX_CHART_LEGEND_TOP)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_BOTTOM",                 LXLSX_CHART_LEGEND_BOTTOM)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_OVERLAY_RIGHT",          LXLSX_CHART_LEGEND_OVERLAY_RIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_OVERLAY_LEFT",           LXLSX_CHART_LEGEND_OVERLAY_LEFT)

#if defined(LXLSX_VERSION_ID) && LXLSX_VERSION_ID >= 95
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LINE_STACKED",                  LXLSX_CHART_LINE_STACKED)
    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LINE_STACKED_PERCENT",          LXLSX_CHART_LINE_STACKED_PERCENT)
#endif

    // PECL Windows version is 0.7.7, but define in 0.7.8
    //REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_TOP_RIGHT",              LXLSX_CHART_LEGEND_TOP_RIGHT)
    //REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LEGEND_OVERLAY_TOP_RIGHT",      LXLSX_CHART_LEGEND_OVERLAY_TOP_RIGHT)

    return SUCCESS;
}
/* }}} */