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

#include "include.h"
#include "chart.h"

zend_class_entry *vtiful_chart_ce;

/* {{{ format_objects_new
 */
static zend_object_handlers chart_handlers;

static zend_always_inline void *vtiful_char_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *chart_objects_new(zend_class_entry *ce)
{
    chart_object *format = (chart_object *)vtiful_char_object_alloc(sizeof(chart_object), ce);

    zend_object_std_init(&format->zo, ce);
    object_properties_init(&format->zo, ce);

    format->ptr.chart   = NULL;
    format->ptr.series  = NULL;
    format->zo.handlers = &chart_handlers;

    return &format->zo;
}
/* }}} */

/* {{{ chart_objects_free
 */
static void chart_objects_free(zend_object *object)
{
    chart_object *intern = php_vtiful_chart_fetch_object(object);

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
ZEND_BEGIN_ARG_INFO_EX(chart_construct_arginfo, 0, 0, 2)
                ZEND_ARG_INFO(0, handle)
                ZEND_ARG_INFO(0, type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(chart_series_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::__construct(resource $handle, int $type)
 */
PHP_METHOD(vtiful_chart, __construct)
{
    zval *handle;
    zend_long type;
    chart_object *obj;
    xls_resource_t *xls_res;

    ZEND_PARSE_PARAMETERS_START(2, 2)
            Z_PARAM_RESOURCE(handle)
            Z_PARAM_LONG(type)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    xls_res = zval_get_resource(handle);
    obj     = Z_CHART_P(getThis());

    if (obj->ptr.chart == NULL) {
        obj->ptr.chart = workbook_add_chart(xls_res->workbook, LXW_CHART_LINE);
    }
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::series()
 */
PHP_METHOD(vtiful_chart, series)
{
    chart_object *obj;
    zend_string *value;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(value)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_CHART_P(getThis());

    chart_add_series(obj->ptr.chart, NULL, ZSTR_VAL(value));
}
/* }}} */

/** {{{ \Vtiful\Kernel\Chart::toResource()
 */
PHP_METHOD(vtiful_chart, toResource)
{
    chart_object *obj = Z_CHART_P(getThis());

    RETURN_RES(zend_register_resource(&obj->ptr, le_xls_writer));
}
/* }}} */

/** {{{ chart_methods
*/
zend_function_entry chart_methods[] = {
        PHP_ME(vtiful_chart, __construct, chart_construct_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, series,      chart_series_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_chart, toResource,  NULL,                    ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

VTIFUL_STARTUP_FUNCTION(chart)
{
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Chart", chart_methods);
    ce.create_object = chart_objects_new;
    vtiful_chart_ce  = zend_register_internal_class(&ce);

    memcpy(&chart_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    chart_handlers.offset   = XtOffsetOf(chart_object, zo);
    chart_handlers.free_obj = chart_objects_free;

    REGISTER_CLASS_CONST_LONG(vtiful_chart_ce, "CHART_LINE", LXW_CHART_LINE)

    return SUCCESS;
}
