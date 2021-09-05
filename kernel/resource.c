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

/* {{{ */
xls_resource_write_t * zval_get_resource(zval *handle)
{
    xls_resource_write_t *res;

    if((res = (xls_resource_write_t *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "XLS resources resolution fail", 210);
    }

    return res;
}
/* }}} */

/* {{{ */
lxw_format * zval_get_format(zval *handle)
{
    lxw_format *res = NULL;

    if (handle == NULL) {
        return NULL;
    }

    if (zval_get_type(handle) != IS_RESOURCE) {
        return NULL;
    }

    if((res = (lxw_format *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "format resources resolution fail", 210);
    }

    return res;
}
/* }}} */

/* {{{ */
xls_resource_chart_t *zval_get_chart(zval *resource)
{
    xls_resource_chart_t *res;

    if((res = (xls_resource_chart_t *)zend_fetch_resource(Z_RES_P(resource), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "chart resources resolution fail", 210);
    }

    return res;
}
/* }}} */

/* {{{ */
lxw_rich_string_tuple *zval_get_rich_string(zval *resource)
{
    lxw_rich_string_tuple *res;

    if((res = (lxw_rich_string_tuple *)zend_fetch_resource(Z_RES_P(resource), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "rich string resources resolution fail", 210);
    }

    return res;
}
/* }}} */

/* {{{ */
lxw_data_validation *zval_get_validation(zval *resource)
{
    lxw_data_validation *res;

    if((res = (lxw_data_validation *)zend_fetch_resource(Z_RES_P(resource), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "validation resources resolution fail", 210);
    }

    return res;
}
/* }}} */