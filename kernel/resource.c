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
    lxw_format *res;

    if((res = (lxw_format *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "format resources resolution fail", 210);
    }

    return res;
}
/* }}} */

xls_resource_chart_t *zval_get_chart(zval *resource)
{
    xls_resource_chart_t *res;

    if((res = (xls_resource_chart_t *)zend_fetch_resource(Z_RES_P(resource), VTIFUL_RESOURCE_NAME, le_xls_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "chart resources resolution fail", 210);
    }

    return res;
}