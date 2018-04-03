#include "../include.h"

/* {{{ */
excel_resource_t * zval_get_resource(zval *handle)
{
    excel_resource_t *res;

    if((res = (excel_resource_t *)zend_fetch_resource(Z_RES_P(handle), VTIFUL_RESOURCE_NAME, le_excel_writer)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Excel resources resolution fail", 210);
    }

    return res;
}
/* }}} */