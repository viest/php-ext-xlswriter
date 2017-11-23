#include <php.h>
#include "zend_exceptions.h"
#include "php_vtiful.h"
#include "exception.h"

zend_class_entry *vtiful_exception_ce;

zend_function_entry exception_methods[] = {
        PHP_FE_END
};

VTIFUL_STARTUP_FUNCTION(vtiful_exception) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Exception", exception_methods);

    vtiful_exception_ce = zend_register_internal_class_ex(&ce, zend_ce_exception);

    return SUCCESS;
}
