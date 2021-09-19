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

zend_class_entry *vtiful_rich_string_ce;

/* {{{ rich_string_objects_new
 */
static zend_object_handlers rich_string_handlers;

static zend_always_inline void *vtiful_rich_string_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *rich_string_objects_new(zend_class_entry *ce)
{
    rich_string_object *rich_string = vtiful_rich_string_object_alloc(sizeof(rich_string_object), ce);

    zend_object_std_init(&rich_string->zo, ce);
    object_properties_init(&rich_string->zo, ce);

    rich_string->ptr.tuple = NULL;
    rich_string->zo.handlers = &rich_string_handlers;

    return &rich_string->zo;
}
/* }}} */

/* {{{ rich_string_objects_free
 */
static void rich_string_objects_free(zend_object *object)
{
    rich_string_object *intern = php_vtiful_rich_string_fetch_object(object);

    if (intern->ptr.tuple != NULL) {
        efree(intern->ptr.tuple);
        intern->ptr.tuple = NULL;
    }

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(rich_string_construct_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, text)
                ZEND_ARG_INFO(0, format_handle)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\RichString::__construct(string $text, resource $format)
 */
PHP_METHOD(vtiful_rich_string, __construct)
{
    zend_string *text = NULL;
    zval *format_handle = NULL;
    rich_string_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 2)
            Z_PARAM_STR(text)
            Z_PARAM_OPTIONAL
            Z_PARAM_RESOURCE_OR_NULL(format_handle)
    ZEND_PARSE_PARAMETERS_END();

    ZVAL_COPY(return_value, getThis());

    obj = Z_RICH_STR_P(getThis());

    if (obj->ptr.tuple != NULL) {
        return;
    }

    lxw_rich_string_tuple *instance = (lxw_rich_string_tuple *)ecalloc(1, sizeof(lxw_rich_string_tuple));

    zend_string *zstr = zend_string_copy(text);

    if (format_handle == NULL) {
        instance->format = NULL;
        instance->string = ZSTR_VAL(zstr);
    } else {
        instance->format = zval_get_format(format_handle);
        instance->string = ZSTR_VAL(zstr);
    }

    obj->ptr.tuple = instance;
}
/* }}} */

/** {{{ rich_string_methods
*/
zend_function_entry rich_string_methods[] = {
        PHP_ME(vtiful_rich_string, __construct, rich_string_construct_arginfo,   ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(rich_string) {
        zend_class_entry ce;

        INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "RichString", rich_string_methods);
        ce.create_object = rich_string_objects_new;
        vtiful_rich_string_ce = zend_register_internal_class(&ce);

        memcpy(&rich_string_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
        rich_string_handlers.offset   = XtOffsetOf(rich_string_object, zo);
        rich_string_handlers.free_obj = rich_string_objects_free;

        return SUCCESS;
}
/* }}} */