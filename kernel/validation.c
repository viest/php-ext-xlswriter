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

zend_class_entry *vtiful_validation_ce;

/* {{{ validation_objects_new
 */
static zend_object_handlers validation_handlers;

static zend_always_inline void *vtiful_validation_object_alloc(size_t obj_size, zend_class_entry *ce) {
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *validation_objects_new(zend_class_entry *ce)
{
    validation_object *validation = vtiful_validation_object_alloc(sizeof(validation_object), ce);

    zend_object_std_init(&validation->zo, ce);
    object_properties_init(&validation->zo, ce);

    validation->ptr.validation = NULL;

    validation->zo.handlers = &validation_handlers;

    return &validation->zo;
}
/* }}} */

/* {{{ validation_objects_free
 */
static void validation_objects_free(zend_object *object)
{
    validation_object *intern = php_vtiful_validation_fetch_object(object);

    if (intern->ptr.validation->value_list != NULL) {
        int index = 0;

        do {
            if (intern->ptr.validation->value_list[index] == NULL) {
                break;
            }

            efree(intern->ptr.validation->value_list[index]);
            index++;
        } while (1);

        efree(intern->ptr.validation->value_list);
    }

    if (intern->ptr.validation != NULL) {
        efree(intern->ptr.validation);
    }

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO
 */
ZEND_BEGIN_ARG_INFO_EX(validation_construct_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_type_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_criteria_type_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_ignore_blank_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, ignore_blank)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_show_input_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, show_input)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_show_error_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, show_error)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_error_type_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, error_type)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_dropdown_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, dropdown)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_value_number_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value_number)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_value_formula_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value_formula)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_value_list_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, value_list)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_value_date_time_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, timestamp)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_minimum_number_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, minimum_number)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_minimum_formula_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, minimum_formula)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_minimum_datetime_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, timestamp)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_maximum_number_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, maximum_number)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_maximum_formula_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, maximum_formula)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_maximum_datetime_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, maximum_datetime)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_input_title_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, input_title)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_input_message_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, input_message)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_error_title_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, error_titile)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_error_message_arginfo, 0, 0, 1)
                ZEND_ARG_INFO(0, error_message)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(validation_to_resource_arginfo, 0, 0, 1)
ZEND_END_ARG_INFO()
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::__construct()
 */
PHP_METHOD(vtiful_validation, __construct)
{
    validation_object *obj = NULL;

    ZVAL_COPY(return_value, getThis());

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        obj->ptr.validation = ecalloc(1, sizeof(lxw_data_validation));
    }

    obj->ptr.validation->value_list = NULL;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::validationType()
 */
PHP_METHOD(vtiful_validation, validationType)
{
    zend_long zl_validate_type = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_validate_type)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    if (zl_validate_type < LXW_VALIDATION_TYPE_NONE || zl_validate_type > LXW_VALIDATION_TYPE_ANY) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->validate = zl_validate_type;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::criteriaType()
 */
PHP_METHOD(vtiful_validation, criteriaType)
{
    zend_long zl_criteria_type = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_criteria_type)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    if (zl_criteria_type < LXW_VALIDATION_CRITERIA_NONE || zl_criteria_type > LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->criteria = zl_criteria_type;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::ignoreBlank()
 */
PHP_METHOD(vtiful_validation, ignoreBlank)
{
    zend_bool zb_ignore_blank = 1;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_BOOL(zb_ignore_blank)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    if (zb_ignore_blank) {
        obj->ptr.validation->ignore_blank = LXW_VALIDATION_ON;

        return;
    }

    obj->ptr.validation->ignore_blank = LXW_VALIDATION_OFF;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::showInput()
 */
PHP_METHOD(vtiful_validation, showInput)
{
    zend_bool zb_show_input = 1;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_BOOL(zb_show_input)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    if (zb_show_input) {
        obj->ptr.validation->show_input = LXW_VALIDATION_ON;

        return;
    }

    obj->ptr.validation->show_input = LXW_VALIDATION_OFF;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::showError()
 */
PHP_METHOD(vtiful_validation, showError)
{
    zend_bool zb_show_error = 1;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_BOOL(zb_show_error)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    if (zb_show_error) {
        obj->ptr.validation->show_error = LXW_VALIDATION_ON;

        return;
    }

    obj->ptr.validation->show_error = LXW_VALIDATION_OFF;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::errorType()
 */
PHP_METHOD(vtiful_validation, errorType)
{
    zend_long zl_error_type = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_error_type)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    if (zl_error_type < LXW_VALIDATION_ERROR_TYPE_STOP || zl_error_type > LXW_VALIDATION_ERROR_TYPE_INFORMATION) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->error_type = zl_error_type;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::dropdown()
 */
PHP_METHOD(vtiful_validation, dropdown)
{
    zend_bool zb_dropdown = 1;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(0, 1)
            Z_PARAM_OPTIONAL
            Z_PARAM_BOOL(zb_dropdown)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    if (zb_dropdown) {
        obj->ptr.validation->dropdown = LXW_VALIDATION_ON;

        return;
    }

    obj->ptr.validation->dropdown = LXW_VALIDATION_OFF;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::valueNumber()
 */
PHP_METHOD(vtiful_validation, valueNumber)
{
    zend_long zl_value_number = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(zl_value_number)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->value_number = zl_value_number;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::valueFormula()
 */
PHP_METHOD(vtiful_validation, valueFormula)
{
    zend_string *zs_value_formula = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_value_formula)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->value_formula = ZSTR_VAL(zs_value_formula);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::valueList()
 */
PHP_METHOD(vtiful_validation, valueList)
{
    int index = 0;
    char **list = NULL;

    Bucket *bucket;
    zval *zv_value_list = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_ARRAY(zv_value_list)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    if (obj->ptr.validation->value_list != NULL) {
        do {
            if (obj->ptr.validation->value_list[index] == NULL) {
                break;
            }

            efree(obj->ptr.validation->value_list[index]);
            index++;
        } while (1);

        efree(obj->ptr.validation->value_list);
        obj->ptr.validation->value_list = NULL;
    }

    ZVAL_COPY(return_value, getThis());

    zend_array *za_value_list = Z_ARR_P(zv_value_list);

    ZEND_HASH_FOREACH_BUCKET(za_value_list, bucket)
            if (Z_TYPE(bucket->val) != IS_STRING) {
                zend_throw_exception(vtiful_exception_ce, "Arrays can only consist of strings.", 300);
                return;
            }
            if (ZSTR_LEN(bucket->val.value.str) == 0 ) {
                zend_throw_exception(vtiful_exception_ce, "Array value is empty string.", 301);
                return;
            }
    ZEND_HASH_FOREACH_END();

    index = 0;
    list = ecalloc(za_value_list->nNumOfElements + 1, sizeof(char *));

    ZEND_HASH_FOREACH_BUCKET(za_value_list, bucket)
            list[index] = ecalloc(1, bucket->val.value.str->len + 1);
            strcpy(list[index],bucket->val.value.str->val);
            index++;
    ZEND_HASH_FOREACH_END();

    list[index] = NULL;

    obj->ptr.validation->value_list = list;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::valueDatetime(int $timestamp)
 */
PHP_METHOD(vtiful_validation, valueDatetime)
{
    zend_long timestamp = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(timestamp)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    obj->ptr.validation->value_datetime = timestamp_to_datetime(timestamp);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::minimumNumber(double $minimumNumber)
 */
PHP_METHOD(vtiful_validation, minimumNumber)
{
    double minimum_number = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_DOUBLE(minimum_number)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->minimum_number = minimum_number;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::minimumFormula()
 */
PHP_METHOD(vtiful_validation, minimumFormula)
{
    zend_string *zs_minimum_formula = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_minimum_formula)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->minimum_formula = ZSTR_VAL(zs_minimum_formula);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::minimumDatetime(int $timestamp)
 */
PHP_METHOD(vtiful_validation, minimumDatetime)
{
    zend_long timestamp = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(timestamp)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    obj->ptr.validation->minimum_datetime = timestamp_to_datetime(timestamp);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::maximumNumber(double $maximumNumber)
 */
PHP_METHOD(vtiful_validation, maximumNumber)
{
    double maximum_number = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_DOUBLE(maximum_number)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->maximum_number = maximum_number;
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::maximumFormula()
 */
PHP_METHOD(vtiful_validation, maximumFormula)
{
    zend_string *zs_maximum_formula = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_maximum_formula)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->maximum_formula = ZSTR_VAL(zs_maximum_formula);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::maximumDatetime(int $timestamp)
 */
PHP_METHOD(vtiful_validation, maximumDatetime)
{
    zend_long timestamp = 0;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_LONG(timestamp)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    obj->ptr.validation->maximum_datetime = timestamp_to_datetime(timestamp);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::inputTitle()
 */
PHP_METHOD(vtiful_validation, inputTitle)
{
    zend_string *zs_input_title = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_input_title)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->input_title = ZSTR_VAL(zs_input_title);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::inputMessage()
 */
PHP_METHOD(vtiful_validation, inputMessage)
{
    zend_string *zs_input_message = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_input_message)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->input_message = ZSTR_VAL(zs_input_message);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::errorTitle()
 */
PHP_METHOD(vtiful_validation, errorTitle)
{
    zend_string *zs_error_title = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_error_title)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->error_title = ZSTR_VAL(zs_error_title);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::errorMessage()
 */
PHP_METHOD(vtiful_validation, errorMessage)
{
    zend_string *zs_error_message = NULL;
    validation_object *obj = NULL;

    ZEND_PARSE_PARAMETERS_START(1, 1)
            Z_PARAM_STR(zs_error_message)
    ZEND_PARSE_PARAMETERS_END();

    obj = Z_VALIDATION_P(getThis());

    if (obj->ptr.validation == NULL) {
        RETURN_NULL();
    }

    ZVAL_COPY(return_value, getThis());

    obj->ptr.validation->error_message = ZSTR_VAL(zs_error_message);
}
/* }}} */

/** {{{ \Vtiful\Kernel\Validation::toResource()
 */
PHP_METHOD(vtiful_validation, toResource)
{
    validation_object *obj = Z_VALIDATION_P(getThis());

    RETURN_RES(zend_register_resource(obj->ptr.validation, le_xls_writer));
}
/* }}} */

/** {{{ validation_methods
*/
zend_function_entry validation_methods[] = {
        PHP_ME(vtiful_validation, __construct,     validation_construct_arginfo,        ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, validationType,  validation_type_arginfo,             ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, criteriaType,    validation_criteria_type_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, ignoreBlank,     validation_ignore_blank_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, showInput,       validation_show_input_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, showError ,      validation_show_error_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, errorType ,      validation_error_type_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, dropdown ,       validation_dropdown_arginfo,         ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, valueNumber,     validation_value_number_arginfo,     ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, valueFormula,    validation_value_formula_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, valueList,       validation_value_list_arginfo,       ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, valueDatetime,   validation_value_date_time_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, minimumNumber,   validation_minimum_number_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, minimumFormula,  validation_minimum_formula_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, minimumDatetime, validation_minimum_datetime_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, maximumNumber,   validation_maximum_number_arginfo,   ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, maximumFormula,  validation_maximum_formula_arginfo,  ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, maximumDatetime, validation_maximum_datetime_arginfo, ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, inputTitle,      validation_input_title_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, inputMessage,    validation_input_message_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, errorTitle,      validation_error_title_arginfo,      ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, errorMessage,    validation_error_message_arginfo,    ZEND_ACC_PUBLIC)
        PHP_ME(vtiful_validation, toResource,      validation_to_resource_arginfo,      ZEND_ACC_PUBLIC)
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(validation) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Validation", validation_methods);
    ce.create_object     = validation_objects_new;
    vtiful_validation_ce = zend_register_internal_class(&ce);

    memcpy(&validation_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    validation_handlers.offset   = XtOffsetOf(validation_object, zo);
    validation_handlers.free_obj = validation_objects_free;

    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_INTEGER",         LXW_VALIDATION_TYPE_INTEGER)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_INTEGER_FORMULA", LXW_VALIDATION_TYPE_INTEGER_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_DECIMAL",         LXW_VALIDATION_TYPE_DECIMAL)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_DECIMAL_FORMULA", LXW_VALIDATION_TYPE_DECIMAL_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_LIST",            LXW_VALIDATION_TYPE_LIST)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_LIST_FORMULA",    LXW_VALIDATION_TYPE_LIST_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_DATE",            LXW_VALIDATION_TYPE_DATE)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_DATE_FORMULA",    LXW_VALIDATION_TYPE_DATE_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_DATE_NUMBER",     LXW_VALIDATION_TYPE_DATE_NUMBER)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_TIME",            LXW_VALIDATION_TYPE_TIME)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_TIME_FORMULA",    LXW_VALIDATION_TYPE_TIME_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_TIME_NUMBER",     LXW_VALIDATION_TYPE_TIME_NUMBER)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_LENGTH",          LXW_VALIDATION_TYPE_LENGTH)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_LENGTH_FORMULA",  LXW_VALIDATION_TYPE_LENGTH_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_CUSTOM_FORMULA",  LXW_VALIDATION_TYPE_CUSTOM_FORMULA)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "TYPE_ANY",             LXW_VALIDATION_TYPE_ANY)

    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_BETWEEN",                  LXW_VALIDATION_CRITERIA_BETWEEN)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_NOT_BETWEEN",              LXW_VALIDATION_CRITERIA_NOT_BETWEEN)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_EQUAL_TO",                 LXW_VALIDATION_CRITERIA_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_NOT_EQUAL_TO",             LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_GREATER_THAN",             LXW_VALIDATION_CRITERIA_GREATER_THAN)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_LESS_THAN",                LXW_VALIDATION_CRITERIA_LESS_THAN)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_GREATER_THAN_OR_EQUAL_TO", LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "CRITERIA_LESS_THAN_OR_EQUAL_TO",    LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO)

    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "ERROR_TYPE_STOP",        LXW_VALIDATION_ERROR_TYPE_STOP)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "ERROR_TYPE_WARNING",     LXW_VALIDATION_ERROR_TYPE_WARNING)
    REGISTER_CLASS_CONST_LONG(vtiful_validation_ce, "ERROR_TYPE_INFORMATION", LXW_VALIDATION_ERROR_TYPE_INFORMATION)

    return SUCCESS;
}
/* }}} */