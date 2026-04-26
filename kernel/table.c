/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension — Excel Table builder.                            |
  +----------------------------------------------------------------------+
  | Configures lxw_table_options + an array of lxw_table_column. Each PHP |
  | column entry is converted to a calloc'd lxw_table_column with its     |
  | string fields strdup'd into the table_object's owned string list.     |
  +----------------------------------------------------------------------+
*/

#include "xlswriter.h"

zend_class_entry *vtiful_table_ce;

/* {{{ object lifecycle */
static zend_object_handlers table_handlers;

static zend_always_inline void *vtiful_table_object_alloc(size_t obj_size, zend_class_entry *ce)
{
    void *obj = emalloc(obj_size);
    memset(obj, 0, obj_size);
    return obj;
}

PHP_VTIFUL_API zend_object *table_objects_new(zend_class_entry *ce)
{
    table_object *intern = vtiful_table_object_alloc(sizeof(table_object), ce);
    zend_object_std_init(&intern->zo, ce);
    object_properties_init(&intern->zo, ce);
    intern->ptr.opts    = NULL;
    intern->ptr.name    = NULL;
    intern->ptr.columns = NULL;
    intern->zo.handlers = &table_handlers;
    return &intern->zo;
}

static void free_columns(lxw_table_column **columns)
{
    if (!columns) return;
    for (int i = 0; columns[i]; i++) {
        if (columns[i]->header)       efree((void *)columns[i]->header);
        if (columns[i]->formula)      efree((void *)columns[i]->formula);
        if (columns[i]->total_string) efree((void *)columns[i]->total_string);
        efree(columns[i]);
    }
    efree(columns);
}

static void table_objects_free(zend_object *object)
{
    table_object *intern = php_vtiful_table_fetch_object(object);
    if (intern->ptr.opts) {
        efree(intern->ptr.opts);
        intern->ptr.opts = NULL;
    }
    if (intern->ptr.name) {
        efree(intern->ptr.name);
        intern->ptr.name = NULL;
    }
    free_columns(intern->ptr.columns);
    intern->ptr.columns = NULL;

    zend_object_std_dtor(&intern->zo);
}
/* }}} */

/* {{{ ARG_INFO */

ZEND_BEGIN_ARG_INFO_EX(table_construct_arginfo, 0, 0, 0)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(table_string_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(table_long_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, value)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(table_bool_arginfo, 0, 0, 0)
    ZEND_ARG_INFO(0, on)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(table_columns_arginfo, 0, 0, 1)
    ZEND_ARG_INFO(0, columns)
ZEND_END_ARG_INFO()

ZEND_BEGIN_ARG_INFO_EX(table_style_arginfo, 0, 0, 2)
    ZEND_ARG_INFO(0, type)
    ZEND_ARG_INFO(0, number)
ZEND_END_ARG_INFO()

/* }}} */

#define TABLE_THIS()  table_object *_obj = Z_TABLE_P(getThis()); \
                       if (!_obj->ptr.opts) { RETURN_NULL(); } \
                       ZVAL_COPY(return_value, getThis())

/* {{{ \Vtiful\Kernel\Table::__construct() */
PHP_METHOD(vtiful_table, __construct)
{
    table_object *obj;
    ZVAL_COPY(return_value, getThis());
    obj = Z_TABLE_P(getThis());
    if (!obj->ptr.opts) {
        obj->ptr.opts = ecalloc(1, sizeof(lxw_table_options));
    }
}
/* }}} */

/* Setters */

PHP_METHOD(vtiful_table, name)
{
    zend_string *v;
    ZEND_PARSE_PARAMETERS_START(1, 1) Z_PARAM_STR(v) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    if (_obj->ptr.name) efree(_obj->ptr.name);
    _obj->ptr.name = estrdup(ZSTR_VAL(v));
    _obj->ptr.opts->name = _obj->ptr.name;
}

PHP_METHOD(vtiful_table, noHeaderRow)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->no_header_row = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, noAutofilter)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->no_autofilter = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, noBandedRows)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->no_banded_rows = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, bandedColumns)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->banded_columns = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, firstColumn)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->first_column = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, lastColumn)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->last_column = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, totalRow)
{
    zend_bool on = 1;
    ZEND_PARSE_PARAMETERS_START(0, 1) Z_PARAM_OPTIONAL Z_PARAM_BOOL(on) ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->total_row = on ? 1 : 0;
}

PHP_METHOD(vtiful_table, style)
{
    zend_long type, number;
    ZEND_PARSE_PARAMETERS_START(2, 2)
        Z_PARAM_LONG(type)
        Z_PARAM_LONG(number)
    ZEND_PARSE_PARAMETERS_END();
    TABLE_THIS();
    _obj->ptr.opts->style_type        = (uint8_t)type;
    _obj->ptr.opts->style_type_number = (uint8_t)number;
}

/* {{{ Table::columns(array $columns)
 *  Each entry is an associative array with optional keys:
 *    header (string), formula (string), total_string (string),
 *    total_function (int), header_format (Format|resource),
 *    format (Format|resource), total_value (double).
 */
PHP_METHOD(vtiful_table, columns)
{
    zval *zv_cols = NULL, *entry;
    int count, idx;
    lxw_table_column **arr;

    ZEND_PARSE_PARAMETERS_START(1, 1)
        Z_PARAM_ARRAY(zv_cols)
    ZEND_PARSE_PARAMETERS_END();

    TABLE_THIS();

    /* Free any prior columns. */
    if (_obj->ptr.columns) {
        free_columns(_obj->ptr.columns);
        _obj->ptr.columns = NULL;
    }

    count = zend_hash_num_elements(Z_ARRVAL_P(zv_cols));
    arr = ecalloc(count + 1, sizeof(lxw_table_column *));
    idx = 0;
    ZEND_HASH_FOREACH_VAL(Z_ARRVAL_P(zv_cols), entry) {
        if (Z_TYPE_P(entry) != IS_ARRAY) continue;

        lxw_table_column *col = ecalloc(1, sizeof(lxw_table_column));
        zval *v;

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "header", sizeof("header") - 1);
        if (v && Z_TYPE_P(v) == IS_STRING) col->header = estrdup(Z_STRVAL_P(v));

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "formula", sizeof("formula") - 1);
        if (v && Z_TYPE_P(v) == IS_STRING) col->formula = estrdup(Z_STRVAL_P(v));

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "total_string", sizeof("total_string") - 1);
        if (v && Z_TYPE_P(v) == IS_STRING) col->total_string = estrdup(Z_STRVAL_P(v));

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "total_function", sizeof("total_function") - 1);
        if (v && Z_TYPE_P(v) == IS_LONG) col->total_function = (uint8_t)Z_LVAL_P(v);

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "total_value", sizeof("total_value") - 1);
        if (v) col->total_value = zval_get_double(v);

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "format", sizeof("format") - 1);
        if (v) {
            if (Z_TYPE_P(v) == IS_RESOURCE) {
                col->format = (lxw_format *)zend_fetch_resource(
                    Z_RES_P(v), VTIFUL_RESOURCE_NAME, le_xls_writer);
            } else if (Z_TYPE_P(v) == IS_OBJECT &&
                       instanceof_function(Z_OBJCE_P(v), vtiful_format_ce)) {
                format_object *fo = Z_FORMAT_P(v);
                col->format = fo->ptr.format;
            }
        }

        v = zend_hash_str_find(Z_ARRVAL_P(entry), "header_format", sizeof("header_format") - 1);
        if (v) {
            if (Z_TYPE_P(v) == IS_RESOURCE) {
                col->header_format = (lxw_format *)zend_fetch_resource(
                    Z_RES_P(v), VTIFUL_RESOURCE_NAME, le_xls_writer);
            } else if (Z_TYPE_P(v) == IS_OBJECT &&
                       instanceof_function(Z_OBJCE_P(v), vtiful_format_ce)) {
                format_object *fo = Z_FORMAT_P(v);
                col->header_format = fo->ptr.format;
            }
        }

        arr[idx++] = col;
    } ZEND_HASH_FOREACH_END();
    arr[idx] = NULL;
    _obj->ptr.columns         = arr;
    _obj->ptr.opts->columns   = arr;
}
/* }}} */

zend_function_entry table_methods[] = {
    PHP_ME(vtiful_table, __construct,    table_construct_arginfo, ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, name,           table_string_arginfo,    ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, noHeaderRow,    table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, noAutofilter,   table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, noBandedRows,   table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, bandedColumns,  table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, firstColumn,    table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, lastColumn,     table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, totalRow,       table_bool_arginfo,      ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, style,          table_style_arginfo,     ZEND_ACC_PUBLIC)
    PHP_ME(vtiful_table, columns,        table_columns_arginfo,   ZEND_ACC_PUBLIC)
    PHP_FE_END
};

VTIFUL_STARTUP_FUNCTION(table)
{
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Table", table_methods);
    ce.create_object = table_objects_new;
    vtiful_table_ce  = zend_register_internal_class(&ce);

    memcpy(&table_handlers, zend_get_std_object_handlers(), sizeof(zend_object_handlers));
    table_handlers.offset   = XtOffsetOf(table_object, zo);
    table_handlers.free_obj = table_objects_free;

    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "STYLE_TYPE_DEFAULT", LXW_TABLE_STYLE_TYPE_DEFAULT)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "STYLE_TYPE_LIGHT",   LXW_TABLE_STYLE_TYPE_LIGHT)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "STYLE_TYPE_MEDIUM",  LXW_TABLE_STYLE_TYPE_MEDIUM)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "STYLE_TYPE_DARK",    LXW_TABLE_STYLE_TYPE_DARK)

    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_NONE",      LXW_TABLE_FUNCTION_NONE)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_AVERAGE",   LXW_TABLE_FUNCTION_AVERAGE)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_COUNT_NUMS",LXW_TABLE_FUNCTION_COUNT_NUMS)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_COUNT",     LXW_TABLE_FUNCTION_COUNT)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_MAX",       LXW_TABLE_FUNCTION_MAX)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_MIN",       LXW_TABLE_FUNCTION_MIN)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_STD_DEV",   LXW_TABLE_FUNCTION_STD_DEV)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_SUM",       LXW_TABLE_FUNCTION_SUM)
    REGISTER_CLASS_CONST_LONG(vtiful_table_ce, "FUNCTION_VAR",       LXW_TABLE_FUNCTION_VAR)

    return SUCCESS;
}
