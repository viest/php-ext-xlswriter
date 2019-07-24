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
xlsxioreader file_open(const char *directory, const char *file_name) {
    char path[strlen(directory) + strlen(file_name) + 2];
    xlsxioreader file;

    lxw_strcpy(path, directory);
    strcat(path, "/");
    strcat(path, file_name);

    if ((file = xlsxioread_open(path)) == NULL) {
        zend_throw_exception(vtiful_exception_ce, "Failed to open file", 100);
        return NULL;
    }

    return file;
}
/* }}} */

/* {{{ */
xlsxioreadersheet sheet_open(xlsxioreader file, const zend_string *zs_sheet_name_t)
{
    if (zs_sheet_name_t == NULL) {
        return xlsxioread_sheet_open(file, NULL, XLSXIOREAD_SKIP_EMPTY_ROWS);
    }

    return xlsxioread_sheet_open(file, ZSTR_VAL(zs_sheet_name_t), XLSXIOREAD_SKIP_EMPTY_ROWS);
}
/* }}} */

/* {{{ */
int sheet_read_row(xlsxioreadersheet sheet)
{
    return xlsxioread_sheet_next_row(sheet);
}
/* }}} */

/* {{{ */
char* sheet_read_column(xlsxioreadersheet sheet)
{
    return xlsxioread_sheet_next_cell(sheet);
}
/* }}} */

/* {{{ */
unsigned int load_sheet_current_row_data(xlsxioreadersheet sheet_t, zval *zv_result_t, unsigned int flag)
{
    char *_string_value = NULL;

    if (flag && !sheet_read_row(sheet_t)) {
        return XLSWRITER_FALSE;
    }

    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while ((_string_value = sheet_read_column(sheet_t)) != NULL)
    {
        add_next_index_stringl(zv_result_t, _string_value, strlen(_string_value));
    }

    return XLSWRITER_TRUE;
}
/* }}} */

/* {{{ */
void load_sheet_all_data(xlsxioreadersheet sheet_t, zval *zv_result_t)
{
    if (Z_TYPE_P(zv_result_t) != IS_ARRAY) {
        array_init(zv_result_t);
    }

    while (sheet_read_row(sheet_t))
    {
        zval _zv_tmp_row;
        load_sheet_current_row_data(sheet_t, &_zv_tmp_row, READ_SKIP_ROW);
        add_next_index_zval(zv_result_t, &_zv_tmp_row);
    }
}
/* }}} */