/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension                                                  |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2018 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <wjx@php.net>                                          |
  +----------------------------------------------------------------------+
*/

#ifndef PHP_READ_INCLUDE_H
#define PHP_READ_INCLUDE_H

#include "common.h"

#define READ_SKIP_ROW    0
#define READ_ROW         0x01
#define SKIP_EMPTY_VALUE 0x100

/* Forward declaration; full definition lives in xlswriter.h. */
struct xls_resource_read_t;

int  is_number(const char *value);
void data_to_null(zval *zv_result_t);
int  sheet_read_row(lxr_worksheet *ws);

void           sheet_list(lxr_workbook *wb, zval *zv_result_t);
void           sheet_list_with_meta(lxr_workbook *wb, zval *zv_result_t);
lxr_workbook  *file_open (const char *directory, const char *file_name);
lxr_worksheet *sheet_open(lxr_workbook *wb, const zend_string *zs_sheet_name_t, const zend_long zl_flag);

void skip_rows          (struct xls_resource_read_t *r, zval *zv_type_t, zend_long data_type_default, zend_long zl_skip_row);
void load_sheet_all_data(struct xls_resource_read_t *r, zend_long sheet_flag, zval *zv_type_t, zend_long data_type_default, zval *zv_result_t);
void load_sheet_row_data(struct xls_resource_read_t *r, zend_long sheet_flag, zval *zv_type_t, zend_long data_type_default, zval *zv_result_t);

unsigned int load_sheet_current_row_data         (struct xls_resource_read_t *r, zval *zv_result_t, zval *zv_type, zend_long data_type_default, unsigned int flag);
unsigned int load_sheet_current_row_data_callback(zend_string *zs_sheet_name_t, lxr_workbook *wb, void *callback_data);

void data_to_custom_type(const char *string_value, const size_t string_value_length, const zend_ulong type, zval *zv_result_t, const zend_ulong zv_hashtable_index);

#endif //PHP_READ_INCLUDE_H
