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

#ifndef PHP_READ_INCLUDE_H
#define PHP_READ_INCLUDE_H

#define READ_SKIP_ROW 0
#define READ_ROW 0x01

int sheet_read_row(xlsxioreadersheet sheet);
char* sheet_read_column(xlsxioreadersheet sheet);
xlsxioreader file_open(const char *directory, const char *file_name);
void load_sheet_all_data(xlsxioreadersheet sheet_t, zval *zv_result_t);
xlsxioreadersheet sheet_open(xlsxioreader file, const zend_string *zs_sheet_name_t);
unsigned int load_sheet_current_row_data(xlsxioreadersheet sheet_t, zval *zv_result_t, unsigned int flag);

#endif //PHP_READ_INCLUDE_H
