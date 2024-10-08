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

#ifndef PHP_EXT_XLS_EXPORT_HELP_H
#define PHP_EXT_XLS_EXPORT_HELP_H

#include "common.h"

unsigned int file_exists(const char *path);
unsigned int directory_exists(const char *path);
zend_long date_double_to_timestamp(double value);
lxw_row_col_options* default_row_col_options();

#endif //PHP_EXT_XLS_EXPORT_HELP_H
