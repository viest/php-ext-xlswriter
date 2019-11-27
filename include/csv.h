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

#ifndef PHP_EXT_XLS_WRITER_CSV_H
#define PHP_EXT_XLS_WRITER_CSV_H

unsigned int xlsx_to_csv(zval *stream_resource, xlsxioreadersheet sheet_t, zval *zv_type_arr_t, unsigned int flag, zend_fcall_info *fci, zend_fcall_info_cache *fci_cache);

#endif // PHP_EXT_XLS_WRITER_CSV_H
