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

unsigned int xlsx_to_csv(
        zval *stream_resource,
        const char *delimiter_str, int delimiter_str_len,
        const char *enclosure_str, int enclosure_str_len,
        const char *escape_str, int escape_str_len,
        xlsxioreadersheet sheet_t,
        zval *zv_type_arr_t, zend_long data_type_default,
        unsigned int flag, zend_fcall_info *fci, zend_fcall_info_cache *fci_cache
);

#endif // PHP_EXT_XLS_WRITER_CSV_H
