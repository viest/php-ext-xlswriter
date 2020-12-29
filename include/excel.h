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

#ifndef VTIFUL_XLS_H
#define VTIFUL_XLS_H

#define V_XLS_HANDLE "handle"
#define V_XLS_FIL    "fileName"
#define V_XLS_COF    "config"
#define V_XLS_PAT    "path"
#define V_XLS_TYPE   "read_row_type"

#define V_XLS_CONST_READ_TYPE_INT      "TYPE_INT"
#define V_XLS_CONST_READ_TYPE_DOUBLE   "TYPE_DOUBLE"
#define V_XLS_CONST_READ_TYPE_STRING   "TYPE_STRING"
#define V_XLS_CONST_READ_TYPE_DATETIME "TYPE_TIMESTAMP"

#define V_XLS_CONST_READ_SKIP_NONE        "SKIP_NONE"
#define V_XLS_CONST_READ_SKIP_EMPTY_ROW   "SKIP_EMPTY_ROW"
#define V_XLS_CONST_READ_SKIP_HIDDEN_ROW  "SKIP_HIDDEN_ROW"
#define V_XLS_CONST_READ_SKIP_EMPTY_CELLS "SKIP_EMPTY_CELLS"
#define V_XLS_CONST_READ_SKIP_EMPTY_VALUE "SKIP_EMPTY_VALUE"

#define READ_TYPE_EMPTY    0x00
#define READ_TYPE_STRING   0x01
#define READ_TYPE_INT      0x02
#define READ_TYPE_DOUBLE   0x04
#define READ_TYPE_DATETIME 0x08

#define GET_CONFIG_PATH(dir_path_res, class_name, object)                                          \
    do {                                                                                           \
        zval rv;                                                                                   \
        zval *_config  = zend_read_property(class_name, object, ZEND_STRL(V_XLS_COF), 0, &rv);     \
        (dir_path_res) = zend_hash_str_find(Z_ARRVAL_P(_config), ZEND_STRL(V_XLS_PAT));            \
    } while(0)

extern zend_class_entry *vtiful_xls_ce;

VTIFUL_STARTUP_FUNCTION(excel);

#endif
