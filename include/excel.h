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

#define GET_CONFIG_PATH(dir_path_res, class_name, object)                                             \
    do {                                                                                              \
        zval *_config  = zend_read_property(class_name, object, ZEND_STRL(V_XLS_COF), 0, NULL);     \
        (dir_path_res) = zend_hash_str_find(Z_ARRVAL_P(_config), ZEND_STRL(V_XLS_PAT));             \
    } while(0)

extern zend_class_entry *vtiful_xls_ce;

VTIFUL_STARTUP_FUNCTION(excel);

#endif
