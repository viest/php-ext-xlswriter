/*
  +----------------------------------------------------------------------+
  | Vtiful Extension                                                     |
  +----------------------------------------------------------------------+
  | Copyright (c) 2017-2017 The Viest                                    |
  +----------------------------------------------------------------------+
  | http://www.viest.me                                                  |
  +----------------------------------------------------------------------+
  | Author: viest <dev@service.viest.me>                                 |
  +----------------------------------------------------------------------+
*/

#ifndef VTIFUL_EXCEL_H
#define VTIFUL_EXCEL_H

#define V_EXCEL_HANDLE "handle"
#define V_EXCEL_FIL "fileName"
#define V_EXCEL_COF "config"
#define V_EXCEL_PAT "path"

#define GET_CONFIG_PATH(dir_path_res, class_name, object)                                             \
    do {                                                                                              \
        zval *_config  = zend_read_property(class_name, object, ZEND_STRL(V_EXCEL_COF), 0, NULL);     \
        (dir_path_res) = zend_hash_str_find(Z_ARRVAL_P(_config), ZEND_STRL(V_EXCEL_PAT));             \
    } while(0)

extern zend_class_entry *vtiful_excel_ce;

VTIFUL_STARTUP_FUNCTION(excel);

#endif
