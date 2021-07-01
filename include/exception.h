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

#ifndef VTIFUL_XLS_EXCEPTION_H
#define VTIFUL_XLS_EXCEPTION_H

extern zend_class_entry *vtiful_exception_ce;

VTIFUL_STARTUP_FUNCTION(exception);

char* exception_message_map(int code);

#endif
