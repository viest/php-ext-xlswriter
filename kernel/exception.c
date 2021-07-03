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

zend_class_entry *vtiful_exception_ce;

/** {{{ exception_methods
*/
zend_function_entry exception_methods[] = {
        PHP_FE_END
};
/* }}} */

/** {{{ VTIFUL_STARTUP_FUNCTION
*/
VTIFUL_STARTUP_FUNCTION(exception) {
    zend_class_entry ce;

    INIT_NS_CLASS_ENTRY(ce, "Vtiful\\Kernel", "Exception", exception_methods);

    vtiful_exception_ce = zend_register_internal_class_ex(&ce, zend_ce_exception);

    return SUCCESS;
}
/* }}} */

/** {{{ exception_message_map
*/
char* exception_message_map(int code) {
    switch (code) {
        case LXW_ERROR_MEMORY_MALLOC_FAILED:
            return "Memory error, failed to malloc() required memory.";
        case LXW_ERROR_CREATING_XLSX_FILE:
            return "Error creating output xlsx file. Usually a permissions error.";
        case LXW_ERROR_CREATING_TMPFILE:
            return "Error encountered when creating a tmpfile during file assembly.";
        case LXW_ERROR_READING_TMPFILE:
            return "Error reading a tmpfile.";
        case LXW_ERROR_ZIP_FILE_OPERATION:
            return "Zlib error with a file operation while creating xlsx file.";
        case LXW_ERROR_ZIP_FILE_ADD:
            return "Zlib error when adding sub file to xlsx file.";
        case LXW_ERROR_ZIP_CLOSE:
            return "Zlib error when closing xlsx file.";
        case LXW_ERROR_NULL_PARAMETER_IGNORED:
            return "NULL function parameter ignored.";
        case LXW_ERROR_PARAMETER_VALIDATION:
            return "Function parameter validation error.";
        case LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED:
            return "Worksheet name exceeds Excel's limit of 31 characters.";
        case LXW_ERROR_INVALID_SHEETNAME_CHARACTER:
            return "Worksheet name contains invalid.";
        case LXW_ERROR_SHEETNAME_ALREADY_USED:
            return "Worksheet name is already in use.";
        case LXW_ERROR_32_STRING_LENGTH_EXCEEDED:
            return "Parameter exceeds Excel's limit of 32 characters.";
        case LXW_ERROR_128_STRING_LENGTH_EXCEEDED:
            return "Parameter exceeds Excel's limit of 128 characters.";
        case LXW_ERROR_255_STRING_LENGTH_EXCEEDED:
            return "Parameter exceeds Excel's limit of 255 characters.";
        case LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED:
            return "String exceeds Excel's limit of 32:767 characters.";
        case LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND:
            return "Error finding internal string index.";
        case LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE:
            return "Worksheet row or column index out of range.";
        case LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED:
            return "Maximum number of worksheet URLs (65530) exceeded.";
        case LXW_ERROR_IMAGE_DIMENSIONS:
            return "Couldn't read image dimensions or DPI.";
        default:
            return "Unknown error";
    }
}
/* }}} */