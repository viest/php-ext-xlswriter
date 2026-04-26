#include "xlsxreader/common.h"

const char *lxr_strerror(lxr_error code)
{
    switch (code) {
    case LXR_NO_ERROR:                   return "no error";
    case LXR_ERROR_MEMORY_MALLOC_FAILED: return "memory allocation failed";
    case LXR_ERROR_FILE_OPEN_FAILED:     return "failed to open file";
    case LXR_ERROR_FILE_NOT_XLSX:        return "file is not a valid XLSX archive";
    case LXR_ERROR_FILE_CORRUPTED:       return "file is corrupted";
    case LXR_ERROR_ZIP_ENTRY_NOT_FOUND:  return "zip entry not found";
    case LXR_ERROR_XML_PARSE:            return "XML parse error";
    case LXR_ERROR_SHEET_NOT_FOUND:      return "sheet not found";
    case LXR_ERROR_NULL_PARAMETER:       return "null parameter";
    case LXR_ERROR_END_OF_DATA:          return "end of data";
    case LXR_ERROR_INVALID_CELL_REF:     return "invalid cell reference";
    case LXR_ERROR_UNSUPPORTED_FEATURE:  return "feature not yet implemented";
    case LXR_MAX_ERRNO:                  break;
    }
    return "unknown error";
}
