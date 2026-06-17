#include "libxlsx/common.h"

const char *lxlsx_reader_strerror(lxlsx_reader_error code)
{
    switch (code) {
    case LXLSX_READER_NO_ERROR:                   return "no error";
    case LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED: return "memory allocation failed";
    case LXLSX_READER_ERROR_FILE_OPEN_FAILED:     return "failed to open file";
    case LXLSX_READER_ERROR_FILE_NOT_XLSX:        return "file is not a valid XLSX archive";
    case LXLSX_READER_ERROR_FILE_CORRUPTED:       return "file is corrupted";
    case LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND:  return "zip entry not found";
    case LXLSX_READER_ERROR_XML_PARSE:            return "XML parse error";
    case LXLSX_READER_ERROR_SHEET_NOT_FOUND:      return "sheet not found";
    case LXLSX_READER_ERROR_NULL_PARAMETER:       return "null parameter";
    case LXLSX_READER_ERROR_END_OF_DATA:          return "end of data";
    case LXLSX_READER_ERROR_INVALID_CELL_REF:     return "invalid cell reference";
    case LXLSX_READER_ERROR_UNSUPPORTED_FEATURE:  return "feature not yet implemented";
    case LXLSX_READER_MAX_ERRNO:                  break;
    }
    return "unknown error";
}
