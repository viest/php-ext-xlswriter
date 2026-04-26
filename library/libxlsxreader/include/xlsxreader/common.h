#ifndef LXR_COMMON_H
#define LXR_COMMON_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef enum {
    LXR_NO_ERROR = 0,
    LXR_ERROR_MEMORY_MALLOC_FAILED,
    LXR_ERROR_FILE_OPEN_FAILED,
    LXR_ERROR_FILE_NOT_XLSX,
    LXR_ERROR_FILE_CORRUPTED,
    LXR_ERROR_ZIP_ENTRY_NOT_FOUND,
    LXR_ERROR_XML_PARSE,
    LXR_ERROR_SHEET_NOT_FOUND,
    LXR_ERROR_NULL_PARAMETER,
    LXR_ERROR_END_OF_DATA,
    LXR_ERROR_INVALID_CELL_REF,
    LXR_ERROR_UNSUPPORTED_FEATURE,
    LXR_MAX_ERRNO
} lxr_error;

const char *lxr_strerror(lxr_error code);

typedef struct {
    const char *ptr;
    size_t      len;
} lxr_str;

typedef enum {
    LXR_CELL_BLANK = 0,
    LXR_CELL_NUMBER,
    LXR_CELL_DATETIME,
    LXR_CELL_STRING,
    LXR_CELL_BOOLEAN,
    LXR_CELL_FORMULA,
    LXR_CELL_ERROR,
    LXR_CELL_INLINE_STRING
} lxr_cell_type;

typedef struct {
    size_t        row;
    size_t        col;
    lxr_cell_type type;
    uint32_t      style_id;

    union {
        double  number;
        int64_t unix_timestamp;
        lxr_str string;
        int     boolean;
        struct {
            lxr_str formula;
            lxr_str cached;
        }       formula;
        char    error_code[8];
    } value;

    lxr_str raw;
} lxr_cell;

#define LXR_SKIP_NONE        0x00
#define LXR_SKIP_EMPTY_ROWS  0x01
#define LXR_SKIP_EMPTY_CELLS 0x02
#define LXR_SKIP_ALL_EMPTY   0x03
#define LXR_SKIP_EXTRA_CELLS 0x04
#define LXR_SKIP_HIDDEN_ROWS 0x08

typedef enum {
    LXR_SST_MODE_FULL = 0,
    LXR_SST_MODE_STREAMING
} lxr_sst_mode;

#ifdef __cplusplus
}
#endif

#endif
