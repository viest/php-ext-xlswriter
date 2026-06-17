#ifndef LXLSX_READER_COMMON_H
#define LXLSX_READER_COMMON_H

#include <stddef.h>
#include <stdint.h>

#include "lxlsx/cell.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef enum {
    LXLSX_READER_NO_ERROR = 0,
    LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED,
    LXLSX_READER_ERROR_FILE_OPEN_FAILED,
    LXLSX_READER_ERROR_FILE_NOT_XLSX,
    LXLSX_READER_ERROR_FILE_CORRUPTED,
    LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND,
    LXLSX_READER_ERROR_XML_PARSE,
    LXLSX_READER_ERROR_SHEET_NOT_FOUND,
    LXLSX_READER_ERROR_NULL_PARAMETER,
    LXLSX_READER_ERROR_END_OF_DATA,
    LXLSX_READER_ERROR_INVALID_CELL_REF,
    LXLSX_READER_ERROR_UNSUPPORTED_FEATURE,
    LXLSX_READER_MAX_ERRNO
} lxlsx_reader_error;

const char *lxlsx_reader_strerror(lxlsx_reader_error code);

#define LXLSX_READER_SKIP_NONE          0x00
#define LXLSX_READER_SKIP_EMPTY_ROWS    0x01
#define LXLSX_READER_SKIP_EMPTY_CELLS   0x02
#define LXLSX_READER_SKIP_ALL_EMPTY     0x03
#define LXLSX_READER_SKIP_EXTRA_CELLS   0x04
#define LXLSX_READER_SKIP_HIDDEN_ROWS   0x08
/* When set, the non-master cells inside a merged region surface as null in
 * row-shaped read APIs (Excel::nextRow / nextRowWithFormula / getSheetData).
 * The master cell — i.e. (first_row, first_col) of the merge — is unaffected. */
#define LXLSX_READER_SKIP_MERGED_FOLLOW 0x10
/* When set, Excel::nextRowWithFormula's "formula" field is an associative
 * array (type/text/ref/si/is_dynamic/cached_value) instead of a plain string.
 * Default OFF preserves existing string-shaped output. */
#define LXLSX_READER_FORMULA_VERBOSE    0x20

typedef enum {
    LXLSX_READER_SST_MODE_FULL = 0,
    LXLSX_READER_SST_MODE_STREAMING
} lxlsx_reader_sst_mode;

/* Inclusive 1-based cell range. (0, 0, 0, 0) means "unset". */
typedef struct {
    size_t first_row;
    size_t first_col;
    size_t last_row;
    size_t last_col;
} lxlsx_reader_range;

/* A formatted run inside a rich-text string cell. All string pointers (text,
 * font_name, color) are owned by the worksheet/SST and remain valid until
 * lxlsx_reader_workbook_close(). text/text_len contain the run's plain text;
 * font_name and color may be NULL when unset. font_size is 0.0 when absent. */
typedef struct {
    const char *text;
    size_t      text_len;
    const char *font_name;       /* or NULL */
    double      font_size;
    int         bold;
    int         italic;
    int         strike;
    int         underline;       /* 0=none, 1=single, 2=double, 3=single-acc, 4=double-acc */
    const char *color;           /* RGB hex like "FF0000" or NULL */
} lxlsx_reader_string_run;

#ifdef __cplusplus
}
#endif

#endif
