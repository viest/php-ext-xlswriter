#ifndef LXLSX_READER_COMMON_H
#define LXLSX_READER_COMMON_H

#include <stddef.h>
#include <stdint.h>

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

typedef struct {
    const char *ptr;
    size_t      len;
} lxlsx_reader_str;

typedef enum {
    LXLSX_READER_CELL_BLANK = 0,
    LXLSX_READER_CELL_NUMBER,
    LXLSX_READER_CELL_DATETIME,
    LXLSX_READER_CELL_STRING,
    LXLSX_READER_CELL_BOOLEAN,
    LXLSX_READER_CELL_FORMULA,
    LXLSX_READER_CELL_ERROR,
    LXLSX_READER_CELL_INLINE_STRING
} lxlsx_reader_cell_type;

/* Formula sub-types per ECMA-376 §18.18.31 (ST_CellFormulaType). */
typedef enum {
    LXLSX_READER_FORMULA_NORMAL    = 0,
    LXLSX_READER_FORMULA_ARRAY     = 1,
    LXLSX_READER_FORMULA_DATATABLE = 2,
    LXLSX_READER_FORMULA_SHARED    = 3
} lxlsx_reader_formula_kind;

typedef struct {
    size_t        row;
    size_t        col;
    lxlsx_reader_cell_type type;
    uint32_t      style_id;

    union {
        double  number;
        int64_t unix_timestamp;
        lxlsx_reader_str string;
        int     boolean;
        struct {
            lxlsx_reader_str          formula;     /* expression text (may be empty
                                             on a shared-formula follower) */
            lxlsx_reader_str          cached;      /* cached <v> value */
            lxlsx_reader_formula_kind kind;
            lxlsx_reader_str          ref;         /* range for array/dataTable, e.g.
                                             "A1:B3"; empty otherwise */
            int              si;          /* shared index, or -1 */
            int              is_dynamic;  /* aca="1" → 1 */
        }       formula;
        char    error_code[8];
    } value;

    lxlsx_reader_str raw;
} lxlsx_reader_cell;

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
