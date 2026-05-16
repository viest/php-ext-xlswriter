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

/* Formula sub-types per ECMA-376 §18.18.31 (ST_CellFormulaType). */
typedef enum {
    LXR_FORMULA_NORMAL    = 0,
    LXR_FORMULA_ARRAY     = 1,
    LXR_FORMULA_DATATABLE = 2,
    LXR_FORMULA_SHARED    = 3
} lxr_formula_kind;

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
            lxr_str          formula;     /* expression text (may be empty
                                             on a shared-formula follower) */
            lxr_str          cached;      /* cached <v> value */
            lxr_formula_kind kind;
            lxr_str          ref;         /* range for array/dataTable, e.g.
                                             "A1:B3"; empty otherwise */
            int              si;          /* shared index, or -1 */
            int              is_dynamic;  /* aca="1" → 1 */
        }       formula;
        char    error_code[8];
    } value;

    lxr_str raw;
} lxr_cell;

#define LXR_SKIP_NONE          0x00
#define LXR_SKIP_EMPTY_ROWS    0x01
#define LXR_SKIP_EMPTY_CELLS   0x02
#define LXR_SKIP_ALL_EMPTY     0x03
#define LXR_SKIP_EXTRA_CELLS   0x04
#define LXR_SKIP_HIDDEN_ROWS   0x08
/* When set, the non-master cells inside a merged region surface as null in
 * row-shaped read APIs (Excel::nextRow / nextRowWithFormula / getSheetData).
 * The master cell — i.e. (first_row, first_col) of the merge — is unaffected. */
#define LXR_SKIP_MERGED_FOLLOW 0x10
/* When set, Excel::nextRowWithFormula's "formula" field is an associative
 * array (type/text/ref/si/is_dynamic/cached_value) instead of a plain string.
 * Default OFF preserves existing string-shaped output. */
#define LXR_FORMULA_VERBOSE    0x20

typedef enum {
    LXR_SST_MODE_FULL = 0,
    LXR_SST_MODE_STREAMING
} lxr_sst_mode;

/* Inclusive 1-based cell range. (0, 0, 0, 0) means "unset". */
typedef struct {
    size_t first_row;
    size_t first_col;
    size_t last_row;
    size_t last_col;
} lxr_range;

/* A formatted run inside a rich-text string cell. All string pointers (text,
 * font_name, color) are owned by the worksheet/SST and remain valid until
 * lxr_workbook_close(). text/text_len contain the run's plain text;
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
} lxr_string_run;

#ifdef __cplusplus
}
#endif

#endif
