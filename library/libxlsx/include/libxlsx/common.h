/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 */

/**
 * @file common.h
 *
 * @brief Common functions and defines for the libxlsxwriter library.
 *
 * <!-- Copyright 2014-2026, John McNamara, jmcnamara@cpan.org -->
 *
 */
#ifndef __LXLSX_COMMON_H__
#define __LXLSX_COMMON_H__

#include <stddef.h>
#include <stdint.h>
#include <stdio.h>
#include <time.h>
#include "third_party/queue.h"
#include "third_party/tree.h"

#ifndef TESTING
#define STATIC static
#else
#define STATIC
#endif

#if __GNUC__ >= 5
#define DEPRECATED(func, msg) func __attribute__ ((deprecated(msg)))
#elif defined(_MSC_VER)
#define DEPRECATED(func, msg) __declspec(deprecated, msg) func
#else
#define DEPRECATED(func, msg) func
#endif

/** Integer data type to represent a row value. Equivalent to `uint32_t`.
 *
 * The maximum row in Excel is 1,048,576.
 */
typedef uint32_t lxlsx_row_t;

/** Integer data type to represent a column value. Equivalent to `uint16_t`.
 *
 * The maximum column in Excel is 16,384.
 */
typedef uint16_t lxlsx_col_t;

/** Boolean values used in libxlsxwriter. */
enum lxlsx_boolean {
    /** False value. */
    LXLSX_FALSE,
    /** True value. */
    LXLSX_TRUE,
    /** False value. Used to turn off a property that is default on, in order
     *  to distinguish it from an uninitialized value. */
    LXLSX_EXPLICIT_FALSE
};

/**
 * @brief Error codes from libxlsxwriter functions.
 *
 * See the `lxlsx_strerror()` function for an example of how to convert the
 * enum number to a descriptive error message string.
 */
typedef enum lxlsx_error {

    /** No error. */
    LXLSX_NO_ERROR = 0,

    /** Memory error, failed to malloc() required memory. */
    LXLSX_ERROR_MEMORY_MALLOC_FAILED,

    /** Error creating output xlsx file. Usually a permissions error. */
    LXLSX_ERROR_CREATING_XLSX_FILE,

    /** Error encountered when creating a tmpfile during file assembly. */
    LXLSX_ERROR_CREATING_TMPFILE,

    /** Error reading a tmpfile. */
    LXLSX_ERROR_READING_TMPFILE,

    /** Zip generic error ZIP_ERRNO while creating the xlsx file. */
    LXLSX_ERROR_ZIP_FILE_OPERATION,

    /** Zip error ZIP_PARAMERROR while creating the xlsx file. */
    LXLSX_ERROR_ZIP_PARAMETER_ERROR,

    /** Zip error ZIP_BADZIPFILE (use_zip64 option may be required). */
    LXLSX_ERROR_ZIP_BAD_ZIP_FILE,

    /** Zip error ZIP_INTERNALERROR while creating the xlsx file. */
    LXLSX_ERROR_ZIP_INTERNAL_ERROR,

    /** File error or unknown zip error when adding sub file to xlsx file. */
    LXLSX_ERROR_ZIP_FILE_ADD,

    /** Unknown zip error when closing xlsx file. */
    LXLSX_ERROR_ZIP_CLOSE,

    /** Feature is not currently supported in this configuration. */
    LXLSX_ERROR_FEATURE_NOT_SUPPORTED,

    /** NULL function parameter ignored. */
    LXLSX_ERROR_NULL_PARAMETER_IGNORED,

    /** Function parameter validation error. */
    LXLSX_ERROR_PARAMETER_VALIDATION,

    /** Function string parameter is empty. */
    LXLSX_ERROR_PARAMETER_IS_EMPTY,

    /** A #lxlsx_datetime parameter has a validation error. */
    LXLSX_ERROR_DATETIME_VALIDATION,

    /** Worksheet name exceeds Excel's limit of 31 characters. */
    LXLSX_ERROR_SHEETNAME_LENGTH_EXCEEDED,

    /** Worksheet name cannot contain invalid characters: '[ ] : * ? / \\' */
    LXLSX_ERROR_INVALID_SHEETNAME_CHARACTER,

    /** Worksheet name cannot start or end with an apostrophe. */
    LXLSX_ERROR_SHEETNAME_START_END_APOSTROPHE,

    /** Worksheet name is already in use. */
    LXLSX_ERROR_SHEETNAME_ALREADY_USED,

    /** Parameter exceeds Excel's limit of 32 characters. */
    LXLSX_ERROR_32_STRING_LENGTH_EXCEEDED,

    /** Parameter exceeds Excel's limit of 128 characters. */
    LXLSX_ERROR_128_STRING_LENGTH_EXCEEDED,

    /** Parameter exceeds Excel's limit of 255 characters. */
    LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED,

    /** String exceeds Excel's limit of 32,767 characters. */
    LXLSX_ERROR_MAX_STRING_LENGTH_EXCEEDED,

    /** Error finding internal string index. */
    LXLSX_ERROR_SHARED_STRING_INDEX_NOT_FOUND,

    /** Worksheet row or column index out of range. */
    LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE,

    /** Maximum hyperlink length (2079) exceeded. */
    LXLSX_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED,

    /** Maximum number of worksheet URLs (65530) exceeded. */
    LXLSX_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED,

    /** Couldn't read image dimensions or DPI. */
    LXLSX_ERROR_IMAGE_DIMENSIONS,

    LXLSX_MAX_ERRNO
} lxlsx_error;

/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
typedef struct lxlsx_datetime {

    /** Year     : 1900 - 9999 */
    int year;
    /** Month    : 1 - 12 */
    int month;
    /** Day      : 1 - 31 */
    int day;
    /** Hour     : 0 - 23 */
    int hour;
    /** Minute   : 0 - 59 */
    int min;
    /** Seconds  : 0 - 59.999 */
    double sec;

} lxlsx_datetime;

enum lxlsx_custom_property_types {
    LXLSX_CUSTOM_NONE,
    LXLSX_CUSTOM_STRING,
    LXLSX_CUSTOM_DOUBLE,
    LXLSX_CUSTOM_INTEGER,
    LXLSX_CUSTOM_BOOLEAN,
    LXLSX_CUSTOM_DATETIME
};

/* Size of MD5 byte arrays. */
#define LXLSX_MD5_SIZE              16

/* Excel sheetname max of 31 chars. */
#define LXLSX_SHEETNAME_MAX         31

/* Max with all worksheet chars 4xUTF-8 bytes + start and end quotes + \0. */
#define LXLSX_MAX_SHEETNAME_LENGTH  ((LXLSX_SHEETNAME_MAX * 4) + 2 + 1)

/* Max col string length. */
#define LXLSX_MAX_COL_NAME_LENGTH   sizeof("$XFD")

/* Max row string length. */
#define LXLSX_MAX_ROW_NAME_LENGTH   sizeof("$1048576")

/* Max cell string length. */
#define LXLSX_MAX_CELL_NAME_LENGTH  sizeof("$XFWD$1048576")

/* Max range: $XFWD$1048576:$XFWD$1048576\0 */
#define LXLSX_MAX_CELL_RANGE_LENGTH (LXLSX_MAX_CELL_NAME_LENGTH * 2)

/* Max range formula Sheet1!$A$1:$C$5$ style. */
#define LXLSX_MAX_FORMULA_RANGE_LENGTH (LXLSX_MAX_SHEETNAME_LENGTH + LXLSX_MAX_CELL_RANGE_LENGTH)

/* Datetime string length. */
#define LXLSX_DATETIME_LENGTH       sizeof("2016-12-12T23:00:00Z")

/* GUID string length. */
#define LXLSX_GUID_LENGTH           sizeof("{12345678-1234-1234-1234-1234567890AB}\0")

#define LXLSX_UINT32_T_LENGTH       sizeof("4294967296")
#define LXLSX_FILENAME_LENGTH       128
#define LXLSX_IGNORE                1

#define LXLSX_PORTRAIT              1
#define LXLSX_LANDSCAPE             0

#define LXLSX_SCHEMA_MS        "http://schemas.microsoft.com/office/2006/relationships"
#define LXLSX_SCHEMA_ROOT      "http://schemas.openxmlformats.org"
#define LXLSX_SCHEMA_DRAWING   LXLSX_SCHEMA_ROOT "/drawingml/2006"
#define LXLSX_SCHEMA_OFFICEDOC LXLSX_SCHEMA_ROOT "/officeDocument/2006"
#define LXLSX_SCHEMA_PACKAGE   LXLSX_SCHEMA_ROOT "/package/2006/relationships"
#define LXLSX_SCHEMA_DOCUMENT  LXLSX_SCHEMA_ROOT "/officeDocument/2006/relationships"
#define LXLSX_SCHEMA_CONTENT   LXLSX_SCHEMA_ROOT "/package/2006/content-types"

/* Use REprintf() for error handling when compiled as an R library. */
#ifdef USE_R_LANG
#include <R.h>
#define LXLSX_PRINTF REprintf
#define LXLSX_STDERR
#else
#define LXLSX_PRINTF fprintf
#define LXLSX_STDERR stderr,
#endif

#define LXLSX_ERROR(message)                      \
    LXLSX_PRINTF(LXLSX_STDERR "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

#define LXLSX_MEM_ERROR()                         \
    LXLSX_ERROR("Memory allocation failed.")

#define GOTO_LABEL_ON_MEM_ERROR(pointer, label) \
    do {                                        \
        if (!pointer) {                         \
            LXLSX_MEM_ERROR();                    \
            goto label;                         \
        }                                       \
    } while (0)

#define RETURN_ON_MEM_ERROR(pointer, error)     \
    do {                                        \
        if (!pointer) {                         \
            LXLSX_MEM_ERROR();                    \
            return error;                       \
        }                                       \
    } while (0)

#define RETURN_VOID_ON_MEM_ERROR(pointer)       \
    do {                                        \
        if (!pointer) {                         \
            LXLSX_MEM_ERROR();                    \
            return;                             \
        }                                       \
    } while (0)

#define RETURN_ON_ERROR(error)                  \
    do {                                        \
        if (error)                              \
            return error;                       \
    } while (0)

#define RETURN_AND_ZIPCLOSE_ON_ERROR(error)     \
        do {                                    \
            if (error) {                        \
                zipClose(self->zipfile, NULL);  \
                return error;                   \
            }                                   \
        } while (0)

#define LXLSX_WARN(message)                       \
    LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: " message "\n")

/* We can't use variadic macros here since we support ANSI C. */
#define LXLSX_WARN_FORMAT(message)                \
    LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: " message "\n")

#define LXLSX_WARN_FORMAT1(message, var)          \
    LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: " message "\n", var)

#define LXLSX_WARN_FORMAT2(message, var1, var2)    \
    LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: " message "\n", var1, var2)

#define LXLSX_WARN_FORMAT3(message, var1, var2, var3) \
    LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: " message "\n", var1, var2, var3)

/* Chart axis type checks. */
#define LXLSX_WARN_CAT_AXIS_ONLY(function)                                   \
    do {                                                                   \
        if (!axis->is_category) {                                          \
            LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: "                            \
                    function "() is only valid for category axes\n");      \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXLSX_WARN_VALUE_AXIS_ONLY(function)                                 \
    do {                                                                   \
        if (!axis->is_value) {                                             \
            LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: "                            \
                function "() is only valid for value axes\n");             \
                return;                                                    \
        }                                                                  \
    } while (0)

#define LXLSX_WARN_DATE_AXIS_ONLY(function)                                  \
    do {                                                                   \
        if (!axis->is_date) {                                              \
            LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: "                            \
                    function "() is only valid for date axes\n");          \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXLSX_WARN_CAT_AND_DATE_AXIS_ONLY(function)                          \
    do {                                                                   \
        if (!axis->is_category && !axis->is_date) {                        \
            LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: "                            \
                function "() is only valid for category and date axes\n"); \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXLSX_WARN_VALUE_AND_DATE_AXIS_ONLY(function)                        \
    do {                                                                   \
        if (!axis->is_value && !axis->is_date) {                           \
            LXLSX_PRINTF(LXLSX_STDERR "[WARNING]: "                            \
                function "() is only valid for value and date axes\n");    \
            return;                                                        \
        }                                                                  \
    } while (0)

#ifndef LXLSX_BIG_ENDIAN
#define LXLSX_UINT16_HOST(n)    (n)
#define LXLSX_UINT32_HOST(n)    (n)
#define LXLSX_UINT16_NETWORK(n) ((((n) & 0x00FF) << 8) | (((n) & 0xFF00) >> 8))
#define LXLSX_UINT32_NETWORK(n) ((((n) & 0xFF)       << 24) | \
                               (((n) & 0xFF00)     <<  8) | \
                               (((n) & 0xFF0000)   >>  8) | \
                               (((n) & 0xFF000000) >> 24))
#else
#define LXLSX_UINT16_NETWORK(n) (n)
#define LXLSX_UINT32_NETWORK(n) (n)
#define LXLSX_UINT16_HOST(n)    ((((n) & 0x00FF) << 8) | (((n) & 0xFF00) >> 8))
#define LXLSX_UINT32_HOST(n)    ((((n) & 0xFF)       << 24) | \
                               (((n) & 0xFF00)     <<  8) | \
                               (((n) & 0xFF0000)   >>  8) | \
                               (((n) & 0xFF000000) >> 24))
#endif

/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/* Compilers that have a native snprintf() can use it directly. */
#ifdef _MSC_VER
#define LXLSX_HAS_SNPRINTF
#endif

#ifdef LXLSX_HAS_SNPRINTF
#define lxlsx_snprintf snprintf
#else
#define lxlsx_snprintf __builtin_snprintf
#endif

/* Define a snprintf for MSVC 2010. */
#if defined(_MSC_VER) && _MSC_VER < 1900

#include <stdarg.h>
#define snprintf msvc2010_snprintf
#define vsnprintf msvc2010_vsnprintf

__inline int
msvc2010_vsnprintf(char *str, size_t size, const char *format, va_list ap)
{
    int count = -1;

    if (size != 0)
        count = _vsnprintf_s(str, size, _TRUNCATE, format, ap);
    if (count == -1)
        count = _vscprintf(format, ap);

    return count;
}

__inline int
msvc2010_snprintf(char *str, size_t size, const char *format, ...)
{
    int count;
    va_list ap;

    va_start(ap, format);
    count = msvc2010_vsnprintf(str, size, format, ap);
    va_end(ap);

    return count;
}

#endif

/* Safer strcpy for fixed width char arrays. */
#define lxlsx_strcpy(dest, src) \
    lxlsx_snprintf(dest, sizeof(dest), "%s", src)

/* Define the queue.h structs for the formats list. */
STAILQ_HEAD(lxlsx_formats, lxlsx_format);

/* Define the queue.h structs for the generic data structs. */
STAILQ_HEAD(lxlsx_tuples, lxlsx_tuple);
STAILQ_HEAD(lxlsx_custom_properties, lxlsx_custom_property);

typedef struct lxlsx_tuple {
    char *key;
    char *value;

    STAILQ_ENTRY (lxlsx_tuple) list_pointers;
} lxlsx_tuple;

/* Define custom property used in workbook.c and custom.c. */
typedef struct lxlsx_custom_property {

    enum lxlsx_custom_property_types type;
    char *name;

    union {
        char *string;
        double number;
        int32_t integer;
        uint8_t boolean;
        lxlsx_datetime datetime;
    } u;

    STAILQ_ENTRY (lxlsx_custom_property) list_pointers;

} lxlsx_custom_property;

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_COMMON_H__ */


/* XLSX read common API. */

#ifndef LXLSX_READER_COMMON_H
#define LXLSX_READER_COMMON_H

#include <stddef.h>
#include <stdint.h>

#include "libxlsx/cell.h"

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
/* When set, worksheet metadata is not preloaded before the first data read.
 * This preserves the lowest-latency pure streaming path, but metadata
 * accessors called while the row stream is active may return empty results.
 * Default OFF makes metadata access independent of call order. */
#define LXLSX_READER_DEFER_METADATA     0x40

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
