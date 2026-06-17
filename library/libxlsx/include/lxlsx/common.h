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
#ifndef __LXW_COMMON_H__
#define __LXW_COMMON_H__

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
typedef uint32_t lxw_row_t;

/** Integer data type to represent a column value. Equivalent to `uint16_t`.
 *
 * The maximum column in Excel is 16,384.
 */
typedef uint16_t lxw_col_t;

/** Boolean values used in libxlsxwriter. */
enum lxw_boolean {
    /** False value. */
    LXW_FALSE,
    /** True value. */
    LXW_TRUE,
    /** False value. Used to turn off a property that is default on, in order
     *  to distinguish it from an uninitialized value. */
    LXW_EXPLICIT_FALSE
};

/**
 * @brief Error codes from libxlsxwriter functions.
 *
 * See the `lxw_strerror()` function for an example of how to convert the
 * enum number to a descriptive error message string.
 */
typedef enum lxw_error {

    /** No error. */
    LXW_NO_ERROR = 0,

    /** Memory error, failed to malloc() required memory. */
    LXW_ERROR_MEMORY_MALLOC_FAILED,

    /** Error creating output xlsx file. Usually a permissions error. */
    LXW_ERROR_CREATING_XLSX_FILE,

    /** Error encountered when creating a tmpfile during file assembly. */
    LXW_ERROR_CREATING_TMPFILE,

    /** Error reading a tmpfile. */
    LXW_ERROR_READING_TMPFILE,

    /** Zip generic error ZIP_ERRNO while creating the xlsx file. */
    LXW_ERROR_ZIP_FILE_OPERATION,

    /** Zip error ZIP_PARAMERROR while creating the xlsx file. */
    LXW_ERROR_ZIP_PARAMETER_ERROR,

    /** Zip error ZIP_BADZIPFILE (use_zip64 option may be required). */
    LXW_ERROR_ZIP_BAD_ZIP_FILE,

    /** Zip error ZIP_INTERNALERROR while creating the xlsx file. */
    LXW_ERROR_ZIP_INTERNAL_ERROR,

    /** File error or unknown zip error when adding sub file to xlsx file. */
    LXW_ERROR_ZIP_FILE_ADD,

    /** Unknown zip error when closing xlsx file. */
    LXW_ERROR_ZIP_CLOSE,

    /** Feature is not currently supported in this configuration. */
    LXW_ERROR_FEATURE_NOT_SUPPORTED,

    /** NULL function parameter ignored. */
    LXW_ERROR_NULL_PARAMETER_IGNORED,

    /** Function parameter validation error. */
    LXW_ERROR_PARAMETER_VALIDATION,

    /** Function string parameter is empty. */
    LXW_ERROR_PARAMETER_IS_EMPTY,

    /** A #lxw_datetime parameter has a validation error. */
    LXW_ERROR_DATETIME_VALIDATION,

    /** Worksheet name exceeds Excel's limit of 31 characters. */
    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED,

    /** Worksheet name cannot contain invalid characters: '[ ] : * ? / \\' */
    LXW_ERROR_INVALID_SHEETNAME_CHARACTER,

    /** Worksheet name cannot start or end with an apostrophe. */
    LXW_ERROR_SHEETNAME_START_END_APOSTROPHE,

    /** Worksheet name is already in use. */
    LXW_ERROR_SHEETNAME_ALREADY_USED,

    /** Parameter exceeds Excel's limit of 32 characters. */
    LXW_ERROR_32_STRING_LENGTH_EXCEEDED,

    /** Parameter exceeds Excel's limit of 128 characters. */
    LXW_ERROR_128_STRING_LENGTH_EXCEEDED,

    /** Parameter exceeds Excel's limit of 255 characters. */
    LXW_ERROR_255_STRING_LENGTH_EXCEEDED,

    /** String exceeds Excel's limit of 32,767 characters. */
    LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED,

    /** Error finding internal string index. */
    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND,

    /** Worksheet row or column index out of range. */
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE,

    /** Maximum hyperlink length (2079) exceeded. */
    LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED,

    /** Maximum number of worksheet URLs (65530) exceeded. */
    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED,

    /** Couldn't read image dimensions or DPI. */
    LXW_ERROR_IMAGE_DIMENSIONS,

    LXW_MAX_ERRNO
} lxw_error;

/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
typedef struct lxw_datetime {

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

} lxw_datetime;

enum lxw_custom_property_types {
    LXW_CUSTOM_NONE,
    LXW_CUSTOM_STRING,
    LXW_CUSTOM_DOUBLE,
    LXW_CUSTOM_INTEGER,
    LXW_CUSTOM_BOOLEAN,
    LXW_CUSTOM_DATETIME
};

/* Size of MD5 byte arrays. */
#define LXW_MD5_SIZE              16

/* Excel sheetname max of 31 chars. */
#define LXW_SHEETNAME_MAX         31

/* Max with all worksheet chars 4xUTF-8 bytes + start and end quotes + \0. */
#define LXW_MAX_SHEETNAME_LENGTH  ((LXW_SHEETNAME_MAX * 4) + 2 + 1)

/* Max col string length. */
#define LXW_MAX_COL_NAME_LENGTH   sizeof("$XFD")

/* Max row string length. */
#define LXW_MAX_ROW_NAME_LENGTH   sizeof("$1048576")

/* Max cell string length. */
#define LXW_MAX_CELL_NAME_LENGTH  sizeof("$XFWD$1048576")

/* Max range: $XFWD$1048576:$XFWD$1048576\0 */
#define LXW_MAX_CELL_RANGE_LENGTH (LXW_MAX_CELL_NAME_LENGTH * 2)

/* Max range formula Sheet1!$A$1:$C$5$ style. */
#define LXW_MAX_FORMULA_RANGE_LENGTH (LXW_MAX_SHEETNAME_LENGTH + LXW_MAX_CELL_RANGE_LENGTH)

/* Datetime string length. */
#define LXW_DATETIME_LENGTH       sizeof("2016-12-12T23:00:00Z")

/* GUID string length. */
#define LXW_GUID_LENGTH           sizeof("{12345678-1234-1234-1234-1234567890AB}\0")

#define LXW_UINT32_T_LENGTH       sizeof("4294967296")
#define LXW_FILENAME_LENGTH       128
#define LXW_IGNORE                1

#define LXW_PORTRAIT              1
#define LXW_LANDSCAPE             0

#define LXW_SCHEMA_MS        "http://schemas.microsoft.com/office/2006/relationships"
#define LXW_SCHEMA_ROOT      "http://schemas.openxmlformats.org"
#define LXW_SCHEMA_DRAWING   LXW_SCHEMA_ROOT "/drawingml/2006"
#define LXW_SCHEMA_OFFICEDOC LXW_SCHEMA_ROOT "/officeDocument/2006"
#define LXW_SCHEMA_PACKAGE   LXW_SCHEMA_ROOT "/package/2006/relationships"
#define LXW_SCHEMA_DOCUMENT  LXW_SCHEMA_ROOT "/officeDocument/2006/relationships"
#define LXW_SCHEMA_CONTENT   LXW_SCHEMA_ROOT "/package/2006/content-types"

/* Use REprintf() for error handling when compiled as an R library. */
#ifdef USE_R_LANG
#include <R.h>
#define LXW_PRINTF REprintf
#define LXW_STDERR
#else
#define LXW_PRINTF fprintf
#define LXW_STDERR stderr,
#endif

#define LXW_ERROR(message)                      \
    LXW_PRINTF(LXW_STDERR "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

#define LXW_MEM_ERROR()                         \
    LXW_ERROR("Memory allocation failed.")

#define GOTO_LABEL_ON_MEM_ERROR(pointer, label) \
    do {                                        \
        if (!pointer) {                         \
            LXW_MEM_ERROR();                    \
            goto label;                         \
        }                                       \
    } while (0)

#define RETURN_ON_MEM_ERROR(pointer, error)     \
    do {                                        \
        if (!pointer) {                         \
            LXW_MEM_ERROR();                    \
            return error;                       \
        }                                       \
    } while (0)

#define RETURN_VOID_ON_MEM_ERROR(pointer)       \
    do {                                        \
        if (!pointer) {                         \
            LXW_MEM_ERROR();                    \
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

#define LXW_WARN(message)                       \
    LXW_PRINTF(LXW_STDERR "[WARNING]: " message "\n")

/* We can't use variadic macros here since we support ANSI C. */
#define LXW_WARN_FORMAT(message)                \
    LXW_PRINTF(LXW_STDERR "[WARNING]: " message "\n")

#define LXW_WARN_FORMAT1(message, var)          \
    LXW_PRINTF(LXW_STDERR "[WARNING]: " message "\n", var)

#define LXW_WARN_FORMAT2(message, var1, var2)    \
    LXW_PRINTF(LXW_STDERR "[WARNING]: " message "\n", var1, var2)

#define LXW_WARN_FORMAT3(message, var1, var2, var3) \
    LXW_PRINTF(LXW_STDERR "[WARNING]: " message "\n", var1, var2, var3)

/* Chart axis type checks. */
#define LXW_WARN_CAT_AXIS_ONLY(function)                                   \
    do {                                                                   \
        if (!axis->is_category) {                                          \
            LXW_PRINTF(LXW_STDERR "[WARNING]: "                            \
                    function "() is only valid for category axes\n");      \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXW_WARN_VALUE_AXIS_ONLY(function)                                 \
    do {                                                                   \
        if (!axis->is_value) {                                             \
            LXW_PRINTF(LXW_STDERR "[WARNING]: "                            \
                function "() is only valid for value axes\n");             \
                return;                                                    \
        }                                                                  \
    } while (0)

#define LXW_WARN_DATE_AXIS_ONLY(function)                                  \
    do {                                                                   \
        if (!axis->is_date) {                                              \
            LXW_PRINTF(LXW_STDERR "[WARNING]: "                            \
                    function "() is only valid for date axes\n");          \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXW_WARN_CAT_AND_DATE_AXIS_ONLY(function)                          \
    do {                                                                   \
        if (!axis->is_category && !axis->is_date) {                        \
            LXW_PRINTF(LXW_STDERR "[WARNING]: "                            \
                function "() is only valid for category and date axes\n"); \
           return;                                                         \
        }                                                                  \
    } while (0)

#define LXW_WARN_VALUE_AND_DATE_AXIS_ONLY(function)                        \
    do {                                                                   \
        if (!axis->is_value && !axis->is_date) {                           \
            LXW_PRINTF(LXW_STDERR "[WARNING]: "                            \
                function "() is only valid for value and date axes\n");    \
            return;                                                        \
        }                                                                  \
    } while (0)

#ifndef LXW_BIG_ENDIAN
#define LXW_UINT16_HOST(n)    (n)
#define LXW_UINT32_HOST(n)    (n)
#define LXW_UINT16_NETWORK(n) ((((n) & 0x00FF) << 8) | (((n) & 0xFF00) >> 8))
#define LXW_UINT32_NETWORK(n) ((((n) & 0xFF)       << 24) | \
                               (((n) & 0xFF00)     <<  8) | \
                               (((n) & 0xFF0000)   >>  8) | \
                               (((n) & 0xFF000000) >> 24))
#else
#define LXW_UINT16_NETWORK(n) (n)
#define LXW_UINT32_NETWORK(n) (n)
#define LXW_UINT16_HOST(n)    ((((n) & 0x00FF) << 8) | (((n) & 0xFF00) >> 8))
#define LXW_UINT32_HOST(n)    ((((n) & 0xFF)       << 24) | \
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
#define LXW_HAS_SNPRINTF
#endif

#ifdef LXW_HAS_SNPRINTF
#define lxw_snprintf snprintf
#else
#define lxw_snprintf __builtin_snprintf
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
#define lxw_strcpy(dest, src) \
    lxw_snprintf(dest, sizeof(dest), "%s", src)

/* Define the queue.h structs for the formats list. */
STAILQ_HEAD(lxw_formats, lxw_format);

/* Define the queue.h structs for the generic data structs. */
STAILQ_HEAD(lxw_tuples, lxw_tuple);
STAILQ_HEAD(lxw_custom_properties, lxw_custom_property);

typedef struct lxw_tuple {
    char *key;
    char *value;

    STAILQ_ENTRY (lxw_tuple) list_pointers;
} lxw_tuple;

/* Define custom property used in workbook.c and custom.c. */
typedef struct lxw_custom_property {

    enum lxw_custom_property_types type;
    char *name;

    union {
        char *string;
        double number;
        int32_t integer;
        uint8_t boolean;
        lxw_datetime datetime;
    } u;

    STAILQ_ENTRY (lxw_custom_property) list_pointers;

} lxw_custom_property;

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_COMMON_H__ */
