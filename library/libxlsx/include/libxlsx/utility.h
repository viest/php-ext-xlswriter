/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 */

/**
 * @file utility.h
 *
 * @brief Utility functions for libxlsxwriter.
 *
 * <!-- Copyright 2014-2026, John McNamara, jmcnamara@cpan.org -->
 *
 */

#ifndef __LXLSX_UTILITY_H__
#define __LXLSX_UTILITY_H__

#include <stdint.h>
#ifndef _MSC_VER
#include <strings.h>
#endif
#include "common.h"
#include "xmlwriter.h"

/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      lxlsx_worksheet_write_string(worksheet, CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      lxlsx_worksheet_write_string(worksheet, 0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
#define CELL(cell) \
    lxlsx_name_to_row(cell), lxlsx_name_to_col(cell)

/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     lxlsx_worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     lxlsx_worksheet_set_column(worksheet, 1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
#define COLS(cols) \
    lxlsx_name_to_col(cols), lxlsx_name_to_col_2(cols)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     lxlsx_worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     lxlsx_worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 */
#define RANGE(range) \
    lxlsx_name_to_row(range), lxlsx_name_to_col(range), \
    lxlsx_name_to_row_2(range), lxlsx_name_to_col_2(range)

/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Retrieve the library version.
 *
 * @return The "X.Y.Z" version string.
 *
 * Get the library version as a "X.Y.Z" version string
 *
 *  @code
 *      printf("Libxlsxwriter version = %s\n", lxlsx_version());
 *  @endcode
 *
 */
const char *lxlsx_version(void);

/**
 * @brief Retrieve the library version ID.
 *
 * @return The version ID.
 *
 * Get the library version such as "X.Y.Z" as a XYZ integer.
 *
 *  @code
 *      printf("Libxlsxwriter version id = %d\n", lxlsx_version_id());
 *  @endcode
 *
 */
uint16_t lxlsx_version_id(void);

/**
 * @brief Converts a libxlsxwriter error number to a string.
 *
 * The `%lxlsx_strerror` function converts a libxlsxwriter error number defined
 * by #lxlsx_error to a pointer to a string description of the error.
 * Similar to the standard library strerror(3) function.
 *
 * For example:
 *
 * @code
 *     lxlsx_error error = lxlsx_workbook_close(workbook);
 *
 *     if (error)
 *         printf("Error in lxlsx_workbook_close().\n"
 *                "Error %d = %s\n", error, lxlsx_strerror(error));
 * @endcode
 *
 * This would produce output like the following if the target file wasn't
 * writable:
 *
 *     Error in lxlsx_workbook_close().
 *     Error 2 = Error creating output xlsx file. Usually a permissions error.
 *
 * @param error_num The error number returned by a libxlsxwriter function.
 *
 * @return A pointer to a statically allocated string. Do not free.
 */
char *lxlsx_strerror(lxlsx_error error_num);

/* Create a quoted version of the worksheet name */
char *lxlsx_quote_sheetname(const char *str);

void lxlsx_col_to_name(char *col_name, lxlsx_col_t col_num, uint8_t absolute);

void lxlsx_rowcol_to_cell(char *cell_name, lxlsx_row_t row, lxlsx_col_t col);

void lxlsx_rowcol_to_cell_abs(char *cell_name,
                            lxlsx_row_t row,
                            lxlsx_col_t col, uint8_t abs_row, uint8_t abs_col);

void lxlsx_rowcol_to_range(char *range,
                         lxlsx_row_t first_row, lxlsx_col_t first_col,
                         lxlsx_row_t last_row, lxlsx_col_t last_col);

void lxlsx_rowcol_to_range_abs(char *range,
                             lxlsx_row_t first_row, lxlsx_col_t first_col,
                             lxlsx_row_t last_row, lxlsx_col_t last_col);

void lxlsx_rowcol_to_formula_abs(char *formula, const char *sheetname,
                               lxlsx_row_t first_row, lxlsx_col_t first_col,
                               lxlsx_row_t last_row, lxlsx_col_t last_col);

uint32_t lxlsx_name_to_row(const char *row_str);
uint16_t lxlsx_name_to_col(const char *col_str);

uint32_t lxlsx_name_to_row_2(const char *row_str);
uint16_t lxlsx_name_to_col_2(const char *col_str);

/**
 * @brief Converts a #lxlsx_datetime to an Excel datetime number.
 *
 * @param datetime A pointer to a #lxlsx_datetime struct.
 *
 * @return A double representing an Excel datetime.
 *
 * The `%lxlsx_datetime_to_excel_datetime()` function converts a datetime in
 * #lxlsx_datetime to an Excel datetime number:
 *
 * @code
 *     lxlsx_datetime datetime = {2013, 2, 28, 12, 0, 0.0};
 *
 *     double excel_datetime = lxlsx_datetime_to_excel_date(&datetime);
 * @endcode
 *
 * See @ref working_with_dates for more details on the Excel datetime format.
 */
double lxlsx_datetime_to_excel_datetime(lxlsx_datetime *datetime);

/**
 * @brief Converts a #lxlsx_datetime to an Excel datetime number with 1900/1904
 * epoch.
 *
 * This function is similar to `lxlsx_datetime_to_excel_datetime()` but it allows
 * you to specify whether to use the 1900 or 1904 epoch. See also the
 * `lxlsx_workbook_use_1904_epoch()` function.
 *
 * @param datetime A pointer to a #lxlsx_datetime struct.
 * @param use_1904_epoch A flag to indicate whether to use the 1904 epoch (true)
 *        or the 1900 epoch (false).
 *
 */
double lxlsx_datetime_to_excel_date_with_epoch(lxlsx_datetime *datetime,
                                             uint8_t use_1904_epoch);

/**
 * @brief Validate a #lxlsx_datetime struct.
 *
 * Validates a #lxlsx_datetime struct to ensure its fields are within acceptable
 * ranges for Excel dates and times.
 *
 * The members of the #lxlsx_datetime struct and the range of their values are:
 *
 * Member   | Value
 * -------- | -----------
 * year     | 1900 - 9999
 * month    | 1 - 12
 * day      | 1 - 31
 * hour     | 0 - 23
 * min      | 0 - 59
 * sec      | 0 - 59.999
 *
 * @param datetime A pointer to a #lxlsx_datetime struct.
 *
 * @return A #lxlsx_error code. Either #LXLSX_NO_ERROR or
 *         #LXLSX_ERROR_DATETIME_VALIDATION if a field is out of range.
 */
lxlsx_error lxlsx_datetime_validate(lxlsx_datetime *datetime);

/**
 * @brief Converts a unix datetime to an Excel datetime number.
 *
 * @param unixtime Unix time (seconds since 1970-01-01)
 *
 * @return A double representing an Excel datetime.
 *
 * The `%lxlsx_unixtime_to_excel_date()` function converts a unix datetime to
 * an Excel datetime number:
 *
 * @code
 *     double excel_datetime = lxlsx_unixtime_to_excel_date(946684800);
 * @endcode
 *
 * See @ref working_with_dates for more details.
 */
double lxlsx_unixtime_to_excel_date(int64_t unixtime);

/**
 * @brief Converts a unix datetime to an Excel datetime number with 1900/1904
 * epoch.
 *
 * This function is similar to `lxlsx_unixtime_to_excel_date()` but it allows
 * you to specify whether to use the 1900 or 1904 epoch. See also the
 * `lxlsx_workbook_use_1904_epoch()` function.
 *
 * @param unixtime Unix time (seconds since 1970-01-01)
 * @param use_1904_epoch A flag to indicate whether to use the 1904 epoch (true)
 *        or the 1900 epoch (false).
 *
 */
double lxlsx_unixtime_to_excel_date_with_epoch(int64_t unixtime,
                                             uint8_t use_1904_epoch);

char *lxlsx_strdup(const char *str);
char *lxlsx_strdup_formula(const char *formula);
size_t lxlsx_utf8_strlen(const char *str);
void lxlsx_str_tolower(char *str);
uint8_t lxlsx_str_is_empty(const char *str);

/* Define a portable version of strcasecmp()/strncasecmp(). MSVC has no POSIX
 * case-insensitive compares; use the _stricmp/_strnicmp equivalents there. */
#ifdef _MSC_VER
#define lxlsx_strcasecmp  _stricmp
#define lxlsx_strncasecmp _strnicmp
#else
#define lxlsx_strcasecmp  strcasecmp
#define lxlsx_strncasecmp strncasecmp
#endif

FILE *lxlsx_tmpfile(const char *tmpdir);
FILE *lxlsx_get_filehandle(char **buf, size_t *size, const char *tmpdir);
FILE *lxlsx_fopen(const char *filename, const char *mode);

/* Use the third party dtoa function to avoid locale issues with sprintf
 * double formatting. Otherwise we use a simple macro that falls back to the
 * default c-lib sprintf.
 */
#ifdef USE_DTOA_LIBRARY
int lxlsx_sprintf_dbl(char *data, double number);
#else
#define lxlsx_sprintf_dbl(data, number) \
        lxlsx_snprintf(data, LXLSX_ATTR_32, "%.16G", number)
#endif

uint16_t lxlsx_hash_password(const char *password);

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_UTILITY_H__ */
