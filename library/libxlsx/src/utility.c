/*****************************************************************************
 * utility - Utility functions for libxlsxwriter.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#ifdef USE_FMEMOPEN
#define _POSIX_C_SOURCE 200809L
#endif

#include <ctype.h>
#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include <stdlib.h>
#include "libxlsx.h"
#include "libxlsx/common.h"
#include "libxlsx/third_party/tmpfileplus.h"

#define LXLSX_TMPFILE_BUFFER_SIZE (64 * 1024)

#ifdef USE_DTOA_LIBRARY
#include "libxlsx/third_party/emyg_dtoa.h"
#endif

char *error_strings[LXLSX_MAX_ERRNO + 1] = {
    "No error.",
    "Memory error, failed to malloc() required memory.",
    "Error creating output xlsx file. Usually a permissions error.",
    "Error encountered when creating a tmpfile during file assembly.",
    "Error reading a tmpfile.",
    "Zip generic error ZIP_ERRNO while creating the xlsx file.",
    "Zip error ZIP_PARAMERROR while creating the xlsx file.",
    "Zip error ZIP_BADZIPFILE (use_zip64 option may be required).",
    "Zip error ZIP_INTERNALERROR while creating the xlsx file.",
    "File error or unknown zip error when adding sub file to xlsx file.",
    "Unknown zip error when closing xlsx file.",
    "Feature is not currently supported in this configuration.",
    "NULL function parameter ignored.",
    "Function parameter validation error.",
    "Function string parameter is empty.",
    "Datetime struct parameter has an invalid field value.",
    "Worksheet name exceeds Excel's limit of 31 characters.",
    "Worksheet name cannot contain invalid characters: '[ ] : * ? / \\'",
    "Worksheet name cannot start or end with an apostrophe.",
    "Worksheet name is already in use.",
    "Parameter exceeds Excel's limit of 32 characters.",
    "Parameter exceeds Excel's limit of 128 characters.",
    "Parameter exceeds Excel's limit of 255 characters.",
    "String exceeds Excel's limit of 32,767 characters.",
    "Error finding internal string index.",
    "Worksheet row or column index out of range.",
    "Maximum hyperlink length (2079) exceeded.",
    "Maximum number of worksheet URLs (65530) exceeded.",
    "Couldn't read image dimensions or DPI.",
    "Unknown error number."
};

char *
lxlsx_strerror(lxlsx_error error_num)
{
    if (error_num > LXLSX_MAX_ERRNO)
        error_num = LXLSX_MAX_ERRNO;

    return error_strings[error_num];
}

/*
 * Convert Excel A-XFD style column name to zero based number.
 */
void
lxlsx_col_to_name(char *col_name, lxlsx_col_t col_num, uint8_t absolute)
{
    uint8_t pos = 0;
    size_t len;
    size_t i;

    /* Change from 0 index to 1 index. */
    col_num++;

    /* Convert the column number to a string in reverse order. */
    while (col_num) {

        /* Get the remainder in base 26. */
        int remainder = col_num % 26;

        if (remainder == 0)
            remainder = 26;

        /* Convert the remainder value to a character. */
        col_name[pos++] = 'A' + remainder - 1;
        col_name[pos] = '\0';

        /* Get the next order of magnitude. */
        col_num = (col_num - 1) / 26;
    }

    if (absolute) {
        col_name[pos] = '$';
        col_name[pos + 1] = '\0';
    }

    /* Reverse the column name string. */
    len = strlen(col_name);
    for (i = 0; i < (len / 2); i++) {
        char tmp = col_name[i];
        col_name[i] = col_name[len - i - 1];
        col_name[len - i - 1] = tmp;
    }
}

/*
 * Convert zero indexed row and column to an Excel style A1 cell reference.
 */
void
lxlsx_rowcol_to_cell(char *cell_name, lxlsx_row_t row, lxlsx_col_t col)
{
    size_t pos;

    /* Add the column to the cell. */
    lxlsx_col_to_name(cell_name, col, 0);

    /* Get the end of the cell. */
    pos = strlen(cell_name);

    /* Add the row to the cell. */
    lxlsx_snprintf(&cell_name[pos], LXLSX_MAX_ROW_NAME_LENGTH, "%d", ++row);
}

/*
 * Convert zero indexed row and column to an Excel style $A$1 cell with
 * an absolute reference.
 */
void
lxlsx_rowcol_to_cell_abs(char *cell_name, lxlsx_row_t row, lxlsx_col_t col,
                       uint8_t abs_row, uint8_t abs_col)
{
    size_t pos;

    /* Add the column to the cell. */
    lxlsx_col_to_name(cell_name, col, abs_col);

    /* Get the end of the cell. */
    pos = strlen(cell_name);

    if (abs_row)
        cell_name[pos++] = '$';

    /* Add the row to the cell. */
    lxlsx_snprintf(&cell_name[pos], LXLSX_MAX_ROW_NAME_LENGTH, "%d", ++row);
}

/*
 * Convert zero indexed row and column pair to an Excel style A1:C5
 * range reference.
 */
void
lxlsx_rowcol_to_range(char *range,
                    lxlsx_row_t first_row, lxlsx_col_t first_col,
                    lxlsx_row_t last_row, lxlsx_col_t last_col)
{
    size_t pos;

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell(range, first_row, first_col);

    /* If the start and end cells are the same just return a single cell. */
    if (first_row == last_row && first_col == last_col)
        return;

    /* Get the end of the cell. */
    pos = strlen(range);

    /* Add the range separator. */
    range[pos++] = ':';

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell(&range[pos], last_row, last_col);
}

/*
 * Convert zero indexed row and column pairs to an Excel style $A$1:$C$5
 * range reference with absolute values.
 */
void
lxlsx_rowcol_to_range_abs(char *range,
                        lxlsx_row_t first_row, lxlsx_col_t first_col,
                        lxlsx_row_t last_row, lxlsx_col_t last_col)
{
    size_t pos;

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell_abs(range, first_row, first_col, 1, 1);

    /* If the start and end cells are the same just return a single cell. */
    if (first_row == last_row && first_col == last_col)
        return;

    /* Get the end of the cell. */
    pos = strlen(range);

    /* Add the range separator. */
    range[pos++] = ':';

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell_abs(&range[pos], last_row, last_col, 1, 1);
}

/*
 * Convert sheetname and zero indexed row and column pairs to an Excel style
 * Sheet1!$A$1:$C$5 formula reference with absolute values.
 */
void
lxlsx_rowcol_to_formula_abs(char *formula, const char *sheetname,
                          lxlsx_row_t first_row, lxlsx_col_t first_col,
                          lxlsx_row_t last_row, lxlsx_col_t last_col)
{
    size_t pos;
    char *quoted_name = lxlsx_quote_sheetname(sheetname);

    strncpy(formula, quoted_name, LXLSX_MAX_FORMULA_RANGE_LENGTH - 1);
    free(quoted_name);

    /* Get the end of the sheetname. */
    pos = strlen(formula);

    /* Add the range separator. */
    formula[pos++] = '!';

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell_abs(&formula[pos], first_row, first_col, 1, 1);

    /* If the start and end cells are the same just return a single cell. */
    if (first_row == last_row && first_col == last_col)
        return;

    /* Get the end of the cell. */
    pos = strlen(formula);

    /* Add the range separator. */
    formula[pos++] = ':';

    /* Add the first cell to the range. */
    lxlsx_rowcol_to_cell_abs(&formula[pos], last_row, last_col, 1, 1);
}

/*
 * Convert an Excel style A1 cell reference to a zero indexed row number.
 */
lxlsx_row_t
lxlsx_name_to_row(const char *row_str)
{
    lxlsx_row_t row_num = 0;

    if (!row_str)
        return row_num;

    /* Skip the column letters and absolute symbol of the A1 cell. */
    while (*row_str && !isdigit((unsigned char) *row_str))
        row_str++;

    /* Convert the row part of the A1 cell to a number. */
    if (*row_str)
        row_num = atoi(row_str);

    if (row_num)
        row_num--;

    return row_num;
}

/*
 * Convert the second row of an Excel range ref to a zero indexed number.
 */
uint32_t
lxlsx_name_to_row_2(const char *row_str)
{
    if (!row_str)
        return 0;

    /* Find the : separator in the range. */
    while (*row_str && *row_str != ':')
        row_str++;

    if (*row_str)
        return lxlsx_name_to_row(++row_str);
    else
        return 0;
}

/*
 * Convert an Excel style A1 cell reference to a zero indexed column number.
 */
lxlsx_col_t
lxlsx_name_to_col(const char *col_str)
{
    lxlsx_col_t col_num = 0;

    if (!col_str)
        return col_num;

    /* Convert leading column letters of A1 cell. Ignore absolute $ marker. */
    while (*col_str && (isupper((unsigned char) *col_str) || *col_str == '$')) {
        if (*col_str != '$')
            col_num = (col_num * 26) + (*col_str - 'A' + 1);
        col_str++;
    }

    if (col_num)
        col_num--;

    return col_num;
}

/*
 * Convert the second column of an Excel range ref to a zero indexed number.
 */
uint16_t
lxlsx_name_to_col_2(const char *col_str)
{
    if (!col_str)
        return 0;

    /* Find the : separator in the range. */
    while (*col_str && *col_str != ':')
        col_str++;

    if (*col_str)
        return lxlsx_name_to_col(++col_str);
    else
        return 0;
}

lxlsx_error
lxlsx_datetime_validate(lxlsx_datetime *datetime)
{
    if (!datetime)
        return LXLSX_ERROR_DATETIME_VALIDATION;

    /*
     * Excel uses the year 1900 as the default epoch but it uses 1899-12-31 as
     * the 0 date and internally we use the 0-0-0 date for time only values.
     */
    if (datetime->year < 1900 &&
        !(datetime->year == 0 &&
          datetime->month == 0 && datetime->day == 0) &&
        !(datetime->year == 1899 &&
          datetime->month == 12 && datetime->day == 31)) {

        LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid year: %d. "
                         "Valid range is 1900-9999.", datetime->year);

        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    if (datetime->year > 9999) {
        LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid year: %d. "
                         "Valid range is 1900-9999.", datetime->year);
        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    if (datetime->year != 0) {
        if (datetime->month < 1 || datetime->month > 12) {
            LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid month: %d. "
                             "Valid range is 1-12.", datetime->month);
            return LXLSX_ERROR_DATETIME_VALIDATION;
        }

        if (datetime->day < 1 || datetime->day > 31) {
            LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid day: %d. "
                             "Valid range is 1-31.", datetime->day);
            return LXLSX_ERROR_DATETIME_VALIDATION;
        }
    }

    if (datetime->hour < 0 || datetime->hour > 23) {
        LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid hour: %d. "
                         "Valid range is 0-23.", datetime->hour);
        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    if (datetime->min < 0 || datetime->min > 59) {
        LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid minute: %d. "
                         "Valid range is 0-59.", datetime->min);
        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    if (datetime->sec < 0.0 || datetime->sec >= 60.0) {
        LXLSX_WARN_FORMAT1("lxlsx_datetime_validate(): invalid seconds: %.3f. "
                         "Valid range is 0.0-59.999.", datetime->sec);
        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    return LXLSX_NO_ERROR;
}

/*
 * Convert a lxlsx_datetime struct to an Excel serial date, with a 1900
 * or 1904 epoch.
 */
double
lxlsx_datetime_to_excel_date_with_epoch(lxlsx_datetime *datetime,
                                      uint8_t use_1904_epoch)
{
    int year = datetime->year;
    int month = datetime->month;
    int day = datetime->day;
    int hour = datetime->hour;
    int min = datetime->min;
    double sec = datetime->sec;
    double seconds;
    int epoch = use_1904_epoch ? 1904 : 1900;
    int offset = use_1904_epoch ? 4 : 0;
    int norm = 300;
    int range;
    /* Set month days and check for leap year. */
    int mdays[] = { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    int leap = 0;
    int days = 0;
    int i;

    if (lxlsx_datetime_validate(datetime) != LXLSX_NO_ERROR)
        return 0.0;

    /* For times without dates set the default date for the epoch. */
    if (!year) {
        if (!use_1904_epoch) {
            year = 1899;
            month = 12;
            day = 31;
        }
        else {
            year = 1904;
            month = 1;
            day = 1;
        }
    }

    /* Convert the Excel seconds to a fraction of the seconds in 24 hours. */
    seconds = (hour * 60 * 60 + min * 60 + sec) / (24 * 60 * 60.0);

    /* Special cases for Excel dates in the 1900 epoch. */
    if (!use_1904_epoch) {
        /* Excel 1900 epoch. */
        if (year == 1899 && month == 12 && day == 31)
            return seconds;

        /* Excel 1900 epoch. */
        if (year == 1900 && month == 1 && day == 0)
            return seconds;

        /* Excel false leapday */
        if (year == 1900 && month == 2 && day == 29)
            return 60 + seconds;
    }

    /* We calculate the date by calculating the number of days since the */
    /* epoch and adjust for the number of leap days. We calculate the */
    /* number of leap days by normalizing the year in relation to the */
    /* epoch. Thus the year 2000 becomes 100 for 4-year and 100-year */
    /* leapdays and 400 for 400-year leapdays. */
    range = year - epoch;

    if (year % 4 == 0 && (year % 100 > 0 || year % 400 == 0)) {
        leap = 1;
        mdays[2] = 29;
    }

    /*
     * Calculate the serial date by accumulating the number of days
     * since the epoch.
     */

    /* Add days for previous months. */
    for (i = 0; i < month; i++) {
        days += mdays[i];
    }
    /* Add days for current month. */
    days += day;
    /* Add days for all previous years.  */
    days += range * 365;
    /* Add 4 year leapdays. */
    days += range / 4;
    /* Remove 100 year leapdays. */
    days -= (range + offset) / 100;
    /* Add 400 year leapdays. */
    days += (range + offset + norm) / 400;
    /* Remove leap days already counted. */
    days -= leap;

    /* Adjust for Excel erroneously treating 1900 as a leap year. */
    if (!use_1904_epoch && days > 59)
        days++;

    return days + seconds;
}

/*
 * Convert a lxlsx_datetime struct to an Excel serial date, for the 1900 epoch.
 */
double
lxlsx_datetime_to_excel_datetime(lxlsx_datetime *datetime)
{
    return lxlsx_datetime_to_excel_date_with_epoch(datetime, LXLSX_FALSE);
}

/*
 * Convert a unix datetime (1970/01/01 epoch) to an Excel serial date, with a
 * 1900 epoch.
 */
double
lxlsx_unixtime_to_excel_date(int64_t unixtime)
{
    return lxlsx_unixtime_to_excel_date_with_epoch(unixtime, LXLSX_FALSE);
}

/*
 * Convert a unix datetime (1970/01/01 epoch) to an Excel serial date, with a
 * 1900 or 1904 epoch.
 */
double
lxlsx_unixtime_to_excel_date_with_epoch(int64_t unixtime,
                                      uint8_t use_1904_epoch)
{
    double excel_datetime = 0.0;
    double epoch = use_1904_epoch ? 24107.0 : 25568.0;

    excel_datetime = epoch + (unixtime / (24 * 60 * 60.0));

    if (!use_1904_epoch && excel_datetime >= 60.0)
        excel_datetime = excel_datetime + 1.0;

    return excel_datetime;
}

/* Simple strdup() implementation since it isn't ANSI C. */
char *
lxlsx_strdup(const char *str)
{
    size_t len;
    char *copy;

    if (!str)
        return NULL;

    len = strlen(str) + 1;
    copy = malloc(len);

    if (copy)
        memcpy(copy, str, len);

    return copy;
}

/* Simple function to strdup() a formula string without the leading "=". */
char *
lxlsx_strdup_formula(const char *formula)
{
    if (!formula)
        return NULL;

    if (formula[0] == '=')
        return lxlsx_strdup(formula + 1);
    else
        return lxlsx_strdup(formula);
}

/* Simple strlen that counts UTF-8 characters. Assumes well formed UTF-8. */
size_t
lxlsx_utf8_strlen(const char *str)
{
    size_t byte_count = 0;
    size_t char_count = 0;

    while (str[byte_count]) {
        if ((str[byte_count] & 0xc0) != 0x80)
            char_count++;

        byte_count++;
    }

    return char_count;
}

/* Simple tolower() for strings. */
void
lxlsx_str_tolower(char *str)
{
    int i;

    for (i = 0; str[i]; i++)
        str[i] = tolower(str[i]);
}

/* Simple check for empty strings. */
uint8_t
lxlsx_str_is_empty(const char *str)
{
    if (str[0] == '\0')
        return 1;
    else
        return 0;
}

/* Create a quoted version of the worksheet name, or return an unmodified
 * copy if it doesn't required quoting. */
char *
lxlsx_quote_sheetname(const char *str)
{

    uint8_t needs_quoting = 0;
    size_t number_of_quotes = 2;
    size_t i, j;
    size_t len = strlen(str);

    /* Don't quote the sheetname if it is already quoted. */
    if (str[0] == '\'')
        return lxlsx_strdup(str);

    /* Check if the sheetname contains any characters that require it
     * to be quoted. Also check for single quotes within the string. */
    for (i = 0; i < len; i++) {
        if (!isalnum((unsigned char) str[i]) && str[i] != '_'
            && str[i] != '.')
            needs_quoting = 1;

        if (str[i] == '\'') {
            needs_quoting = 1;
            number_of_quotes++;
        }
    }

    if (!needs_quoting) {
        return lxlsx_strdup(str);
    }
    else {
        /* Add single quotes to the start and end of the string. */
        char *quoted_name = calloc(1, len + number_of_quotes + 1);
        RETURN_ON_MEM_ERROR(quoted_name, NULL);

        quoted_name[0] = '\'';

        for (i = 0, j = 1; i < len; i++, j++) {
            quoted_name[j] = str[i];

            /* Double quote inline single quotes. */
            if (str[i] == '\'') {
                quoted_name[++j] = '\'';
            }
        }
        quoted_name[j++] = '\'';
        quoted_name[j++] = '\0';

        return quoted_name;
    }
}

/*
 * Thin wrapper for tmpfile() so it can be over-ridden with a user defined
 * version if required for safety or portability.
 */
FILE *
lxlsx_tmpfile(const char *tmpdir)
{
#ifndef USE_STANDARD_TMPFILE
    return tmpfileplus(tmpdir, NULL, NULL, 0);
#else
    (void) tmpdir;
    return tmpfile();
#endif
}

/**
 * Return a memory-backed file if supported, otherwise a temporary one
 */
FILE *
lxlsx_get_filehandle(char **buf, size_t *size, const char *tmpdir)
{
    static size_t s;
#ifndef USE_FMEMOPEN
    FILE *file;
#endif
    if (!size)
        size = &s;
    *buf = NULL;
    *size = 0;
#ifdef USE_FMEMOPEN
    (void) tmpdir;
    return open_memstream(buf, size);
#else
    file = lxlsx_tmpfile(tmpdir);
    if (file)
        (void)setvbuf(file, NULL, _IOFBF, LXLSX_TMPFILE_BUFFER_SIZE);
    return file;
#endif
}

/*
 * Use third party function to handle sprintf of doubles for locale portable
 * code.
 */
#ifdef USE_DTOA_LIBRARY
int
lxlsx_sprintf_dbl(char *data, double number)
{
    emyg_dtoa(number, data);
    return 0;
}
#endif

/*
 * Retrieve runtime library version.
 */
const char *
lxlsx_version(void)
{
    return LXLSX_VERSION;
}

/*
 * Retrieve runtime library version ID.
 */
uint16_t
lxlsx_version_id(void)
{
    return LXLSX_VERSION_ID;
}

/*
 * Hash a worksheet password. Based on the algorithm in ECMA-376-4:2016,
 * Office Open XML File Formats - Transitional Migration Features,
 * Additional attributes for workbookProtection element (Part 1, §18.2.29).
 */
uint16_t
lxlsx_hash_password(const char *password)
{
    uint16_t byte_count = (uint16_t) strlen(password);
    uint16_t hash = 0;
    const char *p = &password[byte_count];

    if (!byte_count)
        return hash;

    while (p-- != password) {
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= *p & 0xFF;
    }

    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= byte_count;
    hash ^= 0xCE4B;

    return hash;
}

/* Make a simple portable version of fopen() for Windows. */
#ifdef __MINGW32__
#undef _WIN32
#endif

#ifdef _WIN32

#include <windows.h>

FILE *
lxlsx_fopen(const char *filename, const char *mode)
{
    int n;
    wchar_t wide_filename[_MAX_PATH + 1] = L"";
    wchar_t wide_mode[_MAX_PATH + 1] = L"";

    n = MultiByteToWideChar(CP_UTF8, 0, filename, (int) strlen(filename),
                            wide_filename, _MAX_PATH);

    if (n == 0) {
        LXLSX_ERROR("MultiByteToWideChar error: filename");
        return NULL;
    }

    n = MultiByteToWideChar(CP_UTF8, 0, mode, (int) strlen(mode),
                            wide_mode, _MAX_PATH);

    if (n == 0) {
        LXLSX_ERROR("MultiByteToWideChar error: mode");
        return NULL;
    }

    return _wfopen(wide_filename, wide_mode);
}
#else
FILE *
lxlsx_fopen(const char *filename, const char *mode)
{
    return fopen(filename, mode);
}
#endif
