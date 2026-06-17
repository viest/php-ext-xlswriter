#include <stdio.h>
#include <stdlib.h>
#include "../../include/lxlsx/utility.h"

/* Compare expected results with the XML data written to the output
 * test file.
 */
#define RUN_XLSX_STREQ(exp, got)                                    \
    fflush(testfile);                                               \
    int file_size = ftell(testfile);                                \
                                                                    \
    got = (char*)calloc(file_size + 1, 1);                          \
                                                                    \
    rewind(testfile);                                               \
    (void)fread(got, file_size, 1, testfile);                       \
                                                                    \
    ASSERT_STR((exp), (got));                                       \
                                                                    \
    if (got)                                                        \
        free(got);                                                  \
                                                                    \
    fclose(testfile)

/* Compare expected results with the XML data written to the output
 * test file. Same as the previous macro but only shows the difference
 * from where it starts. Suitable for long strings of XML data.
 */
#define RUN_XLSX_STREQ_SHORT(exp, got)                              \
    fflush(testfile);                                               \
    int file_size = ftell(testfile);                                \
                                                                    \
    got = (char*)calloc(file_size + 1, 1);                          \
                                                                    \
    rewind(testfile);                                               \
    (void)fread(got, file_size, 1, testfile);                       \
                                                                    \
    /* Start comparison from first difference. */                   \
    char *got_short = got;                                          \
    char *exp_short = exp;                                          \
    while (*exp_short && *exp_short == *got_short) {                \
        exp_short++;                                                \
        got_short++;                                                \
    }                                                               \
                                                                    \
    ASSERT_STR(exp_short, got_short);                               \
                                                                    \
    if (got)                                                        \
        free(got);                                                  \
                                                                    \
    fclose(testfile)


#define TEST_COL_TO_NAME(num, abs, exp)                             \
    lxw_col_to_name(got, num, abs);                                 \
    ASSERT_STR(exp, got);


#define TEST_ROWCOL_TO_CELL(row, col, exp)                          \
    lxw_rowcol_to_cell(got, row, col);                              \
    ASSERT_STR(exp, got);


#define TEST_ROWCOL_TO_CELL_ABS(row, col, row_abs, col_abs, exp)    \
    lxw_rowcol_to_cell_abs(got, row, col, row_abs, col_abs);        \
    ASSERT_STR(exp, got);


#define TEST_ROWCOL_TO_RANGE(row1, col1, row2, col2, exp)           \
    lxw_rowcol_to_range(got, row1, col1, row2, col2);               \
    ASSERT_STR(exp, got);


#define TEST_ROWCOL_TO_RANGE_ABS(row1, col1, row2, col2, exp)       \
    lxw_rowcol_to_range_abs(got, row1, col1, row2, col2);           \
    ASSERT_STR(exp, got);


#define TEST_ROWCOL_TO_FORMULA_ABS(sheet, row1, col1, row2, col2, exp) \
    lxw_rowcol_to_formula_abs(got, sheet, row1, col1, row2, col2);     \
    ASSERT_STR(exp, got);


#define TEST_DATETIME_TIME(_hour, _min, _sec, exp)                  \
    datetime = (lxw_datetime*)calloc(1, sizeof(lxw_datetime));      \
    datetime->hour  = _hour;                                        \
    datetime->min   = _min;                                         \
    datetime->sec   = _sec;                                         \
                                                                    \
    got = lxw_datetime_to_excel_datetime(datetime);                 \
                                                                    \
    ASSERT_DBL_NEAR(exp, got);                                      \
    free(datetime);

#define TEST_DATETIME_DATE(_year, _month, _day, exp)                \
    datetime = (lxw_datetime*)calloc(1, sizeof(lxw_datetime));      \
    datetime->year  = _year;                                        \
    datetime->month = _month;                                       \
    datetime->day   = _day;                                         \
                                                                    \
    got = lxw_datetime_to_excel_datetime(datetime);                 \
                                                                    \
    ASSERT_DBL_NEAR(exp, got);                                      \
    free(datetime);

#define TEST_DATETIME_DATE_1904(_year, _month, _day, exp)           \
    datetime = (lxw_datetime*)calloc(1, sizeof(lxw_datetime));      \
    datetime->year  = _year;                                        \
    datetime->month = _month;                                       \
    datetime->day   = _day;                                         \
                                                                    \
    got = lxw_datetime_to_excel_date_with_epoch(datetime, 1);            \
                                                                    \
    ASSERT_DBL_NEAR(exp, got);                                      \
    free(datetime);

#define TEST_DATETIME(_year, _month, _day, _hour, _min, _sec, exp)  \
    datetime = (lxw_datetime*)calloc(1, sizeof(lxw_datetime));      \
    datetime->year  = _year;                                        \
    datetime->month = _month;                                       \
    datetime->day   = _day;                                         \
    datetime->hour  = _hour;                                        \
    datetime->min   = _min;                                         \
    datetime->sec   = _sec;                                         \
                                                                    \
    got = lxw_datetime_to_excel_datetime(datetime);                 \
                                                                    \
    ASSERT_DBL_NEAR(exp, got);                                      \
    free(datetime);

#define TEST_UNIXTIME(_unixtime, exp)                               \
    got = lxw_unixtime_to_excel_date(_unixtime);                    \
    ASSERT_DBL_NEAR(exp, got);

#define TEST_UNIXTIME_1904(_unixtime, exp)                          \
    got = lxw_unixtime_to_excel_date_with_epoch(_unixtime, 1);           \
    ASSERT_DBL_NEAR(exp, got);

