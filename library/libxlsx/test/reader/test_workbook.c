#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxlsx_reader_test_paths.h"
#include "lxlsx/reader.h"

void setUp(void) {}
void tearDown(void) {}

static void test_open_path(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_error err = lxlsx_reader_workbook_open(LXLSX_READER_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, err);
    TEST_ASSERT_NOT_NULL(wb);
    lxlsx_reader_workbook_close(wb);
}

static void test_open_returns_error_on_missing_file(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_error err = lxlsx_reader_workbook_open("/nonexistent/none.xlsx", &wb);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_FILE_OPEN_FAILED, err);
    TEST_ASSERT_NULL(wb);
}

static void test_sheet_count_and_names(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_workbook_open(LXLSX_READER_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);

    TEST_ASSERT_GREATER_OR_EQUAL_INT(1, lxlsx_reader_workbook_sheet_count(wb));
    TEST_ASSERT_NOT_NULL(lxlsx_reader_workbook_sheet_name(wb, 0));
    TEST_ASSERT_NULL(lxlsx_reader_workbook_sheet_name(wb, 999));

    lxlsx_reader_workbook_close(wb);
}

static void test_uses_1904_dates(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_workbook_open(LXLSX_READER_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);
    /* hidden_row.xlsx is a standard 1900-based workbook. */
    TEST_ASSERT_EQUAL_INT(0, lxlsx_reader_workbook_uses_1904_dates(wb));
    lxlsx_reader_workbook_close(wb);
}

static void test_styles_present(void)
{
    lxlsx_reader_workbook *wb = NULL;
    const lxlsx_reader_styles *st;
    lxlsx_reader_workbook_open(LXLSX_READER_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);
    st = lxlsx_reader_workbook_get_styles(wb);
    TEST_ASSERT_NOT_NULL(st);
    /* xfs_count > 0 means cellXfs was parsed; checked via lookup of index 0. */
    TEST_ASSERT_NOT_NULL(lxlsx_reader_styles_get_xf(st, 0));
    lxlsx_reader_workbook_close(wb);
}

static void test_open_ex_streaming_mode(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_open_options opts = { .sst_mode = LXLSX_READER_SST_MODE_STREAMING };
    lxlsx_reader_error err = lxlsx_reader_workbook_open_ex(LXLSX_READER_TEST_HIDDEN_ROW_XLSX, &opts, &wb);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, err);
    TEST_ASSERT_NOT_NULL(wb);
    lxlsx_reader_workbook_close(wb);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_open_path);
    RUN_TEST(test_open_returns_error_on_missing_file);
    RUN_TEST(test_sheet_count_and_names);
    RUN_TEST(test_uses_1904_dates);
    RUN_TEST(test_styles_present);
    RUN_TEST(test_open_ex_streaming_mode);
    return UNITY_END();
}
