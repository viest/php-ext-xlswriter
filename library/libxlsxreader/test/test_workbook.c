#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "xlsxreader.h"

void setUp(void) {}
void tearDown(void) {}

static void test_open_path(void)
{
    lxr_workbook *wb = NULL;
    lxr_error err = lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, err);
    TEST_ASSERT_NOT_NULL(wb);
    lxr_workbook_close(wb);
}

static void test_open_returns_error_on_missing_file(void)
{
    lxr_workbook *wb = NULL;
    lxr_error err = lxr_workbook_open("/nonexistent/none.xlsx", &wb);
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_FILE_OPEN_FAILED, err);
    TEST_ASSERT_NULL(wb);
}

static void test_sheet_count_and_names(void)
{
    lxr_workbook *wb = NULL;
    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);

    TEST_ASSERT_GREATER_OR_EQUAL_INT(1, lxr_workbook_sheet_count(wb));
    TEST_ASSERT_NOT_NULL(lxr_workbook_sheet_name(wb, 0));
    TEST_ASSERT_NULL(lxr_workbook_sheet_name(wb, 999));

    lxr_workbook_close(wb);
}

static void test_uses_1904_dates(void)
{
    lxr_workbook *wb = NULL;
    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);
    /* hidden_row.xlsx is a standard 1900-based workbook. */
    TEST_ASSERT_EQUAL_INT(0, lxr_workbook_uses_1904_dates(wb));
    lxr_workbook_close(wb);
}

static void test_styles_present(void)
{
    lxr_workbook *wb = NULL;
    const lxr_styles *st;
    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    TEST_ASSERT_NOT_NULL(wb);
    st = lxr_workbook_get_styles(wb);
    TEST_ASSERT_NOT_NULL(st);
    /* xfs_count > 0 means cellXfs was parsed; checked via lookup of index 0. */
    TEST_ASSERT_NOT_NULL(lxr_styles_get_xf(st, 0));
    lxr_workbook_close(wb);
}

static void test_open_ex_streaming_mode(void)
{
    lxr_workbook *wb = NULL;
    lxr_open_options opts = { .sst_mode = LXR_SST_MODE_STREAMING };
    lxr_error err = lxr_workbook_open_ex(LXR_TEST_HIDDEN_ROW_XLSX, &opts, &wb);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, err);
    TEST_ASSERT_NOT_NULL(wb);
    lxr_workbook_close(wb);
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
