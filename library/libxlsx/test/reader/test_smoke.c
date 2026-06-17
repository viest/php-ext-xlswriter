#include <unity.h>

#include "libxlsx.h"

void setUp(void) {}
void tearDown(void) {}

static void test_strerror_returns_message_for_known_codes(void)
{
    TEST_ASSERT_NOT_NULL(lxlsx_reader_strerror(LXLSX_READER_NO_ERROR));
    TEST_ASSERT_NOT_NULL(lxlsx_reader_strerror(LXLSX_READER_ERROR_UNSUPPORTED_FEATURE));
    TEST_ASSERT_NOT_NULL(lxlsx_reader_strerror(LXLSX_READER_ERROR_FILE_OPEN_FAILED));
}

static void test_open_missing_file_returns_open_failed(void)
{
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_error err = lxlsx_reader_workbook_open("/nonexistent/none.xlsx", &wb);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_FILE_OPEN_FAILED, err);
    TEST_ASSERT_NULL(wb);
}

static void test_open_null_args_returns_null_parameter(void)
{
    lxlsx_reader_workbook *wb = NULL;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_NULL_PARAMETER, lxlsx_reader_workbook_open(NULL, &wb));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_NULL_PARAMETER, lxlsx_reader_workbook_open("foo.xlsx", NULL));
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_strerror_returns_message_for_known_codes);
    RUN_TEST(test_open_missing_file_returns_open_failed);
    RUN_TEST(test_open_null_args_returns_null_parameter);
    return UNITY_END();
}
