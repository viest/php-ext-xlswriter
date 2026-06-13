#include <unity.h>

#include "xlsxreader.h"

void setUp(void) {}
void tearDown(void) {}

static void test_strerror_returns_message_for_known_codes(void)
{
    TEST_ASSERT_NOT_NULL(lxr_strerror(LXR_NO_ERROR));
    TEST_ASSERT_NOT_NULL(lxr_strerror(LXR_ERROR_UNSUPPORTED_FEATURE));
    TEST_ASSERT_NOT_NULL(lxr_strerror(LXR_ERROR_FILE_OPEN_FAILED));
}

static void test_open_missing_file_returns_open_failed(void)
{
    lxr_workbook *wb = NULL;
    lxr_error err = lxr_workbook_open("/nonexistent/none.xlsx", &wb);
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_FILE_OPEN_FAILED, err);
    TEST_ASSERT_NULL(wb);
}

static void test_open_null_args_returns_null_parameter(void)
{
    lxr_workbook *wb = NULL;
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_NULL_PARAMETER, lxr_workbook_open(NULL, &wb));
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_NULL_PARAMETER, lxr_workbook_open("foo.xlsx", NULL));
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_strerror_returns_message_for_known_codes);
    RUN_TEST(test_open_missing_file_returns_open_failed);
    RUN_TEST(test_open_null_args_returns_null_parameter);
    return UNITY_END();
}
