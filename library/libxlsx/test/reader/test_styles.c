#include <string.h>

#include <unity.h>

#include "lxlsx_reader_test_paths.h"
#include "lxlsx_reader_styles_priv.h"
#include "lxlsx_reader_numfmt.h"
#include "lxlsx_reader_zip.h"

void setUp(void) {}
void tearDown(void) {}

static const char *XLSX = LXLSX_READER_TEST_HIDDEN_ROW_XLSX;

static void test_styles_load(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_styles *st = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_styles_load(z, "xl/styles.xml", &st));
    TEST_ASSERT_NOT_NULL(st);
    /* hidden_row.xlsx has multiple cellXfs */
    TEST_ASSERT_GREATER_THAN(0, lxlsx_reader_styles_count(st));
    TEST_ASSERT_NOT_NULL(lxlsx_reader_styles_get_xf(st, 0));
    TEST_ASSERT_NULL(lxlsx_reader_styles_get_xf(st, 99999));
    lxlsx_reader_styles_free(st);
    lxlsx_reader_zip_close(z);
}

static void test_styles_no_path(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_styles *st = NULL;
    /* NULL path returns an empty styles handle, not an error. */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_styles_load(z, NULL, &st));
    TEST_ASSERT_NOT_NULL(st);
    TEST_ASSERT_EQUAL_INT(0, lxlsx_reader_styles_count(st));
    lxlsx_reader_styles_free(st);
    lxlsx_reader_zip_close(z);
}

/* numfmt classification spot checks ---------------------------------------- */

static void test_numfmt_builtin_categories(void)
{
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_GENERAL,  lxlsx_reader_numfmt_builtin_category(0));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_DATE,     lxlsx_reader_numfmt_builtin_category(14));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_DATE,     lxlsx_reader_numfmt_builtin_category(15));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_DATETIME, lxlsx_reader_numfmt_builtin_category(22));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_TIME,     lxlsx_reader_numfmt_builtin_category(20));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_PERCENT,  lxlsx_reader_numfmt_builtin_category(9));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_TEXT,     lxlsx_reader_numfmt_builtin_category(49));
}

static void test_numfmt_classify_custom_strings(void)
{
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_DATE,     lxlsx_reader_numfmt_classify("yyyy-mm-dd"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_DATETIME, lxlsx_reader_numfmt_classify("yyyy-mm-dd hh:mm:ss"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_TIME,     lxlsx_reader_numfmt_classify("hh:mm:ss"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_PERCENT,  lxlsx_reader_numfmt_classify("0.00%"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_TEXT,     lxlsx_reader_numfmt_classify("@"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_NUMBER,   lxlsx_reader_numfmt_classify("#,##0.00"));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_GENERAL,  lxlsx_reader_numfmt_classify("General"));
}

static void test_numfmt_classify_ignores_quoted(void)
{
    /* "yyyy" is inside quotes, so it should NOT trigger DATE. */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_NUMBER,
        lxlsx_reader_numfmt_classify("\"yyyy\"0.00"));
    /* [Red] colour spec should not interfere. */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FMT_CATEGORY_NUMBER,
        lxlsx_reader_numfmt_classify("[Red]#,##0.00"));
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_styles_load);
    RUN_TEST(test_styles_no_path);
    RUN_TEST(test_numfmt_builtin_categories);
    RUN_TEST(test_numfmt_classify_custom_strings);
    RUN_TEST(test_numfmt_classify_ignores_quoted);
    return UNITY_END();
}
