#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "lxr_styles_priv.h"
#include "lxr_numfmt.h"
#include "lxr_zip.h"

void setUp(void) {}
void tearDown(void) {}

static const char *XLSX = LXR_TEST_HIDDEN_ROW_XLSX;

static void test_styles_load(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_styles *st = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_styles_load(z, "xl/styles.xml", &st));
    TEST_ASSERT_NOT_NULL(st);
    /* hidden_row.xlsx has multiple cellXfs */
    TEST_ASSERT_GREATER_THAN(0, lxr_styles_count(st));
    TEST_ASSERT_NOT_NULL(lxr_styles_get_xf(st, 0));
    TEST_ASSERT_NULL(lxr_styles_get_xf(st, 99999));
    lxr_styles_free(st);
    lxr_zip_close(z);
}

static void test_styles_no_path(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_styles *st = NULL;
    /* NULL path returns an empty styles handle, not an error. */
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_styles_load(z, NULL, &st));
    TEST_ASSERT_NOT_NULL(st);
    TEST_ASSERT_EQUAL_INT(0, lxr_styles_count(st));
    lxr_styles_free(st);
    lxr_zip_close(z);
}

/* numfmt classification spot checks ---------------------------------------- */

static void test_numfmt_builtin_categories(void)
{
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_GENERAL,  lxr_numfmt_builtin_category(0));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_DATE,     lxr_numfmt_builtin_category(14));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_DATE,     lxr_numfmt_builtin_category(15));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_DATETIME, lxr_numfmt_builtin_category(22));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_TIME,     lxr_numfmt_builtin_category(20));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_PERCENT,  lxr_numfmt_builtin_category(9));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_TEXT,     lxr_numfmt_builtin_category(49));
}

static void test_numfmt_classify_custom_strings(void)
{
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_DATE,     lxr_numfmt_classify("yyyy-mm-dd"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_DATETIME, lxr_numfmt_classify("yyyy-mm-dd hh:mm:ss"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_TIME,     lxr_numfmt_classify("hh:mm:ss"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_PERCENT,  lxr_numfmt_classify("0.00%"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_TEXT,     lxr_numfmt_classify("@"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_NUMBER,   lxr_numfmt_classify("#,##0.00"));
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_GENERAL,  lxr_numfmt_classify("General"));
}

static void test_numfmt_classify_ignores_quoted(void)
{
    /* "yyyy" is inside quotes, so it should NOT trigger DATE. */
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_NUMBER,
        lxr_numfmt_classify("\"yyyy\"0.00"));
    /* [Red] colour spec should not interfere. */
    TEST_ASSERT_EQUAL_INT(LXR_FMT_CATEGORY_NUMBER,
        lxr_numfmt_classify("[Red]#,##0.00"));
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
