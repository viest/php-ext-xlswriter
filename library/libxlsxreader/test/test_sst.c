#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "lxr_sst.h"
#include "lxr_zip.h"

void setUp(void) {}
void tearDown(void) {}

static const char *XLSX = LXR_TEST_HIDDEN_ROW_XLSX;

static void test_sst_open_full(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_sst *s = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_sst_open(z, "xl/sharedStrings.xml", LXR_SST_MODE_FULL, &s));
    TEST_ASSERT_NOT_NULL(s);
    TEST_ASSERT_GREATER_THAN(0, lxr_sst_count(s));
    /* Index 0 must always be a non-NULL string (Excel stores at least one). */
    TEST_ASSERT_NOT_NULL(lxr_sst_get(s, 0));
    lxr_sst_close(s);
    lxr_zip_close(z);
}

static void test_sst_full_vs_streaming_consistency(void)
{
    lxr_zip *z1 = lxr_zip_open_path(XLSX);
    lxr_zip *z2 = lxr_zip_open_path(XLSX);
    lxr_sst *full = NULL;
    lxr_sst *stream = NULL;
    size_t i, count_full;

    TEST_ASSERT_NOT_NULL(z1);
    TEST_ASSERT_NOT_NULL(z2);

    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_sst_open(z1, "xl/sharedStrings.xml", LXR_SST_MODE_FULL, &full));
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_sst_open(z2, "xl/sharedStrings.xml", LXR_SST_MODE_STREAMING, &stream));

    count_full = lxr_sst_count(full);
    TEST_ASSERT_GREATER_THAN(0, count_full);

    /* Pull every index in order; streaming should match FULL byte-for-byte. */
    for (i = 0; i < count_full; i++) {
        const char *a = lxr_sst_get(full, (uint32_t)i);
        const char *b = lxr_sst_get(stream, (uint32_t)i);
        TEST_ASSERT_NOT_NULL(a);
        TEST_ASSERT_NOT_NULL(b);
        TEST_ASSERT_EQUAL_STRING(a, b);
    }

    lxr_sst_close(full);
    lxr_sst_close(stream);
    lxr_zip_close(z1);
    lxr_zip_close(z2);
}

static void test_sst_streaming_lazy_load(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_sst *s = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_sst_open(z, "xl/sharedStrings.xml", LXR_SST_MODE_STREAMING, &s));
    /* Streaming: nothing loaded yet. */
    TEST_ASSERT_EQUAL_INT(0, lxr_sst_loaded_count(s));
    /* Touching index 0 must load at least one string. */
    TEST_ASSERT_NOT_NULL(lxr_sst_get(s, 0));
    TEST_ASSERT_GREATER_OR_EQUAL_INT(1, lxr_sst_loaded_count(s));
    lxr_sst_close(s);
    lxr_zip_close(z);
}

static void test_sst_out_of_range(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_sst *s = NULL;
    lxr_sst_open(z, "xl/sharedStrings.xml", LXR_SST_MODE_FULL, &s);
    TEST_ASSERT_NOT_NULL(s);
    TEST_ASSERT_NULL(lxr_sst_get(s, 99999));
    lxr_sst_close(s);
    lxr_zip_close(z);
}

static void test_sst_missing_entry(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_sst *s = NULL;
    lxr_error err = lxr_sst_open(z, "xl/notthere.xml", LXR_SST_MODE_FULL, &s);
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_ZIP_ENTRY_NOT_FOUND, err);
    TEST_ASSERT_NULL(s);
    lxr_zip_close(z);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_sst_open_full);
    RUN_TEST(test_sst_full_vs_streaming_consistency);
    RUN_TEST(test_sst_streaming_lazy_load);
    RUN_TEST(test_sst_out_of_range);
    RUN_TEST(test_sst_missing_entry);
    return UNITY_END();
}
