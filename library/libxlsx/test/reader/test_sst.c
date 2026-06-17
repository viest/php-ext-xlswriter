#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxlsx_reader_test_paths.h"
#include "lxlsx_reader_sst.h"
#include "lxlsx_reader_zip.h"

void setUp(void) {}
void tearDown(void) {}

static const char *XLSX = LXLSX_READER_TEST_HIDDEN_ROW_XLSX;

static void test_sst_open_full(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_sst *s = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_sst_open(z, "xl/sharedStrings.xml", LXLSX_READER_SST_MODE_FULL, &s));
    TEST_ASSERT_NOT_NULL(s);
    TEST_ASSERT_GREATER_THAN(0, lxlsx_reader_sst_count(s));
    /* Index 0 must always be a non-NULL string (Excel stores at least one). */
    TEST_ASSERT_NOT_NULL(lxlsx_reader_sst_get(s, 0));
    lxlsx_reader_sst_close(s);
    lxlsx_reader_zip_close(z);
}

static void test_sst_full_vs_streaming_consistency(void)
{
    lxlsx_reader_zip *z1 = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_zip *z2 = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_sst *full = NULL;
    lxlsx_reader_sst *stream = NULL;
    size_t i, count_full;

    TEST_ASSERT_NOT_NULL(z1);
    TEST_ASSERT_NOT_NULL(z2);

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_sst_open(z1, "xl/sharedStrings.xml", LXLSX_READER_SST_MODE_FULL, &full));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_sst_open(z2, "xl/sharedStrings.xml", LXLSX_READER_SST_MODE_STREAMING, &stream));

    count_full = lxlsx_reader_sst_count(full);
    TEST_ASSERT_GREATER_THAN(0, count_full);

    /* Pull every index in order; streaming should match FULL byte-for-byte. */
    for (i = 0; i < count_full; i++) {
        const char *a = lxlsx_reader_sst_get(full, (uint32_t)i);
        const char *b = lxlsx_reader_sst_get(stream, (uint32_t)i);
        TEST_ASSERT_NOT_NULL(a);
        TEST_ASSERT_NOT_NULL(b);
        TEST_ASSERT_EQUAL_STRING(a, b);
    }

    lxlsx_reader_sst_close(full);
    lxlsx_reader_sst_close(stream);
    lxlsx_reader_zip_close(z1);
    lxlsx_reader_zip_close(z2);
}

static void test_sst_streaming_lazy_load(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_sst *s = NULL;
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_sst_open(z, "xl/sharedStrings.xml", LXLSX_READER_SST_MODE_STREAMING, &s));
    /* Streaming: nothing loaded yet. */
    TEST_ASSERT_EQUAL_INT(0, lxlsx_reader_sst_loaded_count(s));
    /* Touching index 0 must load at least one string. */
    TEST_ASSERT_NOT_NULL(lxlsx_reader_sst_get(s, 0));
    TEST_ASSERT_GREATER_OR_EQUAL_INT(1, lxlsx_reader_sst_loaded_count(s));
    lxlsx_reader_sst_close(s);
    lxlsx_reader_zip_close(z);
}

static void test_sst_out_of_range(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_sst *s = NULL;
    lxlsx_reader_sst_open(z, "xl/sharedStrings.xml", LXLSX_READER_SST_MODE_FULL, &s);
    TEST_ASSERT_NOT_NULL(s);
    TEST_ASSERT_NULL(lxlsx_reader_sst_get(s, 99999));
    lxlsx_reader_sst_close(s);
    lxlsx_reader_zip_close(z);
}

static void test_sst_missing_entry(void)
{
    lxlsx_reader_zip *z = lxlsx_reader_zip_open_path(XLSX);
    lxlsx_reader_sst *s = NULL;
    lxlsx_reader_error err = lxlsx_reader_sst_open(z, "xl/notthere.xml", LXLSX_READER_SST_MODE_FULL, &s);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND, err);
    TEST_ASSERT_NULL(s);
    lxlsx_reader_zip_close(z);
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
