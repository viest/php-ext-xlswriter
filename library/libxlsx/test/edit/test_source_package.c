#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxlsx.h"
#include "lxlsx/source_package.h"

void setUp(void) {}
void tearDown(void) {}

static const char *SOURCE_XLSX = "fixtures/edit_source.xlsx";
static const char *NOOP_XLSX = "fixtures/edit_noop.xlsx";
static const char *RAW_COPY_XLSX = "fixtures/edit_raw_copy.xlsx";

static void assert_ok(lxlsx_error err)
{
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR, err);
}

static void write_source_workbook(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;

    remove(SOURCE_XLSX);
    workbook = lxlsx_workbook_new(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_add_worksheet(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);
    assert_ok(lxlsx_worksheet_write_number(worksheet, 0, 0, 1.0, NULL));
    assert_ok(lxlsx_worksheet_write_number(worksheet, 0, 1, 2.0, NULL));
    assert_ok(lxlsx_workbook_close(workbook));
}

static unsigned char *slurp(const char *path, size_t *len)
{
    FILE *fp = fopen(path, "rb");
    long size;
    unsigned char *buf;
    TEST_ASSERT_NOT_NULL(fp);
    fseek(fp, 0, SEEK_END);
    size = ftell(fp);
    fseek(fp, 0, SEEK_SET);
    buf = (unsigned char *)malloc((size_t)size);
    TEST_ASSERT_NOT_NULL(buf);
    TEST_ASSERT_EQUAL_INT((int)size, (int)fread(buf, 1, (size_t)size, fp));
    fclose(fp);
    *len = (size_t)size;
    return buf;
}

static void assert_files_equal(const char *left, const char *right)
{
    size_t left_len;
    size_t right_len;
    unsigned char *left_data = slurp(left, &left_len);
    unsigned char *right_data = slurp(right, &right_len);
    TEST_ASSERT_EQUAL_size_t(left_len, right_len);
    TEST_ASSERT_EQUAL_INT(0, memcmp(left_data, right_data, left_len));
    free(left_data);
    free(right_data);
}

static void test_noop_save_is_byte_identical(void)
{
    lxlsx_source_package *package = NULL;

    write_source_workbook();
    remove(NOOP_XLSX);

    assert_ok(lxlsx_source_package_open(SOURCE_XLSX, &package));
    TEST_ASSERT_GREATER_THAN(5, (int)lxlsx_source_package_entry_count(package));
    TEST_ASSERT_GREATER_OR_EQUAL_INT(
        0, lxlsx_source_package_find_first(package, "xl/worksheets/sheet1.xml"));
    assert_ok(lxlsx_source_package_save_copy(package, NOOP_XLSX));
    lxlsx_source_package_close(package);

    assert_files_equal(SOURCE_XLSX, NOOP_XLSX);
}

static void test_replacement_preserves_untouched_raw_local_record(void)
{
    lxlsx_source_package *source = NULL;
    lxlsx_source_package *saved = NULL;
    int sheet_index;
    int untouched_index;
    unsigned char *sheet_xml = NULL;
    size_t sheet_xml_len = 0;
    const unsigned char *source_raw = NULL;
    const unsigned char *saved_raw = NULL;
    size_t source_raw_len = 0;
    size_t saved_raw_len = 0;
    const char suffix[] = "\n<!-- edit replacement -->\n";
    unsigned char *replacement;
    lxlsx_source_package_replacement repl;

    write_source_workbook();
    remove(RAW_COPY_XLSX);

    assert_ok(lxlsx_source_package_open(SOURCE_XLSX, &source));
    sheet_index = lxlsx_source_package_find_first(source, "xl/worksheets/sheet1.xml");
    untouched_index = lxlsx_source_package_find_first(source, "[Content_Types].xml");
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, sheet_index);
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, untouched_index);

    assert_ok(lxlsx_source_package_read_entry(source, (size_t)sheet_index,
                                              &sheet_xml, &sheet_xml_len));
    replacement = (unsigned char *)malloc(sheet_xml_len + sizeof(suffix));
    TEST_ASSERT_NOT_NULL(replacement);
    memcpy(replacement, sheet_xml, sheet_xml_len);
    memcpy(replacement + sheet_xml_len, suffix, sizeof(suffix) - 1);

    repl.entry_index = (size_t)sheet_index;
    repl.data = replacement;
    repl.size = sheet_xml_len + sizeof(suffix) - 1;

    assert_ok(lxlsx_source_package_entry_raw_local_record(
        source, (size_t)untouched_index, &source_raw, &source_raw_len));
    assert_ok(lxlsx_source_package_save_with_replacements(
        source, RAW_COPY_XLSX, &repl, 1));

    assert_ok(lxlsx_source_package_open(RAW_COPY_XLSX, &saved));
    assert_ok(lxlsx_source_package_entry_raw_local_record(
        saved, (size_t)untouched_index, &saved_raw, &saved_raw_len));
    TEST_ASSERT_EQUAL_size_t(source_raw_len, saved_raw_len);
    TEST_ASSERT_EQUAL_INT(0, memcmp(source_raw, saved_raw, source_raw_len));

    lxlsx_source_package_free_buffer(sheet_xml);
    free(replacement);
    lxlsx_source_package_close(saved);
    lxlsx_source_package_close(source);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_noop_save_is_byte_identical);
    RUN_TEST(test_replacement_preserves_untouched_raw_local_record);
    return UNITY_END();
}
