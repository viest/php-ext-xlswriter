#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>
#include <zlib.h>

#include "libxlsx.h"
#include "libxlsx/source_package.h"

void setUp(void) {}
void tearDown(void) {}

static const char *SOURCE_XLSX = "fixtures/edit_source.xlsx";
static const char *NOOP_XLSX = "fixtures/edit_noop.xlsx";
static const char *RAW_COPY_XLSX = "fixtures/edit_raw_copy.xlsx";
static const char *DUP_ZIP = "fixtures/edit_duplicate.zip";
static const char *DUP_NOOP_ZIP = "fixtures/edit_duplicate_noop.zip";

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

static void put_le16_file(FILE *fp, unsigned int value)
{
    unsigned char buf[2];
    buf[0] = (unsigned char)(value & 0xff);
    buf[1] = (unsigned char)((value >> 8) & 0xff);
    TEST_ASSERT_EQUAL_INT(2, (int)fwrite(buf, 1, sizeof(buf), fp));
}

static void put_le32_file(FILE *fp, unsigned long value)
{
    unsigned char buf[4];
    buf[0] = (unsigned char)(value & 0xff);
    buf[1] = (unsigned char)((value >> 8) & 0xff);
    buf[2] = (unsigned char)((value >> 16) & 0xff);
    buf[3] = (unsigned char)((value >> 24) & 0xff);
    TEST_ASSERT_EQUAL_INT(4, (int)fwrite(buf, 1, sizeof(buf), fp));
}

static void write_store_local(FILE *fp, const char *name, const char *payload,
                              unsigned long *offset, unsigned long *crc)
{
    size_t name_len = strlen(name);
    size_t payload_len = strlen(payload);
    *offset = (unsigned long)ftell(fp);
    *crc = crc32(0L, Z_NULL, 0);
    *crc = crc32(*crc, (const Bytef *)payload, (uInt)payload_len);

    put_le32_file(fp, 0x04034b50UL);
    put_le16_file(fp, 20);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le32_file(fp, *crc);
    put_le32_file(fp, (unsigned long)payload_len);
    put_le32_file(fp, (unsigned long)payload_len);
    put_le16_file(fp, (unsigned int)name_len);
    put_le16_file(fp, 0);
    TEST_ASSERT_EQUAL_INT((int)name_len, (int)fwrite(name, 1, name_len, fp));
    TEST_ASSERT_EQUAL_INT((int)payload_len, (int)fwrite(payload, 1, payload_len, fp));
}

static void write_store_central(FILE *fp, const char *name, const char *payload,
                                unsigned long offset, unsigned long crc)
{
    size_t name_len = strlen(name);
    size_t payload_len = strlen(payload);

    put_le32_file(fp, 0x02014b50UL);
    put_le16_file(fp, 20);
    put_le16_file(fp, 20);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le32_file(fp, crc);
    put_le32_file(fp, (unsigned long)payload_len);
    put_le32_file(fp, (unsigned long)payload_len);
    put_le16_file(fp, (unsigned int)name_len);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le32_file(fp, 0);
    put_le32_file(fp, offset);
    TEST_ASSERT_EQUAL_INT((int)name_len, (int)fwrite(name, 1, name_len, fp));
}

static void write_duplicate_zip(void)
{
    FILE *fp;
    unsigned long off1, off2, crc1, crc2;
    unsigned long central_offset;
    unsigned long central_end;

    remove(DUP_ZIP);
    fp = fopen(DUP_ZIP, "wb");
    TEST_ASSERT_NOT_NULL(fp);

    write_store_local(fp, "dup.txt", "first", &off1, &crc1);
    write_store_local(fp, "dup.txt", "second", &off2, &crc2);

    central_offset = (unsigned long)ftell(fp);
    write_store_central(fp, "dup.txt", "first", off1, crc1);
    write_store_central(fp, "dup.txt", "second", off2, crc2);
    central_end = (unsigned long)ftell(fp);

    put_le32_file(fp, 0x06054b50UL);
    put_le16_file(fp, 0);
    put_le16_file(fp, 0);
    put_le16_file(fp, 2);
    put_le16_file(fp, 2);
    put_le32_file(fp, central_end - central_offset);
    put_le32_file(fp, central_offset);
    put_le16_file(fp, 0);

    fclose(fp);
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

static void test_duplicate_entries_noop_save_preserves_bytes(void)
{
    lxlsx_source_package *package = NULL;

    write_duplicate_zip();
    remove(DUP_NOOP_ZIP);

    assert_ok(lxlsx_source_package_open(DUP_ZIP, &package));
    TEST_ASSERT_EQUAL_size_t(2, lxlsx_source_package_entry_count(package));
    assert_ok(lxlsx_source_package_save_copy(package, DUP_NOOP_ZIP));
    lxlsx_source_package_close(package);

    assert_files_equal(DUP_ZIP, DUP_NOOP_ZIP);
}

static void test_zip64_entries_are_explicitly_unsupported(void)
{
    unsigned char data[46 + 5 + 4 + 22];
    size_t pos = 0;
    lxlsx_source_package *package = NULL;

#define PUT16(v) do { data[pos++] = (unsigned char)((v) & 0xff); data[pos++] = (unsigned char)(((v) >> 8) & 0xff); } while (0)
#define PUT32(v) do { unsigned long x = (unsigned long)(v); data[pos++] = (unsigned char)(x & 0xff); data[pos++] = (unsigned char)((x >> 8) & 0xff); data[pos++] = (unsigned char)((x >> 16) & 0xff); data[pos++] = (unsigned char)((x >> 24) & 0xff); } while (0)

    PUT32(0x02014b50UL);
    PUT16(45);
    PUT16(45);
    PUT16(0);
    PUT16(0);
    PUT16(0);
    PUT16(0);
    PUT32(0);
    PUT32(0xffffffffUL);
    PUT32(0xffffffffUL);
    PUT16(5);
    PUT16(4);
    PUT16(0);
    PUT16(0);
    PUT16(0);
    PUT32(0);
    PUT32(0xffffffffUL);
    memcpy(data + pos, "a.txt", 5);
    pos += 5;
    PUT16(0x0001);
    PUT16(0);

    PUT32(0x06054b50UL);
    PUT16(0);
    PUT16(0);
    PUT16(1);
    PUT16(1);
    PUT32(46 + 5 + 4);
    PUT32(0);
    PUT16(0);

#undef PUT16
#undef PUT32

    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_source_package_open_memory(data, pos, &package));
    TEST_ASSERT_NULL(package);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_noop_save_is_byte_identical);
    RUN_TEST(test_replacement_preserves_untouched_raw_local_record);
    RUN_TEST(test_duplicate_entries_noop_save_preserves_bytes);
    RUN_TEST(test_zip64_entries_are_explicitly_unsupported);
    return UNITY_END();
}
