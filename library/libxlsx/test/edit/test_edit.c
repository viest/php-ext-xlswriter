#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "libxlsx.h"
#include "libxlsx/edit.h"
#include "libxlsx/source_package.h"

void setUp(void) {}
void tearDown(void) {}

static const char *SOURCE_XLSX = "fixtures/edit_source.xlsx";
static const char *NOOP_XLSX = "fixtures/edit_noop.xlsx";
static const char *NUMBER_XLSX = "fixtures/edit_number.xlsx";
static const char *SNAPSHOT_XLSX = "fixtures/edit_snapshot.xlsx";
static const char *CUSTOM_TARGET_XLSX = "fixtures/edit_custom_target.xlsx";
static const char *NORMALIZED_TARGET_XLSX = "fixtures/edit_normalized_target.xlsx";

static void assert_ok(lxlsx_error err)
{
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR, err);
}

static void write_edit_workbook(const char *path, double a1, double d1)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_format *blank_format;
    lxlsx_row_col_options row_opts;

    memset(&row_opts, 0, sizeof(row_opts));
    row_opts.level = 1;

    remove(path);
    workbook = lxlsx_workbook_new(path);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_add_worksheet(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    blank_format = lxlsx_workbook_add_format(workbook);
    TEST_ASSERT_NOT_NULL(blank_format);
    lxlsx_format_set_bg_color(blank_format, 0xFFFF00);

    assert_ok(lxlsx_worksheet_set_column(worksheet, 0, 0, 20.0, NULL));
    assert_ok(lxlsx_worksheet_set_row_opt(worksheet, 0, 18.0, NULL, &row_opts));
    assert_ok(lxlsx_worksheet_write_number(worksheet, 0, 0, a1, blank_format));
    assert_ok(lxlsx_worksheet_write_formula_str(worksheet, 0, 1,
                                                "=\"O\"&\"K\"", NULL, "OK"));
    assert_ok(lxlsx_worksheet_write_blank(worksheet, 0, 2, blank_format));
    assert_ok(lxlsx_worksheet_write_number(worksheet, 0, 3, d1, NULL));
    assert_ok(lxlsx_worksheet_write_url(worksheet, 0, 4,
                                        "https://example.com", NULL));
    assert_ok(lxlsx_worksheet_merge_range(worksheet, 2, 0, 2, 1,
                                          "merged", NULL));
    lxlsx_worksheet_protect(worksheet, "pw", NULL);

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

static unsigned char *read_sheet_xml(const char *path, size_t *len)
{
    lxlsx_source_package *package = NULL;
    int sheet_index;
    unsigned char *xml = NULL;

    assert_ok(lxlsx_source_package_open(path, &package));
    sheet_index = lxlsx_source_package_find_first(package, "xl/worksheets/sheet1.xml");
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, sheet_index);
    assert_ok(lxlsx_source_package_read_entry(package, (size_t)sheet_index,
                                              &xml, len));
    lxlsx_source_package_close(package);
    return xml;
}

static void assert_xml_contains(const unsigned char *xml, const char *needle)
{
    TEST_ASSERT_NOT_NULL_MESSAGE(strstr((const char *)xml, needle), needle);
}

static void assert_number_cell_in_sheet(const char *path, const char *sheet_name,
                                        size_t row, size_t col, double expected)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    lxlsx_cell cell;
    int found = 0;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &workbook));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, sheet_name, LXLSX_READER_SKIP_NONE,
                              &worksheet));

    while (lxlsx_reader_worksheet_next_row(worksheet) == LXLSX_READER_NO_ERROR) {
        while (lxlsx_reader_worksheet_next_cell(worksheet, &cell)
               == LXLSX_READER_NO_ERROR) {
            if (cell.row_num == row && cell.col_num == col) {
                TEST_ASSERT_EQUAL_INT(NUMBER_CELL, cell.type);
                TEST_ASSERT_EQUAL_DOUBLE(expected, cell.value.number);
                found = 1;
            }
        }
    }

    TEST_ASSERT_TRUE(found);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

static void assert_number_cell(const char *path, size_t row, size_t col, double expected)
{
    assert_number_cell_in_sheet(path, "Edit", row, col, expected);
}

static void assert_formula_cell(const char *path, size_t row, size_t col,
                                const char *formula, const char *cached)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    lxlsx_cell cell;
    int found = 0;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &workbook));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, "Edit", LXLSX_READER_SKIP_NONE,
                              &worksheet));

    while (lxlsx_reader_worksheet_next_row(worksheet) == LXLSX_READER_NO_ERROR) {
        while (lxlsx_reader_worksheet_next_cell(worksheet, &cell)
               == LXLSX_READER_NO_ERROR) {
            if (cell.row_num == row && cell.col_num == col) {
                TEST_ASSERT_EQUAL_INT(FORMULA_CELL, cell.type);
                TEST_ASSERT_EQUAL_STRING_LEN(formula,
                                             cell.value.formula.formula.ptr,
                                             cell.value.formula.formula.len);
                TEST_ASSERT_EQUAL_STRING_LEN(cached,
                                             cell.value.formula.cached.ptr,
                                             cell.value.formula.cached.len);
                found = 1;
            }
        }
    }

    TEST_ASSERT_TRUE(found);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

static void assert_string_cell(const char *path, size_t row, size_t col,
                               const char *expected)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    lxlsx_cell cell;
    int found = 0;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &workbook));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, "Edit", LXLSX_READER_SKIP_NONE,
                              &worksheet));

    while (lxlsx_reader_worksheet_next_row(worksheet) == LXLSX_READER_NO_ERROR) {
        while (lxlsx_reader_worksheet_next_cell(worksheet, &cell)
               == LXLSX_READER_NO_ERROR) {
            if (cell.row_num == row && cell.col_num == col) {
                TEST_ASSERT_TRUE(cell.type == STRING_CELL ||
                                 cell.type == INLINE_STRING_CELL);
                TEST_ASSERT_EQUAL_STRING_LEN(expected,
                                             cell.value.string.ptr,
                                             cell.value.string.len);
                found = 1;
            }
        }
    }

    TEST_ASSERT_TRUE(found);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

static void assert_boolean_cell(const char *path, size_t row, size_t col,
                                int expected)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    lxlsx_cell cell;
    int found = 0;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &workbook));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, "Edit", LXLSX_READER_SKIP_NONE,
                              &worksheet));

    while (lxlsx_reader_worksheet_next_row(worksheet) == LXLSX_READER_NO_ERROR) {
        while (lxlsx_reader_worksheet_next_cell(worksheet, &cell)
               == LXLSX_READER_NO_ERROR) {
            if (cell.row_num == row && cell.col_num == col) {
                TEST_ASSERT_EQUAL_INT(BOOLEAN_CELL, cell.type);
                TEST_ASSERT_EQUAL_INT(expected ? 1 : 0, cell.value.boolean);
                found = 1;
            }
        }
    }

    TEST_ASSERT_TRUE(found);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

static void test_noop_edit_save_is_byte_identical(void)
{
    lxlsx_edit_session *session;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NOOP_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_save_as(session, NOOP_XLSX));
    lxlsx_edit_close(session);

    assert_files_equal(SOURCE_XLSX, NOOP_XLSX);
}

static void test_number_edit_preserves_worksheet_state(void)
{
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 0, 42.5));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_number_cell(NUMBER_XLSX, 1, 1, 42.5);

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    assert_xml_contains(xml, "<cols>");
    assert_xml_contains(xml, "customHeight=\"1\"");
    assert_xml_contains(xml, "<sheetProtection");
    assert_xml_contains(xml, "<mergeCells");
    assert_xml_contains(xml, "<hyperlinks>");
    assert_xml_contains(xml, "<c r=\"A1\" s=\"");
    assert_xml_contains(xml, "<c r=\"C1\" s=\"");
    assert_xml_contains(xml, "t=\"str\"");
    assert_xml_contains(xml, "<v>OK</v>");
    lxlsx_source_package_free_buffer(xml);
}

static void assert_relationship_target_edit(const char *source_path)
{
    lxlsx_edit_session *session;

    remove(NUMBER_XLSX);
    session = lxlsx_edit_open(source_path);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Custom", 0, 0, 17.0));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_number_cell_in_sheet(NUMBER_XLSX, "Custom", 1, 1, 17.0);
}

static void test_relationship_targets_locate_custom_worksheet_parts(void)
{
    assert_relationship_target_edit(CUSTOM_TARGET_XLSX);
    assert_relationship_target_edit(NORMALIZED_TARGET_XLSX);
}

static void test_formula_edit_roundtrips(void)
{
    lxlsx_edit_session *session;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_formula(session, "Edit", 1, 0, "=A1*2", "85"));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_formula_cell(NUMBER_XLSX, 2, 1, "A1*2", "85");
}

static void test_string_and_boolean_edit_roundtrip(void)
{
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_string(session, "Edit", 0, 0, "new text"));
    assert_ok(lxlsx_edit_set_boolean(session, "Edit", 1, 1, 1));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_string_cell(NUMBER_XLSX, 1, 1, "new text");
    assert_boolean_cell(NUMBER_XLSX, 2, 2, 1);

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    assert_xml_contains(xml, "t=\"inlineStr\"");
    assert_xml_contains(xml, "<v>1</v>");
    lxlsx_source_package_free_buffer(xml);
}

static void test_batched_edits_last_change_wins(void)
{
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 0, 10.0));
    assert_ok(lxlsx_edit_set_string(session, "Edit", 0, 0, "last value"));
    assert_ok(lxlsx_edit_set_boolean(session, "Edit", 1, 1, 1));
    assert_ok(lxlsx_edit_set_formula(session, "Edit", 1, 2, "=A1", "0"));
    assert_ok(lxlsx_edit_set_number(session, "Edit", 9, 3, 99.0));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_string_cell(NUMBER_XLSX, 1, 1, "last value");
    assert_boolean_cell(NUMBER_XLSX, 2, 2, 1);
    assert_formula_cell(NUMBER_XLSX, 2, 3, "A1", "0");
    assert_number_cell(NUMBER_XLSX, 10, 4, 99.0);

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    assert_xml_contains(xml, "<c r=\"A1\" s=\"");
    assert_xml_contains(xml, "t=\"inlineStr\"");
    assert_xml_contains(xml, "<row r=\"10\">");
    lxlsx_source_package_free_buffer(xml);
}

static void test_opened_workbook_uses_standard_write_api(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_format *format;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMBER_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    TEST_ASSERT_TRUE(lxlsx_workbook_is_edit(workbook));

    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    assert_ok(lxlsx_worksheet_write_string(worksheet, 0, 0, "via workbook", NULL));
    assert_ok(lxlsx_worksheet_write_number(worksheet, 1, 0, 123.0, NULL));
    assert_ok(lxlsx_worksheet_write_boolean(worksheet, 1, 1, 1, NULL));
    assert_ok(lxlsx_worksheet_write_formula(worksheet, 1, 2, "=A2*2", NULL));

    format = lxlsx_workbook_add_format(workbook);
    TEST_ASSERT_NOT_NULL(format);
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_write_string(worksheet, 2, 0,
                                                       "formatted", format));

    assert_ok(lxlsx_workbook_save_as(workbook, NUMBER_XLSX));
    lxlsx_workbook_free(workbook);

    assert_string_cell(NUMBER_XLSX, 1, 1, "via workbook");
    assert_number_cell(NUMBER_XLSX, 2, 1, 123.0);
    assert_boolean_cell(NUMBER_XLSX, 2, 2, 1);
    assert_formula_cell(NUMBER_XLSX, 2, 3, "A2*2", "0");
}

static void test_edit_uses_open_time_snapshot(void)
{
    lxlsx_edit_session *session;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(SNAPSHOT_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);

    write_edit_workbook(SOURCE_XLSX, 9.0, 999.0);

    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 0, 2.0));
    assert_ok(lxlsx_edit_save_as(session, SNAPSHOT_XLSX));
    lxlsx_edit_close(session);

    assert_number_cell(SNAPSHOT_XLSX, 1, 1, 2.0);
    assert_number_cell(SNAPSHOT_XLSX, 1, 4, 111.0);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_noop_edit_save_is_byte_identical);
    RUN_TEST(test_number_edit_preserves_worksheet_state);
    RUN_TEST(test_relationship_targets_locate_custom_worksheet_parts);
    RUN_TEST(test_formula_edit_roundtrips);
    RUN_TEST(test_string_and_boolean_edit_roundtrip);
    RUN_TEST(test_batched_edits_last_change_wins);
    RUN_TEST(test_opened_workbook_uses_standard_write_api);
    RUN_TEST(test_edit_uses_open_time_snapshot);
    return UNITY_END();
}
