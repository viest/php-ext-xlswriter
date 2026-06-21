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
static const char *NAMESPACE_XLSX = "fixtures/edit_namespace.xlsx";
static const char *SELF_CLOSING_XLSX = "fixtures/edit_self_closing.xlsx";
static const char *ORDERED_ROWS_XLSX = "fixtures/edit_ordered_rows.xlsx";
static const char *CUSTOM_TARGET_XLSX = "fixtures/edit_custom_target.xlsx";
static const char *NORMALIZED_TARGET_XLSX = "fixtures/edit_normalized_target.xlsx";
static const char *MERGE_XLSX = "fixtures/edit_merge.xlsx";
static const char *NUMFMT_XLSX = "fixtures/edit_numfmt.xlsx";
static const char *DIM_XLSX = "fixtures/edit_dim.xlsx";
static const char *ADDSHEET_XLSX = "fixtures/edit_addsheet.xlsx";
static const char *IMAGE_XLSX = "fixtures/edit_image.xlsx";
static const char *CHART_XLSX = "fixtures/edit_chart.xlsx";

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

static void write_sheet_xml_override(const char *source_path,
                                     const char *output_path,
                                     const char *sheet_xml)
{
    lxlsx_source_package *package = NULL;
    lxlsx_source_package_replacement replacement;
    int sheet_index;

    remove(output_path);
    assert_ok(lxlsx_source_package_open(source_path, &package));
    sheet_index = lxlsx_source_package_find_first(package,
                                                  "xl/worksheets/sheet1.xml");
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, sheet_index);

    replacement.entry_index = (size_t)sheet_index;
    replacement.data = (const unsigned char *)sheet_xml;
    replacement.size = strlen(sheet_xml);
    assert_ok(lxlsx_source_package_save_with_replacements(
        package, output_path, &replacement, 1));

    lxlsx_source_package_close(package);
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
                TEST_ASSERT_EQUAL_DOUBLE(expected, cell.data.reader.value.number);
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
                                             cell.data.reader.value.formula->formula.ptr,
                                             cell.data.reader.value.formula->formula.len);
                TEST_ASSERT_EQUAL_STRING_LEN(cached,
                                             cell.data.reader.value.formula->cached.ptr,
                                             cell.data.reader.value.formula->cached.len);
                found = 1;
            }
        }
    }

    TEST_ASSERT_TRUE(found);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

static void assert_string_cell_in_sheet(const char *path, const char *sheet_name,
                                        size_t row, size_t col,
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
                              workbook, sheet_name, LXLSX_READER_SKIP_NONE,
                              &worksheet));

    while (lxlsx_reader_worksheet_next_row(worksheet) == LXLSX_READER_NO_ERROR) {
        while (lxlsx_reader_worksheet_next_cell(worksheet, &cell)
               == LXLSX_READER_NO_ERROR) {
            if (cell.row_num == row && cell.col_num == col) {
                TEST_ASSERT_TRUE(cell.type == STRING_CELL ||
                                 cell.type == INLINE_STRING_CELL);
                TEST_ASSERT_EQUAL_STRING_LEN(expected,
                                             cell.data.reader.value.string.ptr,
                                             cell.data.reader.value.string.len);
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
    assert_string_cell_in_sheet(path, "Edit", row, col, expected);
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
                TEST_ASSERT_EQUAL_INT(expected ? 1 : 0, cell.data.reader.value.boolean);
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

static void test_edit_tokenizer_handles_namespaces_and_special_sections(void)
{
    static const char sheet_xml[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<x:worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<x:sheetData>"
        "<!-- <x:row r=\"1\"><x:c r=\"A1\"><x:v>999</x:v></x:c></x:row> -->"
        "<![CDATA[<x:row r=\"1\"><x:c r=\"B1\"><x:v>999</x:v></x:c></x:row>]]>"
        "<x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row>"
        "</x:sheetData>"
        "</x:worksheet>";
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    write_sheet_xml_override(SOURCE_XLSX, NAMESPACE_XLSX, sheet_xml);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(NAMESPACE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 0, 77.0));
    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 1, 88.0));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_number_cell(NUMBER_XLSX, 1, 1, 77.0);
    assert_number_cell(NUMBER_XLSX, 1, 2, 88.0);

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    assert_xml_contains(xml, "<!-- <x:row r=\"1\"");
    assert_xml_contains(xml, "<![CDATA[<x:row r=\"1\"");
    assert_xml_contains(xml, "<c r=\"A1\"");
    assert_xml_contains(xml, "<c r=\"B1\"");
    lxlsx_source_package_free_buffer(xml);
    remove(NAMESPACE_XLSX);
}

static void test_edit_expands_namespaced_self_closing_sheet_data(void)
{
    static const char sheet_xml[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<x:worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<x:sheetData/>"
        "</x:worksheet>";
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;
    const char *worksheet;
    const char *sheet_data;
    const char *row;
    const char *close_sheet_data;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    write_sheet_xml_override(SOURCE_XLSX, SELF_CLOSING_XLSX, sheet_xml);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(SELF_CLOSING_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_string(session, "Edit", 0, 0, "a&<>"));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_string_cell(NUMBER_XLSX, 1, 1, "a&<>");

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    worksheet = strstr((const char *)xml, "<x:worksheet");
    sheet_data = strstr((const char *)xml, "<x:sheetData>");
    row = strstr((const char *)xml, "<row r=\"1\"");
    close_sheet_data = strstr((const char *)xml, "</x:sheetData>");
    TEST_ASSERT_NOT_NULL(worksheet);
    TEST_ASSERT_NOT_NULL(sheet_data);
    TEST_ASSERT_NOT_NULL(row);
    TEST_ASSERT_NOT_NULL(close_sheet_data);
    TEST_ASSERT_TRUE(worksheet < sheet_data);
    TEST_ASSERT_TRUE(sheet_data < row);
    TEST_ASSERT_TRUE(row < close_sheet_data);
    assert_xml_contains(xml, "a&amp;&lt;&gt;");
    lxlsx_source_package_free_buffer(xml);
    remove(SELF_CLOSING_XLSX);
}

static void test_edit_inserts_missing_rows_inside_sheet_data_order(void)
{
    static const char sheet_xml[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<sheetData>"
        "<row r=\"5\"><c r=\"A5\"><v>5</v></c></row>"
        "</sheetData>"
        "</worksheet>";
    lxlsx_edit_session *session;
    unsigned char *xml;
    size_t xml_len;
    const char *sheet_data;
    const char *row1;
    const char *row5;
    const char *close_sheet_data;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    write_sheet_xml_override(SOURCE_XLSX, ORDERED_ROWS_XLSX, sheet_xml);
    remove(NUMBER_XLSX);

    session = lxlsx_edit_open(ORDERED_ROWS_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Edit", 0, 0, 1.0));
    assert_ok(lxlsx_edit_set_number(session, "Edit", 4, 0, 55.0));
    assert_ok(lxlsx_edit_save_as(session, NUMBER_XLSX));
    lxlsx_edit_close(session);

    assert_number_cell(NUMBER_XLSX, 1, 1, 1.0);
    assert_number_cell(NUMBER_XLSX, 5, 1, 55.0);

    xml = read_sheet_xml(NUMBER_XLSX, &xml_len);
    (void)xml_len;
    sheet_data = strstr((const char *)xml, "<sheetData>");
    row1 = strstr((const char *)xml, "<row r=\"1\"");
    row5 = strstr((const char *)xml, "<row r=\"5\"");
    close_sheet_data = strstr((const char *)xml, "</sheetData>");
    TEST_ASSERT_NOT_NULL(sheet_data);
    TEST_ASSERT_NOT_NULL(row1);
    TEST_ASSERT_NOT_NULL(row5);
    TEST_ASSERT_NOT_NULL(close_sheet_data);
    TEST_ASSERT_TRUE(sheet_data < row1);
    TEST_ASSERT_TRUE(row1 < row5);
    TEST_ASSERT_TRUE(row5 < close_sheet_data);
    lxlsx_source_package_free_buffer(xml);
    remove(ORDERED_ROWS_XLSX);
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

    /* A format applied in edit mode is now supported (injected into styles). */
    format = lxlsx_workbook_add_format(workbook);
    TEST_ASSERT_NOT_NULL(format);
    lxlsx_format_set_bold(format);
    assert_ok(lxlsx_worksheet_write_string(worksheet, 2, 0, "formatted", format));

    assert_ok(lxlsx_workbook_save_as(workbook, NUMBER_XLSX));
    lxlsx_workbook_free(workbook);

    assert_string_cell(NUMBER_XLSX, 1, 1, "via workbook");
    assert_number_cell(NUMBER_XLSX, 2, 1, 123.0);
    assert_boolean_cell(NUMBER_XLSX, 2, 2, 1);
    assert_formula_cell(NUMBER_XLSX, 2, 3, "A2*2", "0");
}

static void test_edit_rejects_unsupported_worksheet_ops(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_datetime dt;

    memset(&dt, 0, sizeof(dt));

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    TEST_ASSERT_TRUE(lxlsx_workbook_is_edit(workbook));

    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    /* Snapshot-based editing can only patch number/string/boolean/formula
     * cells. Everything else must report FEATURE_NOT_SUPPORTED instead of
     * mutating in-memory state that the snapshot save never serializes. */
    (void) dt;  /* write_datetime is supported in edit mode (see merge/numfmt tests) */
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_write_unixtime(worksheet, 3, 1, 0, NULL));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_write_url(worksheet, 4, 0, "https://x", NULL));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_write_comment(worksheet, 5, 0, "note"));
    /* insert_image and insert_chart are supported in edit mode now (see the
     * add-image / add-chart tests). */
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_autofilter(worksheet, 0, 0, 1, 1));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_add_table(worksheet, 0, 0, 1, 1, NULL));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_data_validation_cell(worksheet, 0, 0, NULL));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_freeze_panes(worksheet, 1, 0));
    TEST_ASSERT_EQUAL_INT(LXLSX_ERROR_FEATURE_NOT_SUPPORTED,
                          lxlsx_worksheet_show_comments(worksheet));

    lxlsx_workbook_free(workbook);
}

static void test_edit_adds_merged_range(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    unsigned char *xml;
    size_t xml_len;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);  /* already merges A3:B3 */
    remove(MERGE_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    /* Merge is supported in edit mode: it is injected on save. */
    assert_ok(lxlsx_worksheet_merge_range(worksheet, 5, 0, 5, 2, "merged", NULL));
    assert_ok(lxlsx_workbook_save_as(workbook, MERGE_XLSX));
    lxlsx_workbook_free(workbook);

    xml = read_sheet_xml(MERGE_XLSX, &xml_len);
    assert_xml_contains(xml, "<mergeCells count=\"2\">"); /* count bumped 1 -> 2 */
    assert_xml_contains(xml, "<mergeCell ref=\"A3:B3\"/>"); /* original preserved */
    assert_xml_contains(xml, "<mergeCell ref=\"A6:C6\"/>"); /* new range added */
    free(xml);
}

static void test_edit_injects_number_format(void)
{
    lxlsx_edit_session *session;
    lxlsx_source_package *package;
    int idx;
    unsigned char *styles = NULL;
    size_t styles_len = 0;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMFMT_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_number(session, "Edit", 5, 0, 1234.5));
    assert_ok(lxlsx_edit_set_number_format(session, "Edit", 5, 0, "0.00"));
    assert_ok(lxlsx_edit_save_as(session, NUMFMT_XLSX));
    lxlsx_edit_close(session);

    /* The value is written and styles.xml carries the injected number format. */
    assert_number_cell(NUMFMT_XLSX, 6, 1, 1234.5);

    assert_ok(lxlsx_source_package_open(NUMFMT_XLSX, &package));
    idx = lxlsx_source_package_find_first(package, "xl/styles.xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(package, (size_t)idx,
                                              &styles, &styles_len));
    assert_xml_contains(styles, "formatCode=\"0.00\"");
    assert_xml_contains(styles, "applyNumberFormat=\"1\"");
    free(styles);
    lxlsx_source_package_close(package);
}

static void test_edit_injects_cell_format(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_format *format;
    lxlsx_source_package *package;
    int idx;
    unsigned char *styles = NULL, *sheet = NULL;
    size_t styles_len = 0, sheet_len = 0;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);
    remove(NUMFMT_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    format = lxlsx_workbook_add_format(workbook);
    TEST_ASSERT_NOT_NULL(format);
    lxlsx_format_set_bold(format);
    lxlsx_format_set_font_color(format, 0xFF0000);
    lxlsx_format_set_fg_color(format, 0x00FF00);
    lxlsx_format_set_pattern(format, LXLSX_PATTERN_SOLID);
    lxlsx_format_set_align(format, LXLSX_ALIGN_CENTER);
    lxlsx_format_set_border(format, LXLSX_BORDER_THIN);
    assert_ok(lxlsx_worksheet_write_string(worksheet, 5, 0, "styled", format));
    assert_ok(lxlsx_workbook_save_as(workbook, NUMFMT_XLSX));
    lxlsx_workbook_free(workbook);

    assert_ok(lxlsx_source_package_open(NUMFMT_XLSX, &package));
    idx = lxlsx_source_package_find_first(package, "xl/styles.xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(package, (size_t)idx, &styles, &styles_len));
    assert_xml_contains(styles, "<b/>");
    assert_xml_contains(styles, "rgb=\"FFFF0000\"");          /* font color */
    assert_xml_contains(styles, "patternType=\"solid\"");     /* fill */
    assert_xml_contains(styles, "rgb=\"FF00FF00\"");          /* fill fg color */
    assert_xml_contains(styles, "horizontal=\"center\"");     /* alignment */
    assert_xml_contains(styles, "<left style=\"thin\">");     /* border */
    assert_xml_contains(styles, "applyBorder=\"1\"");
    free(styles);

    idx = lxlsx_source_package_find_first(package, "xl/worksheets/sheet1.xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(package, (size_t)idx, &sheet, &sheet_len));
    assert_xml_contains(sheet, "<c r=\"A6\" s=");             /* cell repointed */
    free(sheet);
    lxlsx_source_package_close(package);
}

static void test_edit_adds_worksheet(void)
{
    static const char *SHEET_XML =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hi</t></is></c></row></sheetData>"
        "</worksheet>";
    lxlsx_edit_session *session;
    lxlsx_source_package *pkg = NULL;
    unsigned char *wb = NULL;
    size_t wb_len = 0;
    int idx;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);  /* one sheet "Edit" */
    remove(ADDSHEET_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_add_sheet(session, "Added", SHEET_XML, strlen(SHEET_XML)));
    assert_ok(lxlsx_edit_save_as(session, ADDSHEET_XLSX));
    lxlsx_edit_close(session);

    /* Reopening must show two sheets (workbook.xml + rels stayed consistent). */
    session = lxlsx_edit_open(ADDSHEET_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    TEST_ASSERT_EQUAL_INT(2, (int)lxlsx_edit_sheet_count(session));
    TEST_ASSERT_EQUAL_STRING("Added", lxlsx_edit_sheet_name(session, 1));
    lxlsx_edit_close(session);

    /* The new part exists and the metadata files were patched. */
    assert_ok(lxlsx_source_package_open(ADDSHEET_XLSX, &pkg));
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/worksheets/sheet2.xml") >= 0);
    idx = lxlsx_source_package_find_first(pkg, "xl/workbook.xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &wb, &wb_len));
    assert_xml_contains(wb, "<sheet name=\"Added\"");
    free(wb);
    idx = lxlsx_source_package_find_first(pkg, "[Content_Types].xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &wb, &wb_len));
    assert_xml_contains(wb, "/xl/worksheets/sheet2.xml");
    free(wb);
    lxlsx_source_package_close(pkg);
}

/*
 * End-to-end add-sheet via the standard workbook write API (the path the PHP
 * addSheet() takes in edit mode): open, add_worksheet, write with the normal
 * setters, save. The new sheet must serialise to inline strings and be
 * registered without disturbing the existing sheet.
 */
static void test_edit_add_sheet_via_workbook_api(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *added;
    lxlsx_edit_session *session;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);  /* one sheet "Edit" */
    remove(ADDSHEET_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    TEST_ASSERT_TRUE(lxlsx_workbook_is_edit(workbook));

    added = lxlsx_workbook_add_worksheet(workbook, "Added");
    TEST_ASSERT_NOT_NULL(added);
    assert_ok(lxlsx_worksheet_write_string(added, 0, 0, "hello new", NULL));
    assert_ok(lxlsx_worksheet_write_number(added, 1, 0, 99.0, NULL));
    assert_ok(lxlsx_worksheet_write_string(added, 2, 0, "second", NULL));

    assert_ok(lxlsx_workbook_save_as(workbook, ADDSHEET_XLSX));
    lxlsx_workbook_free(workbook);

    /* Two sheets: original first, new one appended in workbook order. */
    session = lxlsx_edit_open(ADDSHEET_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    TEST_ASSERT_EQUAL_INT(2, (int)lxlsx_edit_sheet_count(session));
    TEST_ASSERT_EQUAL_STRING("Edit", lxlsx_edit_sheet_name(session, 0));
    TEST_ASSERT_EQUAL_STRING("Added", lxlsx_edit_sheet_name(session, 1));
    lxlsx_edit_close(session);

    /* Original sheet survives untouched. */
    assert_number_cell_in_sheet(ADDSHEET_XLSX, "Edit", 1, 1, 1.0);
    /* New sheet's inline-string + number values read back. */
    assert_string_cell_in_sheet(ADDSHEET_XLSX, "Added", 1, 1, "hello new");
    assert_number_cell_in_sheet(ADDSHEET_XLSX, "Added", 2, 1, 99.0);
    assert_string_cell_in_sheet(ADDSHEET_XLSX, "Added", 3, 1, "second");
}

static void write_protected_no_merge(const char *path);

/* A minimal valid 1x1 RGBA PNG (no pHYs chunk, so DPI defaults to 96). */
static const unsigned char PNG_1X1[] = {
    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
    0x89,0x00,0x00,0x00,0x0B,0x49,0x44,0x41,0x54,0x78,0x9C,0x63,0xFA,0xCF,0x00,0x00,
    0x02,0x07,0x01,0x02,0x9A,0x1C,0x31,0x71,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
    0xAE,0x42,0x60,0x82
};

/*
 * Insert an image into an existing worksheet whose _rels part does NOT exist yet
 * (no hyperlink in the source), exercising the "create worksheet rels" branch.
 * The media/drawing/drawing-rels parts and the content-type Default are
 * synthesised, the worksheet gets a <drawing> ref, and — crucially — the new
 * rels part must be complete so the reader can still parse the sheet back.
 */
static void test_edit_inserts_image(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_source_package *pkg = NULL;
    unsigned char *xml = NULL;
    size_t xml_len = 0;
    int idx;

    write_protected_no_merge(SOURCE_XLSX);  /* sheet "Edit", no _rels yet */
    remove(IMAGE_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    assert_ok(lxlsx_worksheet_insert_image_buffer_opt(worksheet, 2, 1,
                                                      PNG_1X1, sizeof(PNG_1X1),
                                                      NULL));
    assert_ok(lxlsx_workbook_save_as(workbook, IMAGE_XLSX));
    lxlsx_workbook_free(workbook);

    assert_ok(lxlsx_source_package_open(IMAGE_XLSX, &pkg));
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/media/image1.png") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/_rels/drawing1.xml.rels") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/worksheets/_rels/sheet1.xml.rels") >= 0);

    /* The worksheet references the drawing. */
    idx = lxlsx_source_package_find_first(pkg, "xl/worksheets/sheet1.xml");
    TEST_ASSERT_TRUE(idx >= 0);
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "<drawing r:id=");
    free(xml); xml = NULL;

    /* The drawing anchors the image via a oneCellAnchor + blip. */
    idx = lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "<xdr:oneCellAnchor>");
    assert_xml_contains(xml, "r:embed=\"rId1\"");
    free(xml); xml = NULL;

    /* The drawing rels point at the media part. */
    idx = lxlsx_source_package_find_first(pkg, "xl/drawings/_rels/drawing1.xml.rels");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "Target=\"../media/image1.png\"");
    free(xml); xml = NULL;

    /* Content types declare the png default. */
    idx = lxlsx_source_package_find_first(pkg, "[Content_Types].xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "Extension=\"png\"");
    free(xml); xml = NULL;

    lxlsx_source_package_close(pkg);

    /* The original cell data still reads back through the reader. */
    assert_number_cell_in_sheet(IMAGE_XLSX, "Edit", 1, 1, 1.0);
}

static int count_chart_cb(const lxlsx_reader_chart_meta *info, void *ud)
{
    (void)info;
    (*(int *)ud)++;
    return 0;
}

/* Assert the reader parses back `expected` charts on `sheet` — this exercises
 * the full drawing -> chart rels -> chartN.xml chain through a real XML parser,
 * so a malformed drawing (e.g. mismatched tags) shows up as a count mismatch. */
static void assert_chart_count_in_sheet(const char *path, const char *sheet_name,
                                        int expected)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    int count = 0;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &workbook));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, sheet_name, LXLSX_READER_SKIP_NONE,
                              &worksheet));
    lxlsx_reader_worksheet_iterate_charts(worksheet, count_chart_cb, &count);
    TEST_ASSERT_EQUAL_INT(expected, count);
    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
}

/*
 * Insert a chart into an existing worksheet (no _rels yet). The chart is
 * serialised to xl/charts/chart1.xml and anchored via a graphicFrame; the
 * drawing/rels/content-types are synthesised and the original data still reads
 * back through the reader.
 */
static void test_edit_inserts_chart(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_chart *chart;
    lxlsx_source_package *pkg = NULL;
    unsigned char *xml = NULL;
    size_t xml_len = 0;
    int idx;

    write_protected_no_merge(SOURCE_XLSX);  /* sheet "Edit", A1 = 1.0 */
    remove(CHART_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_COLUMN);
    TEST_ASSERT_NOT_NULL(chart);
    lxlsx_chart_add_series(chart, NULL, "Edit!$A$1:$A$1");

    assert_ok(lxlsx_worksheet_insert_chart(worksheet, 5, 1, chart));
    assert_ok(lxlsx_workbook_save_as(workbook, CHART_XLSX));
    lxlsx_workbook_free(workbook);

    assert_ok(lxlsx_source_package_open(CHART_XLSX, &pkg));
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/charts/chart1.xml") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/_rels/drawing1.xml.rels") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/worksheets/_rels/sheet1.xml.rels") >= 0);

    idx = lxlsx_source_package_find_first(pkg, "xl/worksheets/sheet1.xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "<drawing r:id=");
    free(xml); xml = NULL;

    idx = lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "<xdr:graphicFrame");
    assert_xml_contains(xml, "<c:chart");
    assert_xml_contains(xml, "r:id=\"rId1\"");
    /* The ext must be emitted in full (guards the anchor-buffer size). */
    assert_xml_contains(xml, "<xdr:ext cx=\"4572000\" cy=\"2743200\"/>");
    free(xml); xml = NULL;

    idx = lxlsx_source_package_find_first(pkg, "xl/drawings/_rels/drawing1.xml.rels");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "Target=\"../charts/chart1.xml\"");
    free(xml); xml = NULL;

    idx = lxlsx_source_package_find_first(pkg, "[Content_Types].xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "drawingml.chart+xml");
    free(xml); xml = NULL;

    lxlsx_source_package_close(pkg);

    /* The original cell data still reads back (sheet + new rels are valid). */
    assert_number_cell_in_sheet(CHART_XLSX, "Edit", 1, 1, 1.0);
    /* And the reader parses the chart back through real XML parsing. */
    assert_chart_count_in_sheet(CHART_XLSX, "Edit", 1);
}

/*
 * An image and a chart on the SAME sheet must share one drawing part (one
 * <drawing> ref), holding a pic anchor and a graphicFrame anchor — the shared
 * drawing model. Previously this was rejected.
 */
static void test_edit_mixes_image_and_chart(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    lxlsx_chart *chart;
    lxlsx_source_package *pkg = NULL;
    unsigned char *xml = NULL;
    size_t xml_len = 0;
    int idx;

    write_protected_no_merge(SOURCE_XLSX);
    remove(CHART_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);

    assert_ok(lxlsx_worksheet_insert_image_buffer_opt(worksheet, 2, 1,
                                                      PNG_1X1, sizeof(PNG_1X1),
                                                      NULL));
    chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_COLUMN);
    TEST_ASSERT_NOT_NULL(chart);
    lxlsx_chart_add_series(chart, NULL, "Edit!$A$1:$A$1");
    assert_ok(lxlsx_worksheet_insert_chart(worksheet, 6, 1, chart));

    assert_ok(lxlsx_workbook_save_as(workbook, CHART_XLSX));
    lxlsx_workbook_free(workbook);

    assert_ok(lxlsx_source_package_open(CHART_XLSX, &pkg));
    /* Exactly one drawing, carrying both anchors. */
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml") >= 0);
    TEST_ASSERT_TRUE(lxlsx_source_package_find_first(pkg, "xl/drawings/drawing2.xml") < 0);
    idx = lxlsx_source_package_find_first(pkg, "xl/drawings/drawing1.xml");
    assert_ok(lxlsx_source_package_read_entry(pkg, (size_t)idx, &xml, &xml_len));
    assert_xml_contains(xml, "<xdr:pic>");
    assert_xml_contains(xml, "<xdr:graphicFrame");
    free(xml); xml = NULL;
    lxlsx_source_package_close(pkg);

    /* Reader sees the chart; the original data still reads back. */
    assert_chart_count_in_sheet(CHART_XLSX, "Edit", 1);
    assert_number_cell_in_sheet(CHART_XLSX, "Edit", 1, 1, 1.0);
}

static void test_edit_sets_dimensions(void)
{
    lxlsx_edit_session *session;
    unsigned char *xml = NULL;
    size_t xml_len = 0;

    write_edit_workbook(SOURCE_XLSX, 1.0, 111.0);  /* row 1 has data */
    remove(DIM_XLSX);

    session = lxlsx_edit_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(session);
    assert_ok(lxlsx_edit_set_column(session, "Edit", 0, 2, 25.0));   /* A:C */
    assert_ok(lxlsx_edit_set_row_height(session, "Edit", 0, 30.0));  /* row 1 */
    assert_ok(lxlsx_edit_save_as(session, DIM_XLSX));
    lxlsx_edit_close(session);

    xml = read_sheet_xml(DIM_XLSX, &xml_len);
    assert_xml_contains(xml, "<col min=\"1\" max=\"3\" width=\"25\" customWidth=\"1\"/>");
    assert_xml_contains(xml, "ht=\"30\"");
    assert_xml_contains(xml, "customHeight=\"1\"");
    free(xml);
}

static void write_protected_no_merge(const char *path)
{
    lxlsx_workbook *wb;
    lxlsx_worksheet *ws;
    remove(path);
    wb = lxlsx_workbook_new(path);
    TEST_ASSERT_NOT_NULL(wb);
    ws = lxlsx_workbook_add_worksheet(wb, "Edit");
    TEST_ASSERT_NOT_NULL(ws);
    assert_ok(lxlsx_worksheet_write_number(ws, 0, 0, 1.0, NULL));
    lxlsx_worksheet_protect(ws, "pw", NULL);   /* <sheetProtection>, no merge */
    assert_ok(lxlsx_workbook_close(wb));
}

static void test_edit_merge_orders_after_protection(void)
{
    lxlsx_workbook *workbook;
    lxlsx_worksheet *worksheet;
    unsigned char *xml = NULL;
    size_t xml_len = 0;
    const char *prot, *merge;

    /* Source has <sheetProtection> (after sheetData) but no <mergeCells>. */
    write_protected_no_merge(SOURCE_XLSX);
    remove(MERGE_XLSX);

    workbook = lxlsx_workbook_open(SOURCE_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);
    worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Edit");
    TEST_ASSERT_NOT_NULL(worksheet);
    assert_ok(lxlsx_worksheet_merge_range(worksheet, 4, 0, 4, 2, "m", NULL));
    assert_ok(lxlsx_workbook_save_as(workbook, MERGE_XLSX));
    lxlsx_workbook_free(workbook);

    /* The created <mergeCells> must sort AFTER <sheetProtection>. */
    xml = read_sheet_xml(MERGE_XLSX, &xml_len);
    prot = strstr((const char *)xml, "<sheetProtection");
    merge = strstr((const char *)xml, "<mergeCells");
    TEST_ASSERT_NOT_NULL(prot);
    TEST_ASSERT_NOT_NULL(merge);
    TEST_ASSERT_TRUE(merge > prot);
    free(xml);
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
    RUN_TEST(test_edit_tokenizer_handles_namespaces_and_special_sections);
    RUN_TEST(test_edit_expands_namespaced_self_closing_sheet_data);
    RUN_TEST(test_edit_inserts_missing_rows_inside_sheet_data_order);
    RUN_TEST(test_opened_workbook_uses_standard_write_api);
    RUN_TEST(test_edit_rejects_unsupported_worksheet_ops);
    RUN_TEST(test_edit_adds_merged_range);
    RUN_TEST(test_edit_merge_orders_after_protection);
    RUN_TEST(test_edit_injects_number_format);
    RUN_TEST(test_edit_injects_cell_format);
    RUN_TEST(test_edit_sets_dimensions);
    RUN_TEST(test_edit_adds_worksheet);
    RUN_TEST(test_edit_add_sheet_via_workbook_api);
    RUN_TEST(test_edit_inserts_image);
    RUN_TEST(test_edit_inserts_chart);
    RUN_TEST(test_edit_mixes_image_and_chart);
    RUN_TEST(test_edit_uses_open_time_snapshot);
    return UNITY_END();
}
