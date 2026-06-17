#include <stdio.h>
#include <string.h>

#include <unity.h>

#include "lxlsx.h"
#include "lxlsx/reader.h"

void setUp(void) {}
void tearDown(void) {}

static const char *ROUNDTRIP_XLSX = "fixtures/integration_roundtrip.xlsx";

static void assert_write_ok(lxlsx_error err)
{
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR, err);
}

static void write_roundtrip_workbook(void)
{
    lxlsx_workbook *workbook = NULL;
    lxlsx_worksheet *worksheet = NULL;
    lxlsx_format *number_format = NULL;

    remove(ROUNDTRIP_XLSX);

    workbook = lxlsx_workbook_new(ROUNDTRIP_XLSX);
    TEST_ASSERT_NOT_NULL(workbook);

    worksheet = lxlsx_workbook_add_worksheet(workbook, "RoundTrip");
    TEST_ASSERT_NOT_NULL(worksheet);

    number_format = lxlsx_workbook_add_format(workbook);
    TEST_ASSERT_NOT_NULL(number_format);
    lxlsx_format_set_num_format(number_format, "0.00");

    assert_write_ok(lxlsx_worksheet_write_number(worksheet, 0, 0, 123.5, number_format));
    assert_write_ok(lxlsx_worksheet_write_string(worksheet, 0, 1, "alpha", NULL));
    assert_write_ok(lxlsx_worksheet_write_formula_num(worksheet, 0, 2, "=A1*2", NULL, 247.0));

    assert_write_ok(lxlsx_workbook_close(workbook));
}

static void test_writer_output_is_readable_by_unified_reader(void)
{
    lxlsx_reader_workbook *workbook = NULL;
    lxlsx_reader_worksheet *worksheet = NULL;
    const lxlsx_reader_styles *styles = NULL;
    const lxlsx_reader_xf *xf = NULL;
    lxlsx_cell cell;
    uint32_t number_style_ref = 0;

    write_roundtrip_workbook();

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(ROUNDTRIP_XLSX, &workbook));
    TEST_ASSERT_EQUAL_INT(1, (int)lxlsx_reader_workbook_sheet_count(workbook));
    TEST_ASSERT_EQUAL_STRING("RoundTrip",
                             lxlsx_reader_workbook_sheet_name(workbook, 0));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              workbook, "RoundTrip", LXLSX_READER_SKIP_NONE, &worksheet));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_worksheet_next_row(worksheet));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_worksheet_next_cell(worksheet, &cell));
    TEST_ASSERT_EQUAL_INT(NUMBER_CELL, cell.type);
    TEST_ASSERT_EQUAL_INT(1, (int)cell.row_num);
    TEST_ASSERT_EQUAL_INT(1, (int)cell.col_num);
    TEST_ASSERT_EQUAL_DOUBLE(123.5, cell.value.number);
    TEST_ASSERT_GREATER_THAN(0, (int)cell.style_ref);
    number_style_ref = cell.style_ref;

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_worksheet_next_cell(worksheet, &cell));
    TEST_ASSERT_EQUAL_INT(STRING_CELL, cell.type);
    TEST_ASSERT_EQUAL_INT(1, (int)cell.row_num);
    TEST_ASSERT_EQUAL_INT(2, (int)cell.col_num);
    TEST_ASSERT_EQUAL_STRING_LEN("alpha", cell.value.string.ptr, 5);

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_worksheet_next_cell(worksheet, &cell));
    TEST_ASSERT_EQUAL_INT(FORMULA_CELL, cell.type);
    TEST_ASSERT_EQUAL_INT(1, (int)cell.row_num);
    TEST_ASSERT_EQUAL_INT(3, (int)cell.col_num);
    TEST_ASSERT_EQUAL_STRING_LEN("A1*2",
                                 cell.value.formula.formula.ptr,
                                 cell.value.formula.formula.len);
    TEST_ASSERT_EQUAL_STRING_LEN("247",
                                 cell.value.formula.cached.ptr,
                                 cell.value.formula.cached.len);

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_END_OF_DATA,
                          lxlsx_reader_worksheet_next_cell(worksheet, &cell));

    styles = lxlsx_reader_workbook_get_styles(workbook);
    TEST_ASSERT_NOT_NULL(styles);
    xf = lxlsx_reader_styles_get_xf(styles, number_style_ref);
    TEST_ASSERT_NOT_NULL(xf);
    TEST_ASSERT_NOT_NULL(xf->format_string);
    TEST_ASSERT_EQUAL_STRING("0.00", xf->format_string);

    lxlsx_reader_worksheet_close(worksheet);
    lxlsx_reader_workbook_close(workbook);
    remove(ROUNDTRIP_XLSX);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_writer_output_is_readable_by_unified_reader);
    return UNITY_END();
}
