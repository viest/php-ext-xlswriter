/* Unity tests for Phase 3: formula attrs, data validations, autofilter. */

#include <string.h>

#include <unity.h>

#include "xlsx_test_paths.h"
#include "libxlsx.h"

void setUp(void) {}
void tearDown(void) {}

static void test_formula_array(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_cell       c;
    lxlsx_reader_workbook_open(LXLSX_TEST_PHASE3_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    /* row 1 */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    /* A1 has the array master formula */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(FORMULA_CELL, c.type);
    TEST_ASSERT_EQUAL_INT(LXLSX_FORMULA_ARRAY, c.value.formula.kind);
    TEST_ASSERT_EQUAL_STRING_LEN("A1:B2", c.value.formula.ref.ptr,
                                 c.value.formula.ref.len);
    TEST_ASSERT_EQUAL_STRING_LEN("TRANSPOSE(C1:C2)",
                                 c.value.formula.formula.ptr,
                                 c.value.formula.formula.len);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_formula_shared(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_cell       c;
    lxlsx_reader_workbook_open(LXLSX_TEST_PHASE3_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    /* skip rows 1, 2 */
    lxlsx_reader_worksheet_skip_rows(ws, 2);

    /* row 3: shared master at A3 */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXLSX_FORMULA_SHARED, c.value.formula.kind);
    TEST_ASSERT_EQUAL_INT(0, c.value.formula.si);
    TEST_ASSERT_EQUAL_STRING_LEN("A3:A5", c.value.formula.ref.ptr,
                                 c.value.formula.ref.len);
    TEST_ASSERT_EQUAL_STRING_LEN("B3*2", c.value.formula.formula.ptr,
                                 c.value.formula.formula.len);

    /* row 4: shared follower */
    /* drain remaining cells of row 3 */
    while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXLSX_FORMULA_SHARED, c.value.formula.kind);
    TEST_ASSERT_EQUAL_INT(0, c.value.formula.si);
    /* No expression on the follower. */
    TEST_ASSERT_EQUAL_INT(0, (int)c.value.formula.formula.len);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_data_validations(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_reader_data_validation d;
    lxlsx_reader_workbook_open(LXLSX_TEST_PHASE3_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_size_t(2, lxlsx_reader_worksheet_data_validation_count(ws));

    TEST_ASSERT_EQUAL_INT(1, lxlsx_reader_worksheet_data_validation_get(ws, 0, &d));
    TEST_ASSERT_EQUAL_STRING("whole",   d.type);
    TEST_ASSERT_EQUAL_STRING("between", d.operator_);
    TEST_ASSERT_EQUAL_STRING("D2:D8",   d.sqref);
    TEST_ASSERT_EQUAL_STRING("1",       d.formula1);
    TEST_ASSERT_EQUAL_STRING("100",     d.formula2);
    TEST_ASSERT_EQUAL_INT(1, d.allow_blank);
    TEST_ASSERT_EQUAL_INT(1, d.show_input_message);

    TEST_ASSERT_EQUAL_INT(1, lxlsx_reader_worksheet_data_validation_get(ws, 1, &d));
    TEST_ASSERT_EQUAL_STRING("list",    d.type);
    TEST_ASSERT_EQUAL_STRING("C1:C2",   d.sqref);
    TEST_ASSERT_EQUAL_STRING("\"alpha,beta,gamma\"", d.formula1);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_autofilter(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_reader_autofilter af;
    lxlsx_reader_workbook_open(LXLSX_TEST_PHASE3_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(1, lxlsx_reader_worksheet_autofilter(ws, &af));
    TEST_ASSERT_EQUAL_STRING("A1:D8", af.range);
    TEST_ASSERT_EQUAL_size_t(2, af.columns_count);

    /* Column 0: list filter A,B */
    TEST_ASSERT_EQUAL_INT(0, af.columns[0].col_id);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FILTER_LIST, af.columns[0].kind);
    TEST_ASSERT_NOT_NULL(af.columns[0].values);
    TEST_ASSERT_EQUAL_STRING("A", af.columns[0].values[0]);
    TEST_ASSERT_EQUAL_STRING("B", af.columns[0].values[1]);
    TEST_ASSERT_NULL(af.columns[0].values[2]);

    /* Column 3: custom AND with two predicates */
    TEST_ASSERT_EQUAL_INT(3, af.columns[1].col_id);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_FILTER_CUSTOM, af.columns[1].kind);
    TEST_ASSERT_EQUAL_INT(1, af.columns[1].custom_and);
    TEST_ASSERT_EQUAL_STRING("greaterThan", af.columns[1].custom_op_1);
    TEST_ASSERT_EQUAL_STRING("20",          af.columns[1].custom_val_1);
    TEST_ASSERT_EQUAL_STRING("lessThan",    af.columns[1].custom_op_2);
    TEST_ASSERT_EQUAL_STRING("60",          af.columns[1].custom_val_2);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_formula_array);
    RUN_TEST(test_formula_shared);
    RUN_TEST(test_data_validations);
    RUN_TEST(test_autofilter);
    return UNITY_END();
}
