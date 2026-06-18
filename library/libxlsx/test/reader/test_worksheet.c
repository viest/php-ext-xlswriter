#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "xlsx_test_paths.h"
#include "libxlsx.h"
#include "libxlsx/source_package.h"

void setUp(void) {}
void tearDown(void) {}

/* ------------------------------------------------------------------------- */
/* Basic flow against hidden_row.xlsx                                        */
/* ------------------------------------------------------------------------- */

static void test_open_by_index_and_close(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_workbook_open(LXLSX_TEST_HIDDEN_ROW_XLSX, &wb));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws));
    TEST_ASSERT_NOT_NULL(ws);
    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_iterate_all_rows(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    int rows = 0;

    lxlsx_reader_workbook_open(LXLSX_TEST_HIDDEN_ROW_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    while (lxlsx_reader_worksheet_next_row(ws) == LXLSX_READER_NO_ERROR) {
        lxlsx_cell c;
        rows++;
        TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
        TEST_ASSERT_EQUAL_INT(STRING_CELL, c.type);
        TEST_ASSERT_EQUAL_INT(rows, (int)c.row_num);
        TEST_ASSERT_EQUAL_INT(1, (int)c.col_num);
        /* No more cells in this row */
        TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_END_OF_DATA,
            lxlsx_reader_worksheet_next_cell(ws, &c));
    }
    TEST_ASSERT_EQUAL_INT(4, rows);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_skip_hidden_rows(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    int rows = 0;

    lxlsx_reader_workbook_open(LXLSX_TEST_HIDDEN_ROW_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_HIDDEN_ROWS, &ws);

    while (lxlsx_reader_worksheet_next_row(ws) == LXLSX_READER_NO_ERROR) {
        lxlsx_cell c;
        rows++;
        lxlsx_reader_worksheet_next_cell(ws, &c);
        /* Drain row */
        while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    }
    TEST_ASSERT_EQUAL_INT(2, rows);  /* rows 2 & 3 are hidden */
    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_skip_rows_budget(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    int rows = 0;

    lxlsx_reader_workbook_open(LXLSX_TEST_HIDDEN_ROW_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);
    lxlsx_reader_worksheet_skip_rows(ws, 2);

    while (lxlsx_reader_worksheet_next_row(ws) == LXLSX_READER_NO_ERROR) rows++;
    TEST_ASSERT_EQUAL_INT(2, rows);  /* 4 total - 2 skipped */

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

/* ------------------------------------------------------------------------- */
/* Callback mode                                                             */
/* ------------------------------------------------------------------------- */

typedef struct {
    int cells;
    int rows;
} cb_state;

static int cell_cb(const lxlsx_cell *c, void *ud)
{
    cb_state *s = (cb_state *)ud;
    s->cells++;
    (void)c;
    return 0;
}

static int row_cb(size_t row, size_t max_col, void *ud)
{
    cb_state *s = (cb_state *)ud;
    s->rows++;
    (void)row;
    (void)max_col;
    return 0;
}

static void test_callback_mode(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    cb_state s = {0, 0};

    lxlsx_reader_workbook_open(LXLSX_TEST_HIDDEN_ROW_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_worksheet_process(ws, cell_cb, row_cb, &s));
    TEST_ASSERT_EQUAL_INT(4, s.cells);
    TEST_ASSERT_EQUAL_INT(4, s.rows);
    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

/* ------------------------------------------------------------------------- */
/* Comprehensive type fixture                                                */
/* ------------------------------------------------------------------------- */

typedef struct {
    int rows_seen;
    int cells_seen;
    int n_number;
    int n_string;
    int n_boolean;
    int n_error;
    int n_formula;
    int n_inline;
    int n_datetime;
    char first_string[64];
    char err_code[16];
    int  bool_true_seen;
    int  bool_false_seen;
    char rich_text[64];
} type_state;

static int type_cell_cb(const lxlsx_cell *c, void *ud)
{
    type_state *s = (type_state *)ud;
    s->cells_seen++;
    switch (c->type) {
    case NUMBER_CELL:   s->n_number++;   break;
    case STRING_CELL:
        s->n_string++;
        if (c->row_num == 2 && c->col_num == 1) {
            size_t n = c->value.string.len;
            if (n >= sizeof(s->first_string)) n = sizeof(s->first_string) - 1;
            memcpy(s->first_string, c->value.string.ptr, n);
            s->first_string[n] = 0;
        }
        if (c->row_num == 10 && c->col_num == 1) {
            size_t n = c->value.string.len;
            if (n >= sizeof(s->rich_text)) n = sizeof(s->rich_text) - 1;
            memcpy(s->rich_text, c->value.string.ptr, n);
            s->rich_text[n] = 0;
        }
        break;
    case BOOLEAN_CELL:
        s->n_boolean++;
        if (c->value.boolean) s->bool_true_seen = 1;
        else                  s->bool_false_seen = 1;
        break;
    case ERROR_CELL:
        s->n_error++;
        strncpy(s->err_code, c->value.error_code, sizeof(s->err_code) - 1);
        break;
    case FORMULA_CELL:   s->n_formula++;  break;
    case INLINE_STRING_CELL: s->n_inline++; break;
    case DATETIME_CELL:  s->n_datetime++; break;
    default: break;
    }
    return 0;
}

static int type_row_cb(size_t r, size_t mc, void *ud)
{
    type_state *s = (type_state *)ud;
    (void)r; (void)mc;
    s->rows_seen++;
    return 0;
}

static void test_all_cell_types(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    type_state s;
    memset(&s, 0, sizeof(s));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_workbook_open(LXLSX_TEST_TYPES_XLSX, &wb));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_workbook_get_worksheet_by_name(wb, "Types", LXLSX_READER_SKIP_NONE, &ws));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_worksheet_process(ws, type_cell_cb, type_row_cb, &s));

    TEST_ASSERT_EQUAL_INT(10, s.rows_seen);
    /* Row 1 has 2 plain numbers (42, 3.14); row 9 percent (xf[4], numFmtId=9)
     * is classified PERCENT — emit_cell treats everything non-date as NUMBER,
     * so it counts here. Total: 3 numbers. */
    TEST_ASSERT_EQUAL_INT_MESSAGE(3, s.n_number,   "row 1 numbers + row 9 percent");
    TEST_ASSERT_EQUAL_INT_MESSAGE(3, s.n_string,   "expect SST + SST + rich");
    TEST_ASSERT_EQUAL_INT_MESSAGE(2, s.n_boolean,  "expect true and false");
    TEST_ASSERT_EQUAL_INT_MESSAGE(1, s.n_error,    "#DIV/0!");
    TEST_ASSERT_EQUAL_INT_MESSAGE(1, s.n_formula,  "SUM formula");
    TEST_ASSERT_EQUAL_INT_MESSAGE(1, s.n_inline,   "inline string");
    TEST_ASSERT_EQUAL_INT_MESSAGE(3, s.n_datetime, "rows 6,7,8 are dates by style");
    /* row 9 percent: numFmtId 9 -> PERCENT, classified as NUMBER (not date) */
    TEST_ASSERT_EQUAL_STRING("hello",          s.first_string);
    TEST_ASSERT_EQUAL_STRING("rich-text run",  s.rich_text);
    TEST_ASSERT_EQUAL_STRING("#DIV/0!",        s.err_code);
    TEST_ASSERT_TRUE(s.bool_true_seen);
    TEST_ASSERT_TRUE(s.bool_false_seen);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_pull_pulls_correct_values(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_cell c;
    lxlsx_reader_workbook_open(LXLSX_TEST_TYPES_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_name(wb, "Types", LXLSX_READER_SKIP_NONE, &ws);

    /* Row 1: 42, 3.14 */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(NUMBER_CELL, c.type);
    TEST_ASSERT_EQUAL_DOUBLE(42.0, c.value.number);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(NUMBER_CELL, c.type);
    TEST_ASSERT_EQUAL_DOUBLE(3.14, c.value.number);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_END_OF_DATA, lxlsx_reader_worksheet_next_cell(ws, &c));

    /* Row 2: shared strings */
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(STRING_CELL, c.type);
    TEST_ASSERT_EQUAL_STRING_LEN("hello", c.value.string.ptr, 5);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_unified_cell_formula_and_style_ref(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_cell c;

    lxlsx_reader_workbook_open(LXLSX_TEST_TYPES_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_name(wb, "Types", LXLSX_READER_SKIP_NONE, &ws);

    lxlsx_reader_worksheet_skip_rows(ws, 4);
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(FORMULA_CELL, c.type);
    TEST_ASSERT_EQUAL_STRING_LEN("SUM(A1:A2)",
                                 c.value.formula.formula.ptr,
                                 c.value.formula.formula.len);
    TEST_ASSERT_EQUAL_STRING_LEN("44",
                                 c.value.formula.cached.ptr,
                                 c.value.formula.cached.len);

    while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    while (lxlsx_reader_worksheet_next_cell(ws, &c) == LXLSX_READER_NO_ERROR) ;
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(NUMBER_CELL, c.type);
    TEST_ASSERT_EQUAL_INT(4, (int)c.style_ref);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_invalid_cell_ref_is_rejected(void)
{
    const char *path = "fixtures/invalid_cell_ref.xlsx";
    lxlsx_source_package *package = NULL;
    unsigned char *sheet_xml = NULL;
    size_t sheet_xml_len = 0;
    int sheet_index;
    char *cell_ref;
    lxlsx_source_package_replacement replacement;
    lxlsx_reader_workbook *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_cell cell;

    remove(path);
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR,
                          lxlsx_source_package_open(LXLSX_TEST_TYPES_XLSX,
                                                     &package));
    sheet_index = lxlsx_source_package_find_first(package,
                                                  "xl/worksheets/sheet1.xml");
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, sheet_index);
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR,
                          lxlsx_source_package_read_entry(
                              package, (size_t)sheet_index, &sheet_xml,
                              &sheet_xml_len));

    cell_ref = strstr((char *)sheet_xml, "r=\"A1\"");
    TEST_ASSERT_NOT_NULL(cell_ref);
    cell_ref[3] = '1';
    cell_ref[4] = 'A';

    replacement.entry_index = (size_t)sheet_index;
    replacement.data = sheet_xml;
    replacement.size = sheet_xml_len;
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR,
                          lxlsx_source_package_save_with_replacements(
                              package, path, &replacement, 1));

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_open(path, &wb));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_workbook_get_worksheet_by_name(
                              wb, "Types", LXLSX_READER_SKIP_NONE, &ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
                          lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_ERROR_INVALID_CELL_REF,
                          lxlsx_reader_worksheet_next_cell(ws, &cell));

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
    lxlsx_source_package_free_buffer(sheet_xml);
    lxlsx_source_package_close(package);
    remove(path);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_open_by_index_and_close);
    RUN_TEST(test_iterate_all_rows);
    RUN_TEST(test_skip_hidden_rows);
    RUN_TEST(test_skip_rows_budget);
    RUN_TEST(test_callback_mode);
    RUN_TEST(test_all_cell_types);
    RUN_TEST(test_pull_pulls_correct_values);
    RUN_TEST(test_unified_cell_formula_and_style_ref);
    RUN_TEST(test_invalid_cell_ref_is_rejected);
    return UNITY_END();
}
