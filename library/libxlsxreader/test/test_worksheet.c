#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "xlsxreader.h"

void setUp(void) {}
void tearDown(void) {}

/* ------------------------------------------------------------------------- */
/* Basic flow against hidden_row.xlsx                                        */
/* ------------------------------------------------------------------------- */

static void test_open_by_index_and_close(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb));
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws));
    TEST_ASSERT_NOT_NULL(ws);
    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_iterate_all_rows(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    int rows = 0;

    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    while (lxr_worksheet_next_row(ws) == LXR_NO_ERROR) {
        lxr_cell c;
        rows++;
        TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_cell(ws, &c));
        TEST_ASSERT_EQUAL_INT(LXR_CELL_STRING, c.type);
        TEST_ASSERT_EQUAL_INT(rows, (int)c.row);
        TEST_ASSERT_EQUAL_INT(1, (int)c.col);
        /* No more cells in this row */
        TEST_ASSERT_EQUAL_INT(LXR_ERROR_END_OF_DATA,
            lxr_worksheet_next_cell(ws, &c));
    }
    TEST_ASSERT_EQUAL_INT(4, rows);

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_skip_hidden_rows(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    int rows = 0;

    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_HIDDEN_ROWS, &ws);

    while (lxr_worksheet_next_row(ws) == LXR_NO_ERROR) {
        lxr_cell c;
        rows++;
        lxr_worksheet_next_cell(ws, &c);
        /* Drain row */
        while (lxr_worksheet_next_cell(ws, &c) == LXR_NO_ERROR) ;
    }
    TEST_ASSERT_EQUAL_INT(2, rows);  /* rows 2 & 3 are hidden */
    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_skip_rows_budget(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    int rows = 0;

    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);
    lxr_worksheet_skip_rows(ws, 2);

    while (lxr_worksheet_next_row(ws) == LXR_NO_ERROR) rows++;
    TEST_ASSERT_EQUAL_INT(2, rows);  /* 4 total - 2 skipped */

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

/* ------------------------------------------------------------------------- */
/* Callback mode                                                             */
/* ------------------------------------------------------------------------- */

typedef struct {
    int cells;
    int rows;
} cb_state;

static int cell_cb(const lxr_cell *c, void *ud)
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
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    cb_state s = {0, 0};

    lxr_workbook_open(LXR_TEST_HIDDEN_ROW_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_worksheet_process(ws, cell_cb, row_cb, &s));
    TEST_ASSERT_EQUAL_INT(4, s.cells);
    TEST_ASSERT_EQUAL_INT(4, s.rows);
    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
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

static int type_cell_cb(const lxr_cell *c, void *ud)
{
    type_state *s = (type_state *)ud;
    s->cells_seen++;
    switch (c->type) {
    case LXR_CELL_NUMBER:   s->n_number++;   break;
    case LXR_CELL_STRING:
        s->n_string++;
        if (c->row == 2 && c->col == 1) {
            size_t n = c->value.string.len;
            if (n >= sizeof(s->first_string)) n = sizeof(s->first_string) - 1;
            memcpy(s->first_string, c->value.string.ptr, n);
            s->first_string[n] = 0;
        }
        if (c->row == 10 && c->col == 1) {
            size_t n = c->value.string.len;
            if (n >= sizeof(s->rich_text)) n = sizeof(s->rich_text) - 1;
            memcpy(s->rich_text, c->value.string.ptr, n);
            s->rich_text[n] = 0;
        }
        break;
    case LXR_CELL_BOOLEAN:
        s->n_boolean++;
        if (c->value.boolean) s->bool_true_seen = 1;
        else                  s->bool_false_seen = 1;
        break;
    case LXR_CELL_ERROR:
        s->n_error++;
        strncpy(s->err_code, c->value.error_code, sizeof(s->err_code) - 1);
        break;
    case LXR_CELL_FORMULA:   s->n_formula++;  break;
    case LXR_CELL_INLINE_STRING: s->n_inline++; break;
    case LXR_CELL_DATETIME:  s->n_datetime++; break;
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
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    type_state s;
    memset(&s, 0, sizeof(s));

    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_open(LXR_TEST_TYPES_XLSX, &wb));
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_get_worksheet_by_name(wb, "Types", LXR_SKIP_NONE, &ws));

    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_worksheet_process(ws, type_cell_cb, type_row_cb, &s));

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

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_pull_pulls_correct_values(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_cell c;
    lxr_workbook_open(LXR_TEST_TYPES_XLSX, &wb);
    lxr_workbook_get_worksheet_by_name(wb, "Types", LXR_SKIP_NONE, &ws);

    /* Row 1: 42, 3.14 */
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXR_CELL_NUMBER, c.type);
    TEST_ASSERT_EQUAL_DOUBLE(42.0, c.value.number);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXR_CELL_NUMBER, c.type);
    TEST_ASSERT_EQUAL_DOUBLE(3.14, c.value.number);
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_END_OF_DATA, lxr_worksheet_next_cell(ws, &c));

    /* Row 2: shared strings */
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXR_CELL_STRING, c.type);
    TEST_ASSERT_EQUAL_STRING_LEN("hello", c.value.string.ptr, 5);

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
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
    return UNITY_END();
}
