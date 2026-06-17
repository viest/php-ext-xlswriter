/* Unity tests for Phase 4: page setup, rich runs, comments (legacy + threaded). */

#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxlsx_reader_test_paths.h"
#include "lxlsx/reader.h"

void setUp(void) {}
void tearDown(void) {}

static void test_page_setup(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_reader_page_setup p;
    lxlsx_reader_workbook_open(LXLSX_READER_TEST_PHASE4_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(1, lxlsx_reader_worksheet_page_setup(ws, &p));
    TEST_ASSERT_EQUAL_INT(1, p.has_margins);
    TEST_ASSERT_EQUAL_DOUBLE(0.7, p.margin_left);
    TEST_ASSERT_EQUAL_DOUBLE(0.75, p.margin_top);
    TEST_ASSERT_EQUAL_INT(1, p.has_setup);
    TEST_ASSERT_EQUAL_INT(9, p.paper_size);
    TEST_ASSERT_EQUAL_INT(125, p.scale);
    TEST_ASSERT_EQUAL_INT(1, p.orientation_landscape);
    TEST_ASSERT_EQUAL_STRING("&COdd Header",  p.odd_header);
    TEST_ASSERT_EQUAL_STRING("&CEven Header", p.even_header);
    TEST_ASSERT_EQUAL_INT(1, p.different_odd_even);
    TEST_ASSERT_EQUAL_INT(1, p.print_horizontal_centered);
    TEST_ASSERT_EQUAL_INT(1, p.print_grid_lines);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

static void test_rich_sst(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    lxlsx_reader_cell c;
    size_t   n;
    lxlsx_reader_string_run runs[8];
    lxlsx_reader_workbook_open(LXLSX_READER_TEST_PHASE4_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_row(ws));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR, lxlsx_reader_worksheet_next_cell(ws, &c));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_CELL_STRING, c.type);

    n = lxlsx_reader_cell_string_runs(ws, &c, runs, 8);
    TEST_ASSERT_EQUAL_size_t(2, n);
    TEST_ASSERT_EQUAL_STRING("Hello ", runs[0].text);
    TEST_ASSERT_EQUAL_INT(1,  runs[0].bold);
    TEST_ASSERT_EQUAL_INT(0,  runs[0].italic);
    TEST_ASSERT_EQUAL_DOUBLE(12.0, runs[0].font_size);
    TEST_ASSERT_EQUAL_STRING("FFFF0000", runs[0].color);
    TEST_ASSERT_EQUAL_STRING("Arial", runs[0].font_name);

    TEST_ASSERT_EQUAL_STRING("world", runs[1].text);
    TEST_ASSERT_EQUAL_INT(0,  runs[1].bold);
    TEST_ASSERT_EQUAL_INT(1,  runs[1].italic);
    TEST_ASSERT_EQUAL_DOUBLE(14.0, runs[1].font_size);

    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

typedef struct {
    int count;
    lxlsx_reader_comment_info collected[8];
    char  *texts[8];
    char  *authors[8];
    char  *parents[8];
} cb_collect;

static int comment_collector(const lxlsx_reader_comment_info *info, void *ud)
{
    cb_collect *c = (cb_collect *)ud;
    if (c->count >= 8) return 1;
    c->collected[c->count] = *info;
    c->texts[c->count]   = info->text      ? strdup(info->text)      : NULL;
    c->authors[c->count] = info->author    ? strdup(info->author)    : NULL;
    c->parents[c->count] = info->parent_id ? strdup(info->parent_id) : NULL;
    c->count++;
    return 0;
}

static void test_comments(void)
{
    lxlsx_reader_workbook  *wb = NULL;
    lxlsx_reader_worksheet *ws = NULL;
    cb_collect     col;
    int            i;

    lxlsx_reader_workbook_open(LXLSX_READER_TEST_PHASE4_XLSX, &wb);
    lxlsx_reader_workbook_get_worksheet_by_index(wb, 0, LXLSX_READER_SKIP_NONE, &ws);

    memset(&col, 0, sizeof(col));
    TEST_ASSERT_EQUAL_INT(LXLSX_READER_NO_ERROR,
        lxlsx_reader_worksheet_iterate_comments(ws, comment_collector, &col));

    TEST_ASSERT_EQUAL_INT(4, col.count);

    /* Legacy 1: A1 from Eve, visible. */
    TEST_ASSERT_EQUAL_size_t(1, col.collected[0].row);
    TEST_ASSERT_EQUAL_size_t(1, col.collected[0].col);
    TEST_ASSERT_EQUAL_STRING("Eve",                 col.authors[0]);
    TEST_ASSERT_EQUAL_STRING("From Eve, visible",   col.texts[0]);
    TEST_ASSERT_EQUAL_INT(1, col.collected[0].visible);
    TEST_ASSERT_EQUAL_INT(0, col.collected[0].threaded);

    /* Legacy 2: C2 from Bob, hidden. */
    TEST_ASSERT_EQUAL_size_t(2, col.collected[1].row);
    TEST_ASSERT_EQUAL_size_t(3, col.collected[1].col);
    TEST_ASSERT_EQUAL_INT(0, col.collected[1].visible);

    /* Threaded 1: parent. */
    TEST_ASSERT_EQUAL_size_t(3, col.collected[2].row);
    TEST_ASSERT_EQUAL_INT(1,    col.collected[2].threaded);
    TEST_ASSERT_NULL(col.parents[2]);
    TEST_ASSERT_EQUAL_STRING("Top-level threaded comment", col.texts[2]);

    /* Threaded 2: reply, parent_id set. */
    TEST_ASSERT_EQUAL_INT(1, col.collected[3].threaded);
    TEST_ASSERT_EQUAL_STRING("{aaaaaaaa}", col.parents[3]);
    TEST_ASSERT_EQUAL_STRING("Reply to it", col.texts[3]);

    for (i = 0; i < col.count; i++) {
        free(col.texts[i]); free(col.authors[i]); free(col.parents[i]);
    }
    lxlsx_reader_worksheet_close(ws);
    lxlsx_reader_workbook_close(wb);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_page_setup);
    RUN_TEST(test_rich_sst);
    RUN_TEST(test_comments);
    return UNITY_END();
}
