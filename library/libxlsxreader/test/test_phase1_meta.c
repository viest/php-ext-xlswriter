/* Unity tests for Phase 1 metadata: visibility, defined names, merged
 * cells, hyperlinks, sheet protection, row/col options, default sizes. */

#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "xlsxreader.h"

void setUp(void) {}
void tearDown(void) {}

static void test_sheet_visibility(void)
{
    lxr_workbook *wb = NULL;
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb));
    TEST_ASSERT_EQUAL_size_t(2, lxr_workbook_sheet_count(wb));
    TEST_ASSERT_EQUAL_INT(LXR_SHEET_VISIBLE,
        lxr_workbook_sheet_visibility(wb, 0));
    TEST_ASSERT_EQUAL_INT(LXR_SHEET_HIDDEN,
        lxr_workbook_sheet_visibility(wb, 1));
    lxr_workbook_close(wb);
}

static void test_defined_names(void)
{
    lxr_workbook *wb = NULL;
    lxr_defined_name dn;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    TEST_ASSERT_EQUAL_size_t(3, lxr_workbook_defined_name_count(wb));

    TEST_ASSERT_EQUAL_INT(1, lxr_workbook_defined_name_get(wb, 0, &dn));
    TEST_ASSERT_EQUAL_STRING("GlobalArea", dn.name);
    TEST_ASSERT_EQUAL_STRING("Visible!$A$1:$B$2", dn.formula);
    TEST_ASSERT_EQUAL_INT(-1, dn.scope_sheet_index);
    TEST_ASSERT_EQUAL_INT(0, dn.hidden);

    TEST_ASSERT_EQUAL_INT(1, lxr_workbook_defined_name_get(wb, 1, &dn));
    TEST_ASSERT_EQUAL_STRING("LocalToHidden", dn.name);
    TEST_ASSERT_EQUAL_INT(1, dn.scope_sheet_index);

    TEST_ASSERT_EQUAL_INT(1, lxr_workbook_defined_name_get(wb, 2, &dn));
    TEST_ASSERT_EQUAL_STRING("HiddenDef", dn.name);
    TEST_ASSERT_EQUAL_INT(1, dn.hidden);

    lxr_workbook_close(wb);
}

static void test_merged_cells(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_range      r;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR,
        lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws));
    TEST_ASSERT_EQUAL_size_t(2, lxr_worksheet_merged_count(ws));

    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_merged_get(ws, 0, &r));
    TEST_ASSERT_EQUAL_size_t(1, r.first_row);
    TEST_ASSERT_EQUAL_size_t(1, r.first_col);
    TEST_ASSERT_EQUAL_size_t(1, r.last_row);
    TEST_ASSERT_EQUAL_size_t(3, r.last_col);

    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_merged_get(ws, 1, &r));
    TEST_ASSERT_EQUAL_size_t(2, r.first_row);
    TEST_ASSERT_EQUAL_size_t(4, r.first_col);
    TEST_ASSERT_EQUAL_size_t(5, r.last_row);
    TEST_ASSERT_EQUAL_size_t(5, r.last_col);

    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_merged_get(ws, 2, &r));

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_hyperlinks(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_hyperlink  h;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_size_t(2, lxr_worksheet_hyperlink_count(ws));

    /* External link at A5 (row=5,col=1). */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_hyperlink_get(ws, 0, &h));
    TEST_ASSERT_EQUAL_STRING("https://example.com/x", h.url);
    TEST_ASSERT_EQUAL_STRING("Click me", h.tooltip);
    TEST_ASSERT_EQUAL_size_t(5, h.range.first_row);
    TEST_ASSERT_EQUAL_size_t(1, h.range.first_col);

    /* Internal link at B5. */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_hyperlink_get(ws, 1, &h));
    TEST_ASSERT_NULL(h.url);
    TEST_ASSERT_EQUAL_STRING("Hidden!A1", h.location);

    /* Hyperlink-by-cell lookup. */
    TEST_ASSERT_EQUAL_STRING("https://example.com/x",
        lxr_worksheet_hyperlink_url(ws, 5, 1));
    /* Internal location is returned when no external URL. */
    TEST_ASSERT_EQUAL_STRING("Hidden!A1",
        lxr_worksheet_hyperlink_url(ws, 5, 2));
    /* No link at A1 — NULL. */
    TEST_ASSERT_NULL(lxr_worksheet_hyperlink_url(ws, 1, 1));

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_in_merge_follow(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    /* Phase1 fixture has two merges: A1:C1 and D2:E5. */

    /* Master cells -> 0 (not a follow position). */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_in_merge_follow(ws, 1, 1)); /* A1 master */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_in_merge_follow(ws, 2, 4)); /* D2 master */

    /* Follow positions inside merges -> 1. */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_in_merge_follow(ws, 1, 2)); /* B1 inside A1:C1 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_in_merge_follow(ws, 1, 3)); /* C1 inside A1:C1 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_in_merge_follow(ws, 5, 5)); /* E5 inside D2:E5 */

    /* Cells outside any merge -> 0. */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_in_merge_follow(ws, 1, 4)); /* D1 outside */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_in_merge_follow(ws, 6, 1));

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_sheet_protection(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_protection p;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_protection(ws, &p));
    TEST_ASSERT_EQUAL_INT(1, p.is_present);
    TEST_ASSERT_EQUAL_STRING("DAA7", p.password_hash);
    TEST_ASSERT_EQUAL_INT(1, p.sheet);
    TEST_ASSERT_EQUAL_INT(1, p.objects);
    TEST_ASSERT_EQUAL_INT(1, p.scenarios);
    /* formatCells="0" sets the corresponding flag to 0 in our raw model. */
    TEST_ASSERT_EQUAL_INT(0, p.format_cells);

    lxr_worksheet_close(ws);

    /* Sheet 2 has no <sheetProtection>. */
    lxr_workbook_get_worksheet_by_index(wb, 1, LXR_SKIP_NONE, &ws);
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_protection(ws, &p));
    TEST_ASSERT_EQUAL_INT(0, p.is_present);

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_row_options(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_row_options ro;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    /* Row 2: ht=30.5 customHeight=1 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_row_options(ws, 2, &ro));
    TEST_ASSERT_EQUAL_INT(1, ro.has_height);
    TEST_ASSERT_EQUAL_DOUBLE(30.5, ro.height);
    TEST_ASSERT_EQUAL_INT(1, ro.custom_height);

    /* Row 3: hidden */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_row_options(ws, 3, &ro));
    TEST_ASSERT_EQUAL_INT(1, ro.hidden);

    /* Row 4: outlineLevel=2 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_row_options(ws, 4, &ro));
    TEST_ASSERT_EQUAL_INT(2, ro.outline_level);

    /* Row 1 and 5 have no metadata. */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_row_options(ws, 1, &ro));
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_row_options(ws, 5, &ro));

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_col_options(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    lxr_col_options co;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    /* Col B (=2): width 22.5 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_col_options(ws, 2, &co));
    TEST_ASSERT_EQUAL_INT(1, co.has_width);
    TEST_ASSERT_EQUAL_DOUBLE(22.5, co.width);
    TEST_ASSERT_EQUAL_INT(0, co.hidden);

    /* Col C and D (3..4) share width=11 hidden=1 */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_col_options(ws, 3, &co));
    TEST_ASSERT_EQUAL_INT(1, co.hidden);
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_col_options(ws, 4, &co));
    TEST_ASSERT_EQUAL_INT(1, co.hidden);

    /* Col E (=5): outline 2. */
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_col_options(ws, 5, &co));
    TEST_ASSERT_EQUAL_INT(2, co.outline_level);

    /* Col A (=1) has no entry. */
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_col_options(ws, 1, &co));

    lxr_worksheet_close(ws);
    lxr_workbook_close(wb);
}

static void test_defaults(void)
{
    lxr_workbook  *wb = NULL;
    lxr_worksheet *ws = NULL;
    double         d = 0.0;
    lxr_workbook_open(LXR_TEST_PHASE1_XLSX, &wb);
    lxr_workbook_get_worksheet_by_index(wb, 0, LXR_SKIP_NONE, &ws);

    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_default_row_height(ws, &d));
    TEST_ASSERT_EQUAL_DOUBLE(18.5, d);
    TEST_ASSERT_EQUAL_INT(1, lxr_worksheet_default_col_width(ws, &d));
    TEST_ASSERT_EQUAL_DOUBLE(9.25, d);

    lxr_worksheet_close(ws);

    /* Sheet 2 has no sheetFormatPr. */
    lxr_workbook_get_worksheet_by_index(wb, 1, LXR_SKIP_NONE, &ws);
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_default_row_height(ws, &d));
    TEST_ASSERT_EQUAL_INT(0, lxr_worksheet_default_col_width(ws, &d));
    lxr_worksheet_close(ws);

    lxr_workbook_close(wb);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_sheet_visibility);
    RUN_TEST(test_defined_names);
    RUN_TEST(test_merged_cells);
    RUN_TEST(test_hyperlinks);
    RUN_TEST(test_in_merge_follow);
    RUN_TEST(test_sheet_protection);
    RUN_TEST(test_row_options);
    RUN_TEST(test_col_options);
    RUN_TEST(test_defaults);
    return UNITY_END();
}
