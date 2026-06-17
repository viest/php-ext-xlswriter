/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/worksheet.h"


/* 1. Test the _write_sheet_views() method. */
CTEST(worksheet, write_sheet_views01) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 2. Test the _write_sheet_views() method. */
CTEST(worksheet, write_sheet_views02) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    worksheet_set_zoom(worksheet, 100);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 3. Test the _write_sheet_views() method. With zoom. */
CTEST(worksheet, write_sheet_views03) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" zoomScale=\"200\" zoomScaleNormal=\"200\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    worksheet_set_zoom(worksheet, 200);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 4. Test the _write_sheet_views() method. Right to left. */
CTEST(worksheet, write_sheet_views04) {
    char* got;
    char exp[] = "<sheetViews><sheetView rightToLeft=\"1\" tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    worksheet_right_to_left(worksheet);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 5. Test the _write_sheet_views() method. Hide zeroes. */
CTEST(worksheet, write_sheet_views05) {
    char* got;
    char exp[] = "<sheetViews><sheetView showZeros=\"0\" tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    worksheet_hide_zero(worksheet);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 6. Test the _write_sheet_views() method. Set page view mode. */
CTEST(worksheet, write_sheet_views06) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" view=\"pageLayout\" workbookViewId=\"0\"/></sheetViews>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_select(worksheet);
    worksheet_set_page_view(worksheet);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}
