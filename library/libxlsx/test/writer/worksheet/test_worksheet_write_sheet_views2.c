/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/worksheet.h"

/* 1. Test the _write_sheet_views() method with freeze panes. */
CTEST(worksheet, write_freeze_panes01) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_freeze_panes(worksheet, 1, 0);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 2. Test the _write_sheet_views() method with freeze panes. */
CTEST(worksheet, write_freeze_panes02) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_freeze_panes(worksheet, 0, 1);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 3. Test the _write_sheet_views() method with freeze panes. */
CTEST(worksheet, write_freeze_panes03) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/><selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/><selection pane=\"bottomRight\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_freeze_panes(worksheet, 1, 1);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 4. Test the _write_sheet_views() method with freeze panes. */
CTEST(worksheet, write_freeze_panes04) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"6\" ySplit=\"3\" topLeftCell=\"G4\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"topRight\" activeCell=\"G1\" sqref=\"G1\"/><selection pane=\"bottomLeft\" activeCell=\"A4\" sqref=\"A4\"/><selection pane=\"bottomRight\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_freeze_panes(worksheet, CELL("G4"));
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 5. Test the _write_sheet_views() method with freeze panes. */
CTEST(worksheet, write_freeze_panes05) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"6\" ySplit=\"3\" topLeftCell=\"G4\" activePane=\"bottomRight\" state=\"frozenSplit\"/><selection pane=\"topRight\" activeCell=\"G1\" sqref=\"G1\"/><selection pane=\"bottomLeft\" activeCell=\"A4\" sqref=\"A4\"/><selection pane=\"bottomRight\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_freeze_panes_opt(worksheet, 3, 6, 3, 6, 1);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}
