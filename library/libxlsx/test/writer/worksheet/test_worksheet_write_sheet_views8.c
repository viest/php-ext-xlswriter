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

/* 1. Test the _write_sheet_views() method with split panes + selection. */
CTEST(worksheet, set_selection31) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"600\" topLeftCell=\"A2\" activePane=\"bottomLeft\"/><selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_set_selection(worksheet, RANGE("A2:A2"));
    lxlsx_worksheet_split_panes(worksheet, 15, 0);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 2. Test the _write_sheet_views() method with split panes + selection. */
CTEST(worksheet, set_selection32) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"1350\" topLeftCell=\"B1\" activePane=\"topRight\"/><selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_set_selection(worksheet, RANGE("B1:B1"));
    lxlsx_worksheet_split_panes(worksheet, 0, 8.43);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 3. Test the _write_sheet_views() method with split panes + selection. */
CTEST(worksheet, set_selection33) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"6150\" ySplit=\"1200\" topLeftCell=\"G4\" activePane=\"bottomRight\"/><selection pane=\"topRight\" activeCell=\"G1\" sqref=\"G1\"/><selection pane=\"bottomLeft\" activeCell=\"A4\" sqref=\"A4\"/><selection pane=\"bottomRight\" activeCell=\"G4\" sqref=\"G4\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_set_selection(worksheet, RANGE("G4:G4"));
    lxlsx_worksheet_split_panes(worksheet, 45, 54.14);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* 4. Test the _write_sheet_views() method with split panes + selection. */
CTEST(worksheet, set_selection34) {
    char* got;
    char exp[] = "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"6150\" ySplit=\"1200\" topLeftCell=\"G4\" activePane=\"bottomRight\"/><selection pane=\"topRight\" activeCell=\"G1\" sqref=\"G1\"/><selection pane=\"bottomLeft\" activeCell=\"A4\" sqref=\"A4\"/><selection pane=\"bottomRight\" activeCell=\"I5\" sqref=\"I5\"/></sheetView></sheetViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_select(worksheet);
    lxlsx_worksheet_set_selection(worksheet, RANGE("I5:I5"));
    lxlsx_worksheet_split_panes(worksheet, 45, 54.14);
    _worksheet_write_sheet_views(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}
