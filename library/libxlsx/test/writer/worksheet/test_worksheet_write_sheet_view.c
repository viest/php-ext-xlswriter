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

// Test the _write_sheet_view() function.
CTEST(worksheet, write_sheet_view1) {

    char* got;
    char exp[] = "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    _worksheet_write_sheet_view(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}

// Test the _write_sheet_view() function.
CTEST(worksheet, write_sheet_view2) {

    char* got;
    char exp[] = "<sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    lxlsx_worksheet_gridlines(worksheet, LXLSX_HIDE_ALL_GRIDLINES);

    _worksheet_write_sheet_view(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}



