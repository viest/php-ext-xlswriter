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


/* 1. Test the _write_sheet_pr() method. */
CTEST(worksheet, write_write_sheet_pr01) {
    char* got;
    char exp[] = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet->fit_page = 1;

    _worksheet_write_sheet_pr(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 2. Test the _write_sheet_pr() method. */
CTEST(worksheet, write_write_sheet_pr02) {
    char* got;
    char exp[] = "<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_set_tab_color(worksheet, LXW_COLOR_RED);
    _worksheet_write_sheet_pr(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 3. Test the _write_sheet_pr() method. */
CTEST(worksheet, write_write_sheet_pr03) {
    char* got;
    char exp[] = "<sheetPr><tabColor rgb=\"FFFF0000\"/><pageSetUpPr fitToPage=\"1\"/></sheetPr>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet->fit_page = 1;

    worksheet_set_tab_color(worksheet, LXW_COLOR_RED);
    _worksheet_write_sheet_pr(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


