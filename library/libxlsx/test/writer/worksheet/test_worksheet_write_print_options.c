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

// Test the _write_print_options() function.
CTEST(worksheet, write_print_options1) {

    char* got;
    char exp[] = "<printOptions horizontalCentered=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_center_horizontally(worksheet);

    _worksheet_write_print_options(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}

// Test the _write_print_options() function.
CTEST(worksheet, write_print_options2) {

    char* got;
    char exp[] = "<printOptions verticalCentered=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_center_vertically(worksheet);

    _worksheet_write_print_options(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}

// Test the _write_print_options() function.
CTEST(worksheet, write_print_options3) {

    char* got;
    char exp[] = "<printOptions horizontalCentered=\"1\" verticalCentered=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_center_horizontally(worksheet);
    worksheet_center_vertically(worksheet);

    _worksheet_write_print_options(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


// Test the _write_print_options() function.
CTEST(worksheet, write_print_options4) {

    char* got;
    char exp[] = "<printOptions gridLines=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_gridlines(worksheet, LXW_SHOW_PRINT_GRIDLINES);

    _worksheet_write_print_options(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}

