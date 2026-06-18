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


/* Test the _write_page_setup() method. Without any page setup. */
CTEST(worksheet, write_page_setup01) {
    char* got;
    char exp[] = "";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_page_setup(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* Test the _write_page_setup() method. With set_landscape(); */
CTEST(worksheet, write_page_setup02) {
    char* got;
    char exp[] = "<pageSetup orientation=\"landscape\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_landscape(worksheet);
    _worksheet_write_page_setup(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* Test the _write_page_setup() method. With set_portrait(); */
CTEST(worksheet, write_page_setup03) {
    char* got;
    char exp[] = "<pageSetup orientation=\"portrait\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_portrait(worksheet);
    _worksheet_write_page_setup(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* Test the _write_page_setup() method. With set_paper(); */
CTEST(worksheet, write_page_setup04) {
    char* got;
    char exp[] = "<pageSetup paperSize=\"9\" orientation=\"portrait\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_paper(worksheet, 9);
    _worksheet_write_page_setup(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


/* Test the _write_page_setup() method. With print_across(); */
CTEST(worksheet, write_page_setup05) {
    char* got;
    char exp[] = "<pageSetup pageOrder=\"overThenDown\" orientation=\"portrait\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_print_across(worksheet);
    _worksheet_write_page_setup(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}
