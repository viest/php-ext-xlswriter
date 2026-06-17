/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/workbook.h"

// Test the _write_workbook_view() function.
CTEST(workbook, write_workbook_view1) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function.
CTEST(workbook, write_workbook_view2) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\" activeTab=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;
    workbook->active_sheet = 1;

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function.
CTEST(workbook, write_workbook_view3) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\" firstSheet=\"2\" activeTab=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;
    workbook->active_sheet = 1;
    workbook->first_sheet = 2;

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function with set_size().
CTEST(workbook, write_workbook_view4) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_set_size(workbook, 0, 0);

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function with set_size().
CTEST(workbook, write_workbook_view5) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_set_size(workbook, 1073, 644);


    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function with set_size().
CTEST(workbook, write_workbook_view6) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"1845\" windowHeight=\"1050\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_set_size(workbook, 123, 70);

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

// Test the _write_workbook_view() function with set_size().
CTEST(workbook, write_workbook_view7) {

    char* got;
    char exp[] = "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"10785\" windowHeight=\"7350\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_set_size(workbook, 719, 490);

    _write_workbook_view(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}
