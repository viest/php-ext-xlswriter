/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/workbook.h"

// Test the _write_book_views() function.
CTEST(workbook, write_book_views) {


    char* got;
    char exp[] = "<bookViews><workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/></bookViews>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    _write_book_views(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}

