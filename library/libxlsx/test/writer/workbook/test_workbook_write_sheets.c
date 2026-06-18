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

// Test the _write_sheets() function.
CTEST(workbook, write_sheets) {


    char* got;
    char exp[] = "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    lxlsx_workbook_add_worksheet(workbook, NULL);

    _write_sheets(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}

