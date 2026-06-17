/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/workbook.h"

/* Test the _write_defined_name() method. */
CTEST(workbook, write_defined_name) {
    char* got;
    char exp[] = "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Sheet1!$1:$1</definedName>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    lxlsx_defined_name defined_name = {0, 0, "_xlnm.Print_Titles", "", "Sheet1!$1:$1", "", "", {NULL, NULL}};


    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    _write_defined_name(workbook, &defined_name);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}

