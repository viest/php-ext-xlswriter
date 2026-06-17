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

// Test the _write_workbook_pr() function.
CTEST(workbook, write_workbook_pr) {

    char* got;
    char exp[] = "<workbookPr defaultThemeVersion=\"124226\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    _write_workbook_pr(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxw_workbook_free(workbook);
}

