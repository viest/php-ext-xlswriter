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

// Test the _write_dimension() function.
CTEST(worksheet, write_dimension) {

    char* got;
    char exp[] = "<dimension ref=\"A1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_dimension(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}
