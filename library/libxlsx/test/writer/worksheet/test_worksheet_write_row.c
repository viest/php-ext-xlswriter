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

// Test the _write_row() function.
CTEST(worksheet, write_row) {

    char* got;
    char exp[] = "<row r=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_row *row = _get_row_list(worksheet->table, 0);

    _write_row(worksheet, row, NULL);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}
