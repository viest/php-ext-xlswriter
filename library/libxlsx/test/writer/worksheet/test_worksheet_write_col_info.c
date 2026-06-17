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
#include "../../../include/lxlsx/format.h"

// Test the _write_col_info() function.
CTEST(worksheet, write_col_info01) {

    char* got;
    char exp[] = "<col min=\"2\" max=\"4\" width=\"5.7109375\" customWidth=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_col_options col_options = {1, 3, 5, NULL, 0, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


CTEST(worksheet, write_col_info02) {

    char* got;
    char exp[] = "<col min=\"6\" max=\"6\" width=\"8.7109375\" hidden=\"1\" customWidth=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_col_options col_options = {5, 5, 8, NULL, 1, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


CTEST(worksheet, write_col_info03) {

    char* got;
    char exp[] = "<col min=\"8\" max=\"8\" width=\"9.140625\" style=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_format *format = lxw_format_new();
    format->xf_index = 1;

    lxw_col_options col_options = {7, 7, LXW_DEF_COL_WIDTH, format, 0, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


CTEST(worksheet, write_col_info04) {

    char* got;
    char exp[] = "<col min=\"9\" max=\"9\" width=\"9.140625\" style=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_format *format = lxw_format_new();
    format->xf_index = 1;

    lxw_col_options col_options = {8, 8, LXW_DEF_COL_WIDTH, format, 0, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


CTEST(worksheet, write_col_info05) {

    char* got;
    char exp[] = "<col min=\"10\" max=\"10\" width=\"2.7109375\" customWidth=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_col_options col_options = {9, 9, 2, NULL, 0, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


CTEST(worksheet, write_col_info06) {

    char* got;
    char exp[] = "<col min=\"12\" max=\"12\" width=\"0\" hidden=\"1\" customWidth=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_col_options col_options = {11, 11, LXW_DEF_COL_WIDTH, NULL, 1, 0, 0};

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    _worksheet_write_col_info(worksheet, &col_options);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}
