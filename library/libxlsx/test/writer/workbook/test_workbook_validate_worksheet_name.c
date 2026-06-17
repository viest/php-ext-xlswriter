/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/workbook.h"
#include "../../../include/lxlsx/shared_strings.h"


/* Test a valid sheet name. */
CTEST(workbook, validate_worksheet_name01) {

    const char* sheetname = "123456789_123456789_123456789_1";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_NO_ERROR;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that is too long. */
CTEST(workbook, validate_worksheet_name02) {

    const char* sheetname = "123456789_123456789_123456789_12";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_LENGTH_EXCEEDED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name contains invalid characters. */
CTEST(workbook, validate_worksheet_name03) {

    const char* sheetname = "Sheet[1]";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_INVALID_SHEETNAME_CHARACTER;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that already exists. */
CTEST(workbook, validate_worksheet_name04) {

    const char* sheetname = "Sheet1";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_workbook_add_worksheet(workbook, sheetname);

    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_ALREADY_USED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that starts with an apostrophe. */
CTEST(workbook, validate_worksheet_name05) {

    const char* sheetname = "'Sheet1";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_START_END_APOSTROPHE;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that ends with an apostrophe. */
CTEST(workbook, validate_worksheet_name06) {

    const char* sheetname = "Sheet1'";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_START_END_APOSTROPHE;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
CTEST(workbook, validate_worksheet_name07) {

    const char* sheetname = "Sheet1";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_workbook_add_worksheet(workbook, sheetname);

    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_ALREADY_USED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, "sheet1");

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
CTEST(workbook, validate_worksheet_name08) {

    const char* sheetname = "Café";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_workbook_add_worksheet(workbook, sheetname);

    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_ALREADY_USED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, "café");

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
CTEST(workbook, validate_worksheet_name09) {

    const char* sheetname = "abcde";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_workbook_add_worksheet(workbook, sheetname);

    lxlsx_error exp = LXLSX_ERROR_SHEETNAME_ALREADY_USED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, "ABCDE");

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test for empty sheet name. */
CTEST(workbook, validate_worksheet_name10) {

    const char* sheetname = "";

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_PARAMETER_IS_EMPTY;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test for NULL sheet name. */
CTEST(workbook, validate_worksheet_name11) {

    const char* sheetname = NULL;

    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    lxlsx_error exp = LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    lxlsx_error got = lxlsx_workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxlsx_workbook_free(workbook);
}
