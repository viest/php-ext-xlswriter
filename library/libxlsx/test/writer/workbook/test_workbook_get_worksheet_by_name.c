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

/* Test finding a worksheet that does exist (implicit naming). */
CTEST(workbook, get_worksheet_by_name01) {
    lxw_workbook *workbook = workbook_new(NULL);

    lxw_worksheet *exp = workbook_add_worksheet(workbook, NULL);
    lxw_worksheet *got = workbook_get_worksheet_by_name(workbook, "Sheet1");

    ASSERT_TRUE(got == exp);

    lxw_workbook_free(workbook);
}

/* Test finding a worksheet that does exist (explicit naming). */
CTEST(workbook, get_worksheet_by_name02) {

    lxw_workbook *workbook = workbook_new(NULL);

    lxw_worksheet *exp = workbook_add_worksheet(workbook, "FOO");
    lxw_worksheet *got = workbook_get_worksheet_by_name(workbook, "FOO");

    ASSERT_TRUE(got == exp);

    lxw_workbook_free(workbook);
}

/* Test finding a worksheet that doesn't exist. */
CTEST(workbook, get_worksheet_by_name03) {

    lxw_workbook *workbook = workbook_new(NULL);

    workbook_add_worksheet(workbook, NULL);
    lxw_worksheet *got = workbook_get_worksheet_by_name(workbook, "FOO");

    ASSERT_TRUE(got == NULL);

    lxw_workbook_free(workbook);
}

/* Test finding a worksheet when no worksheets exist. */
CTEST(workbook, get_worksheet_by_name04) {

    lxw_workbook *workbook = workbook_new(NULL);

    lxw_worksheet *got = workbook_get_worksheet_by_name(workbook, "FOO");

    ASSERT_TRUE(got == NULL);

    lxw_workbook_free(workbook);
}

/* Test finding a worksheet with a NULL name. */
CTEST(workbook, get_worksheet_by_name05) {

    lxw_workbook *workbook = workbook_new(NULL);

    workbook_add_worksheet(workbook, NULL);
    lxw_worksheet *got = workbook_get_worksheet_by_name(workbook, NULL);

    ASSERT_TRUE(got == NULL);

    lxw_workbook_free(workbook);
}
