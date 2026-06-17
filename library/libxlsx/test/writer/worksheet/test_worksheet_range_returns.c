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


//  Test the return value for various worksheet functions that handle 1 or 2D ranges.
CTEST(worksheet, bound_checks01) {

    int err;
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_row_t MAX_ROW = 1048576;
    lxlsx_col_t MAX_COL = 16384;

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    lxlsx_format *format = lxlsx_format_new();

    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    err = lxlsx_worksheet_write_number(worksheet, 0, MAX_COL, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_number(worksheet, MAX_ROW, 0, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_number(worksheet, MAX_ROW, MAX_COL, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_string(worksheet, MAX_ROW, 0, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_string(worksheet, 0, MAX_COL, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_string(worksheet, MAX_ROW, MAX_COL, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_number(worksheet, MAX_ROW, 0, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_number(worksheet, 0, MAX_COL, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_number(worksheet, MAX_ROW, MAX_COL, 123, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_blank(worksheet, MAX_ROW, 0, format);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_blank(worksheet, 0, MAX_COL, format);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_blank(worksheet, MAX_ROW, MAX_COL, format);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_formula(worksheet, MAX_ROW, 0, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_formula(worksheet, 0, MAX_COL, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_formula(worksheet, MAX_ROW, MAX_COL, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, 0, MAX_COL, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, MAX_ROW, 0, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, MAX_COL, 0, 0, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, MAX_ROW, 0, 0, 0, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, MAX_ROW, MAX_COL, MAX_ROW, MAX_COL, "=A1", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_merge_range(worksheet, 0, 0, 0, MAX_COL, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_merge_range(worksheet, 0, 0, MAX_ROW, 0, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_merge_range(worksheet, 0, MAX_COL, 0, 0, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_merge_range(worksheet, MAX_ROW, 0, 0, 0, "Foo", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_set_column(worksheet, 6, MAX_COL, 17, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    err = lxlsx_worksheet_set_column(worksheet, MAX_COL, 6, 17, NULL);
    ASSERT_EQUAL(LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE, err);

    /* Tests array formula strings. */
    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_PARAMETER_IS_EMPTY, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "}", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_PARAMETER_IS_EMPTY, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{}", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_PARAMETER_IS_EMPTY, err);

    err = lxlsx_worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{=}", NULL);
    ASSERT_EQUAL(LXLSX_ERROR_PARAMETER_IS_EMPTY, err);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    lxlsx_worksheet_free(worksheet);
}
