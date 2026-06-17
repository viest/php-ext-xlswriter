/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/utility.h"


// Test lxlsx_rowcol_to_formula_abs().
CTEST(utility, lxlsx_rowcol_to_formula_abs) {

    char got[LXLSX_MAX_FORMULA_RANGE_LENGTH];

    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 0, 0, 0, 1, "Sheet1!$A$1:$B$1");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 0, 2, 0, 9, "Sheet1!$C$1:$J$1");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 1, 0, 2, 0, "Sheet1!$A$2:$A$3");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 9, 0, 1, 24, "Sheet1!$A$10:$Y$2");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 7, 25, 9, 26, "Sheet1!$Z$8:$AA$10");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 1, 254, 1, 255, "Sheet1!$IU$2:$IV$2");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 1, 256, 0, 16383, "Sheet1!$IW$2:$XFD$1");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 0, 0, 1048576, 16384, "Sheet1!$A$1:$XFE$1048577");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 1048575, 16383, 1048576, 16384, "Sheet1!$XFD$1048576:$XFE$1048577");

    TEST_ROWCOL_TO_FORMULA_ABS("New data",   1, 2, 8, 2, "'New data'!$C$2:$C$9");
    TEST_ROWCOL_TO_FORMULA_ABS("'New data'", 1, 2, 8, 2, "'New data'!$C$2:$C$9");


    // Test ranges that resolve to single cells.
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 0, 0, 0, 0, "Sheet1!$A$1");
    TEST_ROWCOL_TO_FORMULA_ABS("Sheet1", 1048576, 16384, 1048576, 16384, "Sheet1!$XFE$1048577");

}
