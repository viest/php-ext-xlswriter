/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/utility.h"

// Test xl_rowcol_to_cell().
CTEST(utility, lxlsx_rowcol_to_cell) {

    char got[LXLSX_MAX_CELL_NAME_LENGTH];

    TEST_ROWCOL_TO_CELL(0, 0, "A1");
    TEST_ROWCOL_TO_CELL(0, 1, "B1");
    TEST_ROWCOL_TO_CELL(0, 2, "C1");
    TEST_ROWCOL_TO_CELL(0, 9, "J1");
    TEST_ROWCOL_TO_CELL(1, 0, "A2");
    TEST_ROWCOL_TO_CELL(2, 0, "A3");
    TEST_ROWCOL_TO_CELL(9, 0, "A10");
    TEST_ROWCOL_TO_CELL(1, 24, "Y2");
    TEST_ROWCOL_TO_CELL(7, 25, "Z8");
    TEST_ROWCOL_TO_CELL(9, 26, "AA10");
    TEST_ROWCOL_TO_CELL(1, 254, "IU2");
    TEST_ROWCOL_TO_CELL(1, 255, "IV2");
    TEST_ROWCOL_TO_CELL(1, 256, "IW2");
    TEST_ROWCOL_TO_CELL(0, 16383, "XFD1");
    TEST_ROWCOL_TO_CELL(1048576, 16384, "XFE1048577");
}

// Test xl_rowcol_to_cell_abs().
CTEST(utility, lxlsx_rowcol_to_cell_abs) {

    char got[LXLSX_MAX_CELL_NAME_LENGTH];

    TEST_ROWCOL_TO_CELL_ABS(0, 0, 0, 0, "A1");
    TEST_ROWCOL_TO_CELL_ABS(0, 1, 0, 0, "B1");
    TEST_ROWCOL_TO_CELL_ABS(0, 2, 0, 0, "C1");
    TEST_ROWCOL_TO_CELL_ABS(0, 9, 0, 0, "J1");
    TEST_ROWCOL_TO_CELL_ABS(1, 0, 0, 0, "A2");
    TEST_ROWCOL_TO_CELL_ABS(2, 0, 0, 0, "A3");
    TEST_ROWCOL_TO_CELL_ABS(9, 0, 0, 0, "A10");
    TEST_ROWCOL_TO_CELL_ABS(1, 24, 0, 0, "Y2");
    TEST_ROWCOL_TO_CELL_ABS(7, 25, 0, 0, "Z8");
    TEST_ROWCOL_TO_CELL_ABS(9, 26, 0, 0, "AA10");
    TEST_ROWCOL_TO_CELL_ABS(1, 254, 0, 0, "IU2");
    TEST_ROWCOL_TO_CELL_ABS(1, 255, 0, 0, "IV2");
    TEST_ROWCOL_TO_CELL_ABS(1, 256, 0, 0, "IW2");
    TEST_ROWCOL_TO_CELL_ABS(0, 16383, 0, 0, "XFD1");
    TEST_ROWCOL_TO_CELL_ABS(1048576, 16384, 0, 0, "XFE1048577");

    TEST_ROWCOL_TO_CELL_ABS(0, 0, 1, 0, "A$1");
    TEST_ROWCOL_TO_CELL_ABS(0, 0, 0, 1, "$A1");
    TEST_ROWCOL_TO_CELL_ABS(0, 0, 1, 1, "$A$1");

    TEST_ROWCOL_TO_CELL_ABS(1048576, 16384, 1, 0, "XFE$1048577");
    TEST_ROWCOL_TO_CELL_ABS(1048576, 16384, 0, 1, "$XFE1048577");
    TEST_ROWCOL_TO_CELL_ABS(1048576, 16384, 1, 1, "$XFE$1048577");
}
