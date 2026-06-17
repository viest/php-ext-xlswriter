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

// Test lxlsx_rowcol_to_range().
CTEST(utility, lxlsx_rowcol_to_range) {

    char got[LXLSX_MAX_CELL_RANGE_LENGTH];

    TEST_ROWCOL_TO_RANGE(0, 0, 0, 1, "A1:B1");
    TEST_ROWCOL_TO_RANGE(0, 2, 0, 9, "C1:J1");
    TEST_ROWCOL_TO_RANGE(1, 0, 2, 0, "A2:A3");
    TEST_ROWCOL_TO_RANGE(9, 0, 1, 24, "A10:Y2");
    TEST_ROWCOL_TO_RANGE(7, 25, 9, 26, "Z8:AA10");
    TEST_ROWCOL_TO_RANGE(1, 254, 1, 255, "IU2:IV2");
    TEST_ROWCOL_TO_RANGE(1, 256, 0, 16383, "IW2:XFD1");
    TEST_ROWCOL_TO_RANGE(0, 0, 1048576, 16384, "A1:XFE1048577");
    TEST_ROWCOL_TO_RANGE(1048575, 16383, 1048576, 16384, "XFD1048576:XFE1048577");

    // Test ranges that resolve to single cells.
    TEST_ROWCOL_TO_RANGE(0, 0, 0, 0, "A1");
    TEST_ROWCOL_TO_RANGE(1048576, 16384, 1048576, 16384, "XFE1048577");

}

// Test lxlsx_rowcol_to_range_abs().
CTEST(utility, lxlsx_rowcol_to_range_abs) {

    char got[LXLSX_MAX_CELL_RANGE_LENGTH];

    TEST_ROWCOL_TO_RANGE_ABS(0, 0, 0, 1, "$A$1:$B$1");
    TEST_ROWCOL_TO_RANGE_ABS(0, 2, 0, 9, "$C$1:$J$1");
    TEST_ROWCOL_TO_RANGE_ABS(1, 0, 2, 0, "$A$2:$A$3");
    TEST_ROWCOL_TO_RANGE_ABS(9, 0, 1, 24, "$A$10:$Y$2");
    TEST_ROWCOL_TO_RANGE_ABS(7, 25, 9, 26, "$Z$8:$AA$10");
    TEST_ROWCOL_TO_RANGE_ABS(1, 254, 1, 255, "$IU$2:$IV$2");
    TEST_ROWCOL_TO_RANGE_ABS(1, 256, 0, 16383, "$IW$2:$XFD$1");
    TEST_ROWCOL_TO_RANGE_ABS(0, 0, 1048576, 16384, "$A$1:$XFE$1048577");
    TEST_ROWCOL_TO_RANGE_ABS(1048575, 16383, 1048576, 16384, "$XFD$1048576:$XFE$1048577");

    // Test ranges that resolve to single cells.
    TEST_ROWCOL_TO_RANGE_ABS(0, 0, 0, 0, "$A$1");
    TEST_ROWCOL_TO_RANGE_ABS(1048576, 16384, 1048576, 16384, "$XFE$1048577");

}
