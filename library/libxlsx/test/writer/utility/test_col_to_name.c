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

// Test _xl_col_to_name().
CTEST(utility, lxlsx_col_to_name) {

    char got[LXLSX_MAX_COL_NAME_LENGTH];

    TEST_COL_TO_NAME(0, 0, "A");
    TEST_COL_TO_NAME(1, 0, "B");
    TEST_COL_TO_NAME(2, 0, "C");
    TEST_COL_TO_NAME(9, 0, "J");
    TEST_COL_TO_NAME(24, 0, "Y");
    TEST_COL_TO_NAME(25, 0, "Z");
    TEST_COL_TO_NAME(26, 0, "AA");
    TEST_COL_TO_NAME(254, 0, "IU");
    TEST_COL_TO_NAME(255, 0, "IV");
    TEST_COL_TO_NAME(256, 0, "IW");
    TEST_COL_TO_NAME(16383, 0, "XFD");
    TEST_COL_TO_NAME(16384, 0, "XFE");

    TEST_COL_TO_NAME(0, 1, "$A");
    TEST_COL_TO_NAME(1, 1, "$B");
    TEST_COL_TO_NAME(2, 1, "$C");
    TEST_COL_TO_NAME(9, 1, "$J");
    TEST_COL_TO_NAME(24, 1, "$Y");
    TEST_COL_TO_NAME(25, 1, "$Z");
    TEST_COL_TO_NAME(26, 1, "$AA");
    TEST_COL_TO_NAME(254, 1, "$IU");
    TEST_COL_TO_NAME(255, 1, "$IV");
    TEST_COL_TO_NAME(256, 1, "$IW");
    TEST_COL_TO_NAME(16383, 1, "$XFD");
    TEST_COL_TO_NAME(16384, 1, "$XFE");
}
