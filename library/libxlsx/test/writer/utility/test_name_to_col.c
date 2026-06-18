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

// Test lxlsx_name_to_col().
CTEST(utility, lxlsx_name_to_col) {

    ASSERT_EQUAL(0,     lxlsx_name_to_col(NULL));
    ASSERT_EQUAL(0,     lxlsx_name_to_col(""));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("1"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("A"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("A1"));
    ASSERT_EQUAL(1,     lxlsx_name_to_col("B1"));
    ASSERT_EQUAL(2,     lxlsx_name_to_col("C1"));
    ASSERT_EQUAL(9,     lxlsx_name_to_col("J1"));
    ASSERT_EQUAL(24,    lxlsx_name_to_col("Y1"));
    ASSERT_EQUAL(25,    lxlsx_name_to_col("Z1"));
    ASSERT_EQUAL(26,    lxlsx_name_to_col("AA1"));
    ASSERT_EQUAL(254,   lxlsx_name_to_col("IU1"));
    ASSERT_EQUAL(255,   lxlsx_name_to_col("IV1"));
    ASSERT_EQUAL(256,   lxlsx_name_to_col("IW1"));
    ASSERT_EQUAL(16383, lxlsx_name_to_col("XFD1"));
    ASSERT_EQUAL(16384, lxlsx_name_to_col("XFE1"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("$A1"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("A$1"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col("$A$1"));
}


// Test lxlsx_name_to_col_2().
CTEST(utility, lxlsx_name_to_col_2) {

    ASSERT_EQUAL(0,     lxlsx_name_to_col_2(NULL));
    ASSERT_EQUAL(0,     lxlsx_name_to_col_2(""));
    ASSERT_EQUAL(0,     lxlsx_name_to_col_2("AAA"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col_2("AAA:"));
    ASSERT_EQUAL(0,     lxlsx_name_to_col_2("AAA:A"));
    ASSERT_EQUAL(1,     lxlsx_name_to_col_2("AAA:B"));
    ASSERT_EQUAL(2,     lxlsx_name_to_col_2("AAA:C"));
    ASSERT_EQUAL(9,     lxlsx_name_to_col_2("AAA:J"));
    ASSERT_EQUAL(24,    lxlsx_name_to_col_2("AAA:Y"));
    ASSERT_EQUAL(25,    lxlsx_name_to_col_2("AAA:Z"));
    ASSERT_EQUAL(26,    lxlsx_name_to_col_2("AAA:AA"));
    ASSERT_EQUAL(254,   lxlsx_name_to_col_2("AAA:IU"));
    ASSERT_EQUAL(255,   lxlsx_name_to_col_2("AAA:IV"));
    ASSERT_EQUAL(256,   lxlsx_name_to_col_2("AAA:IW"));
    ASSERT_EQUAL(16383, lxlsx_name_to_col_2("AAA:XFD"));
    ASSERT_EQUAL(16384, lxlsx_name_to_col_2("AAA:XFE"));
    ASSERT_EQUAL(16384, lxlsx_name_to_col_2("AAA1:XFE1"));
    ASSERT_EQUAL(16384, lxlsx_name_to_col_2("$AAA:$XFE"));
}
