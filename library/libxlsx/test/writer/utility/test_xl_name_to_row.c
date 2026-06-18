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

// Test lxlsx_name_to_row().
CTEST(utility, lxlsx_name_to_row) {

    ASSERT_EQUAL(0,       lxlsx_name_to_row(NULL));
    ASSERT_EQUAL(0,       lxlsx_name_to_row(""));
    ASSERT_EQUAL(0,       lxlsx_name_to_row("A"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row("A0"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row("A1"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row("$A$1"));
    ASSERT_EQUAL(1,       lxlsx_name_to_row("B2"));
    ASSERT_EQUAL(2,       lxlsx_name_to_row("C3"));
    ASSERT_EQUAL(9,       lxlsx_name_to_row("J10"));
    ASSERT_EQUAL(24,      lxlsx_name_to_row("Y25"));
    ASSERT_EQUAL(25,      lxlsx_name_to_row("Z26"));
    ASSERT_EQUAL(26,      lxlsx_name_to_row("AA27"));
    ASSERT_EQUAL(254,     lxlsx_name_to_row("IU255"));
    ASSERT_EQUAL(255,     lxlsx_name_to_row("IV256"));
    ASSERT_EQUAL(256,     lxlsx_name_to_row("IW257"));
    ASSERT_EQUAL(16383,   lxlsx_name_to_row("XFD16384"));
    ASSERT_EQUAL(16384,   lxlsx_name_to_row("XFE16385"));
    ASSERT_EQUAL(1048576, lxlsx_name_to_row("XFE1048577"));
    ASSERT_EQUAL(1048576, lxlsx_name_to_row("$XFE$1048577"));
}

// Test lxlsx_name_to_row().
CTEST(utility, lxlsx_name_to_row_2) {

    ASSERT_EQUAL(0,       lxlsx_name_to_row_2(NULL));
    ASSERT_EQUAL(0,       lxlsx_name_to_row_2(""));
    ASSERT_EQUAL(0,       lxlsx_name_to_row_2("A1:A"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row_2("A1:A0"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row_2("A1:A1"));
    ASSERT_EQUAL(0,       lxlsx_name_to_row_2("A1:$A$1"));
    ASSERT_EQUAL(1,       lxlsx_name_to_row_2("A1:B2"));
    ASSERT_EQUAL(2,       lxlsx_name_to_row_2("A1:C3"));
    ASSERT_EQUAL(9,       lxlsx_name_to_row_2("A1:J10"));
    ASSERT_EQUAL(24,      lxlsx_name_to_row_2("A1:Y25"));
    ASSERT_EQUAL(25,      lxlsx_name_to_row_2("A1:Z26"));
    ASSERT_EQUAL(26,      lxlsx_name_to_row_2("A1:AA27"));
    ASSERT_EQUAL(254,     lxlsx_name_to_row_2("A1:IU255"));
    ASSERT_EQUAL(255,     lxlsx_name_to_row_2("A1:IV256"));
    ASSERT_EQUAL(256,     lxlsx_name_to_row_2("A1:IW257"));
    ASSERT_EQUAL(16383,   lxlsx_name_to_row_2("A1:XFD16384"));
    ASSERT_EQUAL(16384,   lxlsx_name_to_row_2("A1:XFE16385"));
    ASSERT_EQUAL(1048576, lxlsx_name_to_row_2("A1:XFE1048577"));
    ASSERT_EQUAL(1048576, lxlsx_name_to_row_2("A1:$XFE$1048577"));
}

