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

// Test lxlsx_utf8_strlen().
CTEST(utility, lxlsx_utf8_strlen) {

    ASSERT_EQUAL(0,  (long)lxlsx_utf8_strlen(""));
    ASSERT_EQUAL(3,  (long)lxlsx_utf8_strlen("Foo"));
    ASSERT_EQUAL(4,  (long)lxlsx_utf8_strlen("café"));
    ASSERT_EQUAL(4,  (long)lxlsx_utf8_strlen("cake"));
    ASSERT_EQUAL(21, (long)lxlsx_utf8_strlen("Это фраза на русском!"));

}

