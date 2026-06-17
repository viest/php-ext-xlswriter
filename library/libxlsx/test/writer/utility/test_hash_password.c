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


// Test lxw_hash_password().
CTEST(utility, lxw_hash_password) {

    ASSERT_EQUAL(0x83AF, lxw_hash_password("password"));
    ASSERT_EQUAL(0xD14E, lxw_hash_password("This is a longer phrase"));
    ASSERT_EQUAL(0xCE2A, lxw_hash_password("0"));
    ASSERT_EQUAL(0xCEED, lxw_hash_password("01"));
    ASSERT_EQUAL(0xCF7C, lxw_hash_password("012"));
    ASSERT_EQUAL(0xCC4B, lxw_hash_password("0123"));
    ASSERT_EQUAL(0xCACA, lxw_hash_password("01234"));
    ASSERT_EQUAL(0xC789, lxw_hash_password("012345"));
    ASSERT_EQUAL(0xDC88, lxw_hash_password("0123456"));
    ASSERT_EQUAL(0xEB87, lxw_hash_password("01234567"));
    ASSERT_EQUAL(0x9B86, lxw_hash_password("012345678"));
    ASSERT_EQUAL(0xFF84, lxw_hash_password("0123456789"));
    ASSERT_EQUAL(0xFF86, lxw_hash_password("01234567890"));
    ASSERT_EQUAL(0xEF87, lxw_hash_password("012345678901"));
    ASSERT_EQUAL(0xAF8A, lxw_hash_password("0123456789012"));
    ASSERT_EQUAL(0xEF90, lxw_hash_password("01234567890123"));
    ASSERT_EQUAL(0xEFA5, lxw_hash_password("012345678901234"));
    ASSERT_EQUAL(0xEFD0, lxw_hash_password("0123456789012345"));
    ASSERT_EQUAL(0xEF09, lxw_hash_password("01234567890123456"));
    ASSERT_EQUAL(0xEEB2, lxw_hash_password("012345678901234567"));
    ASSERT_EQUAL(0xED33, lxw_hash_password("0123456789012345678"));
    ASSERT_EQUAL(0xEA14, lxw_hash_password("01234567890123456789"));
    ASSERT_EQUAL(0xE615, lxw_hash_password("012345678901234567890"));
    ASSERT_EQUAL(0xFE96, lxw_hash_password("0123456789012345678901"));
    ASSERT_EQUAL(0xCC97, lxw_hash_password("01234567890123456789012"));
    ASSERT_EQUAL(0xAA98, lxw_hash_password("012345678901234567890123"));
    ASSERT_EQUAL(0xFA98, lxw_hash_password("0123456789012345678901234"));
    ASSERT_EQUAL(0xD298, lxw_hash_password("01234567890123456789012345"));
    ASSERT_EQUAL(0xD2D3, lxw_hash_password("0123456789012345678901234567890"));
}
