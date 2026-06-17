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

// Test valid datetime.
CTEST(utility, test_datetime_validate01) {

    lxlsx_datetime datetime = {2025, 10, 30, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_NO_ERROR;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test valid datetime (time only).
CTEST(utility, test_datetime_validate02) {

    lxlsx_datetime datetime = {0, 0, 0, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_NO_ERROR;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test valid datetime (1900 epoch).
CTEST(utility, test_datetime_validate03) {

    lxlsx_datetime datetime = {1899, 12, 31, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_NO_ERROR;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test invalid year.
CTEST(utility, test_datetime_validate04) {

    lxlsx_datetime datetime = {1800, 10, 30, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test invalid month.
CTEST(utility, test_datetime_validate05) {

    lxlsx_datetime datetime = {1900, 13, 30, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}


// Test invalid day.
CTEST(utility, test_datetime_validate06) {

    lxlsx_datetime datetime = {1900, 10, 32, 21, 07, 0.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test invalid hour.
CTEST(utility, test_datetime_validate07) {

    lxlsx_datetime datetime = {1900, 1, 1, 24, 07, 0.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test invalid minute.
CTEST(utility, test_datetime_validate08) {
    lxlsx_datetime datetime = {1900, 1, 1, 21, 60, 0.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}

// Test invalid second.
CTEST(utility, test_datetime_validate09) {
    lxlsx_datetime datetime = {1900, 1, 1, 21, 07, 60.0};

    lxlsx_error exp = LXLSX_ERROR_DATETIME_VALIDATION;
    lxlsx_error got = lxlsx_datetime_validate(&datetime);

    ASSERT_EQUAL(exp, got);
}