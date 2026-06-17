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

// Test lxlsx_strerror() to ensure the error_string array doesn't go out of sync.
CTEST(utility, lxlsx_strerror) {

    ASSERT_STR("No error.",
               lxlsx_strerror(LXLSX_NO_ERROR));

    ASSERT_STR("Error encountered when creating a tmpfile during file assembly.",
               lxlsx_strerror(LXLSX_ERROR_CREATING_TMPFILE));

    ASSERT_STR("Maximum number of worksheet URLs (65530) exceeded.",
               lxlsx_strerror(LXLSX_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED));

    ASSERT_STR("Unknown error number.",
               lxlsx_strerror(LXLSX_MAX_ERRNO));

    ASSERT_STR("Unknown error number.",
               lxlsx_strerror(LXLSX_MAX_ERRNO + 1));
}
