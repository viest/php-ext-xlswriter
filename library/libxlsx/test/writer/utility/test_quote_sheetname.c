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


// Test _datetime_to_excel_date().
CTEST(utility, lxw_quote_sheetname) {

    ASSERT_STR("Sheet1",     lxw_quote_sheetname("Sheet1"));
    ASSERT_STR("Sheet.2",    lxw_quote_sheetname("Sheet.2"));
    ASSERT_STR("Sheet_3",    lxw_quote_sheetname("Sheet_3"));
    ASSERT_STR("'Sheet4'",   lxw_quote_sheetname("'Sheet4'"));
    ASSERT_STR("'Sheet 5'",  lxw_quote_sheetname("Sheet 5"));
    ASSERT_STR("'Sheet!6'",  lxw_quote_sheetname("Sheet!6"));
    ASSERT_STR("'Sheet''7'", lxw_quote_sheetname("Sheet'7"));
    ASSERT_STR("'a''''''''''''''''''''''''''''''''''''''''''''''''''''''''''b'",
               lxw_quote_sheetname("a'''''''''''''''''''''''''''''b"));
}
