/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/styles.h"

// Test the _write_border() function.
CTEST(styles, write_border) {

    char* got;
    char exp[] = "<border><left/><right/><top/><bottom/><diagonal/></border>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    styles->file = testfile;

    _write_border(styles, format, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

