/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/styles.h"

// Test the _write_fills() function.
CTEST(styles, write_fills) {

    char* got;
    char exp[] = "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    styles->fill_count = 2;
    styles->file = testfile;

    _write_fills(styles);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

