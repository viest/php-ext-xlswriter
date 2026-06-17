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

// Test the _write_scheme() function.
CTEST(styles, write_scheme) {


    char* got;
    char exp[] = "<scheme val=\"minor\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    styles->file = testfile;

    _write_font_scheme(styles, "minor");

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

