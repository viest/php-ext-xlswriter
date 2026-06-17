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

// Test the _write_name() function.
CTEST(styles, write_name) {


    char* got;
    char exp[] = "<name val=\"Calibri\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    styles->file = testfile;

    _write_font_name(styles, "Calibri", LXW_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}

