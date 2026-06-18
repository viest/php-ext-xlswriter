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

// Test the _write_color() function.
CTEST(styles, write_color) {


    char* got;
    char exp[] = "<color theme=\"1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    styles->file = testfile;

    _write_font_color_theme(styles, 1);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

