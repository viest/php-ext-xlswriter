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

// Test the _write_borders() function.
CTEST(styles, write_borders) {

    char* got;
    char exp[] = "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    format->has_border = 1;

    STAILQ_INSERT_TAIL(styles->xf_formats, format, list_pointers);

    styles->file = testfile;

    styles->border_count = 1;

    _write_borders(styles);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

