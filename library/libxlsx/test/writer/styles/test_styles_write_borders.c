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

// Test the _write_borders() function.
CTEST(styles, write_borders) {

    char* got;
    char exp[] = "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    lxw_format *format = lxw_format_new();

    format->has_border = 1;

    STAILQ_INSERT_TAIL(styles->xf_formats, format, list_pointers);

    styles->file = testfile;

    styles->border_count = 1;

    _write_borders(styles);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}

