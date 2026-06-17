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

// Test the _write_sz() function.
CTEST(styles, write_sz) {


    char* got;
    char exp[] = "<sz val=\"11\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    styles->file = testfile;

    _write_font_size(styles, 11);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}

