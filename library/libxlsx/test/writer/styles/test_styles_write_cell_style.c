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

// Test the _write_cell_style() function.
CTEST(styles, write_cell_style) {

    char* got;
    char exp[] = "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    styles->file = testfile;

    _write_cell_style(styles, "Normal", 0, 0);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}

