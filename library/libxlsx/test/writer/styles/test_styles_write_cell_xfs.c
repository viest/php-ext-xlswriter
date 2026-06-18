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

// Test the _write_cell_xfs() function.
CTEST(styles, write_cell_xfs) {

    char* got;
    char exp[] = "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    STAILQ_INSERT_TAIL(styles->xf_formats, format, list_pointers);

    styles->file = testfile;

    styles->xf_count = 1;

    _write_cell_xfs(styles);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
}

