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

// Test the _write_font() function.
CTEST(styles, write_fonts01) {

    char* got;
    char exp[] = "<fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    lxw_format *format = lxw_format_new();

    format->has_font = 1;

    STAILQ_INSERT_TAIL(styles->xf_formats, format, list_pointers);

    styles->file = testfile;
    styles->font_count = 1;

    _write_fonts(styles);

    RUN_XLSX_STREQ(exp, got);

    lxw_styles_free(styles);
}
