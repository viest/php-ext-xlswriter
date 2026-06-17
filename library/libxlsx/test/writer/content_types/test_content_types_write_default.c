/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/content_types.h"

// Test the _write_default() function.
CTEST(content_types, write_default) {

    char* got;
    char exp[] = "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_content_types *content_types = lxlsx_content_types_new();
    content_types->file = testfile;

    _write_default(content_types, "xml", "application/xml");

    RUN_XLSX_STREQ(exp, got);

    lxlsx_content_types_free(content_types);
}

