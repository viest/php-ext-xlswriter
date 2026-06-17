/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/content_types.h"

// Test the _write_override() function.
CTEST(content_types, write_override) {

    char* got;
    char exp[] = "<Override PartName=\"/docProps/core.xml\" ContentType=\"app...\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_content_types *content_types = lxw_content_types_new();
    content_types->file = testfile;

    _write_override(content_types, "/docProps/core.xml", "app...");

    RUN_XLSX_STREQ(exp, got);

    lxw_content_types_free(content_types);
}

