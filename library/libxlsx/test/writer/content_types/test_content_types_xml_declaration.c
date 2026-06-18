/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/content_types.h"

// Test _xml_declaration().
CTEST(content_types, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_content_types *content_types = lxlsx_content_types_new();
    content_types->file = testfile;

    _content_types_xml_declaration(content_types);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_content_types_free(content_types);
}
