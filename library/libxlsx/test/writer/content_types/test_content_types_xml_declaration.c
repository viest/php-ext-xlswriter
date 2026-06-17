/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/content_types.h"

// Test _xml_declaration().
CTEST(content_types, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_content_types *content_types = lxw_content_types_new();
    content_types->file = testfile;

    _content_types_xml_declaration(content_types);

    RUN_XLSX_STREQ(exp, got);

    lxw_content_types_free(content_types);
}
