/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/relationships.h"

// Test _xml_declaration().
CTEST(relationships, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_relationships *relationships = lxlsx_relationships_new();
    relationships->file = testfile;

    _relationships_xml_declaration(relationships);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_free_relationships(relationships);
}
