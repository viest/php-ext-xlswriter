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
CTEST(relationships, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_relationships *relationships = lxw_relationships_new();
    relationships->file = testfile;

    _relationships_xml_declaration(relationships);

    RUN_XLSX_STREQ(exp, got);

    lxw_free_relationships(relationships);
}
