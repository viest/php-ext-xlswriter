/*
 * Tests for the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "lxlsx/rich_value_rel.h"

// Test _xml_declaration().
CTEST(lxlsx_rich_value_rel, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxlsx_rich_value_rel *lxlsx_rich_value_rel = lxlsx_rich_value_rel_new();
    lxlsx_rich_value_rel->file = testfile;

    _rich_value_rel_xml_declaration(lxlsx_rich_value_rel);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_rich_value_rel_free(lxlsx_rich_value_rel);
}
