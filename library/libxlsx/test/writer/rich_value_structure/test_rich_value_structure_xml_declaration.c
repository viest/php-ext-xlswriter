/*
 * Tests for the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "lxlsx/rich_value_structure.h"

// Test _xml_declaration().
CTEST(rich_value_structure, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxw_rich_value_structure *rich_value_structure = lxw_rich_value_structure_new();
    rich_value_structure->file = testfile;

    _rich_value_structure_xml_declaration(rich_value_structure);

    RUN_XLSX_STREQ(exp, got);

    lxw_rich_value_structure_free(rich_value_structure);
}
