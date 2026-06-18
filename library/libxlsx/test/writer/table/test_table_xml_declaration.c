/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "libxlsx/table.h"

// Test _xml_declaration().
CTEST(table, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxlsx_table *table = lxlsx_table_new();
    table->file = testfile;

    _table_xml_declaration(table);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_table_free(table);
}
