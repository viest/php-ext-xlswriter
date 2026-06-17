/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/custom.h"

// Test _xml_declaration().
CTEST(custom, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_custom *custom = lxlsx_custom_new();
    custom->file = testfile;

    _custom_xml_declaration(custom);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_custom_free(custom);
}
