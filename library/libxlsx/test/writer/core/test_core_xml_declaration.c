/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/core.h"

// Test _xml_declaration().
CTEST(core, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_core *core = lxlsx_core_new();
    core->file = testfile;

    _core_xml_declaration(core);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_core_free(core);
}
