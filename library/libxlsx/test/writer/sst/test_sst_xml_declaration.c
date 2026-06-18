/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/shared_strings.h"

// Test _xml_declaration().
CTEST(sst, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_sst *sst = lxlsx_sst_new();
    sst->file = testfile;

    _sst_xml_declaration(sst);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_sst_free(sst);
}
