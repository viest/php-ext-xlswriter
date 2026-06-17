/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/shared_strings.h"

// Test _xml_declaration().
CTEST(sst, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_sst *sst = lxw_sst_new();
    sst->file = testfile;

    _sst_xml_declaration(sst);

    RUN_XLSX_STREQ(exp, got);

    lxw_sst_free(sst);
}
