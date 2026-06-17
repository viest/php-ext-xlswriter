/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/chartsheet.h"

// Test _xml_declaration().
CTEST(chartsheet, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;

    _chartsheet_xml_declaration(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}
