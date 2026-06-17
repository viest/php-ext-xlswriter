/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/chart.h"

// Test _xml_declaration().
CTEST(chart, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_chart *chart = lxlsx_chart_new(LXLSX_CHART_AREA);
    chart->file = testfile;

    _chart_xml_declaration(chart);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_chart_free(chart);
}
