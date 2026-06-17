/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/shared_strings.h"

// Test assembling a complete SharedStrings file.
CTEST(sst, sst01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"7\" uniqueCount=\"3\">"
          "<si>"
            "<t>neptune</t>"
          "</si>"
          "<si>"
            "<t>mars</t>"
          "</si>"
          "<si>"
            "<t>venus</t>"
          "</si>"
        "</sst>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_sst *sst = lxw_sst_new();
    sst->file = testfile;

    lxw_get_sst_index(sst, "neptune", LXW_FALSE);
    lxw_get_sst_index(sst, "neptune", LXW_FALSE);
    lxw_get_sst_index(sst, "neptune", LXW_FALSE);
    lxw_get_sst_index(sst, "mars", LXW_FALSE);
    lxw_get_sst_index(sst, "mars", LXW_FALSE);
    lxw_get_sst_index(sst, "venus", LXW_FALSE);
    lxw_get_sst_index(sst, "venus", LXW_FALSE);

    lxw_sst_assemble_xml_file(sst);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_sst_free(sst);
}

// Test assembling a complete SharedStrings file.
CTEST(sst, sst02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">"
          "<si>"
            "<t>abcdefg</t>"
          "</si>"
          "<si>"
            "<t xml:space=\"preserve\">   abcdefg</t>"
          "</si>"
          "<si>"
            "<t xml:space=\"preserve\">abcdefg   </t>"
          "</si>"
        "</sst>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_sst *sst = lxw_sst_new();
    sst->file = testfile;

    // Test strings with whitespace that must be preserved.
    lxw_get_sst_index(sst, "abcdefg", LXW_FALSE);
    lxw_get_sst_index(sst, "   abcdefg", LXW_FALSE);
    lxw_get_sst_index(sst, "abcdefg   ", LXW_FALSE);

    lxw_sst_assemble_xml_file(sst);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_sst_free(sst);
}
