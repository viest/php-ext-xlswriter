/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/shared_strings.h"

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

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_sst *sst = lxlsx_sst_new();
    sst->file = testfile;

    lxlsx_get_sst_index(sst, "neptune", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "neptune", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "neptune", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "mars", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "mars", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "venus", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "venus", LXLSX_FALSE);

    lxlsx_sst_assemble_xml_file(sst);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_sst_free(sst);
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

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_sst *sst = lxlsx_sst_new();
    sst->file = testfile;

    // Test strings with whitespace that must be preserved.
    lxlsx_get_sst_index(sst, "abcdefg", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "   abcdefg", LXLSX_FALSE);
    lxlsx_get_sst_index(sst, "abcdefg   ", LXLSX_FALSE);

    lxlsx_sst_assemble_xml_file(sst);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_sst_free(sst);
}
