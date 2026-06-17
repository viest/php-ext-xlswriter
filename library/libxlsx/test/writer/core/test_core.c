/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/core.h"
#include "../../../include/libxlsx/workbook.h"

#ifdef _WIN32
#define timegm _mkgmtime
#define strdup _strdup
#endif

// Test assembling a complete core file.
CTEST(core, core01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
          "<dc:creator>A User</dc:creator>"
          "<cp:lastModifiedBy>A User</cp:lastModifiedBy>"
          "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2010-01-01T00:00:00Z</dcterms:created>"
          "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2010-01-01T00:00:00Z</dcterms:modified>"
        "</cp:coreProperties>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_core *core = lxlsx_core_new();
    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);

    core->file = testfile;
    core->properties = workbook->properties;

    // Add data to the core->properties.
    struct tm tmp_tm;
    tmp_tm.tm_year  = 110;
    tmp_tm.tm_mon   = 0;
    tmp_tm.tm_mday  = 1;
    tmp_tm.tm_hour  = 0;
    tmp_tm.tm_min   = 0;
    tmp_tm.tm_sec   = 0;
    tmp_tm.tm_isdst = -1;

    core->properties->created = timegm(&tmp_tm);
    core->properties->author  = strdup("A User");

    lxlsx_core_assemble_xml_file(core);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_workbook_free(workbook);
    lxlsx_core_free(core);
}


// Test assembling a complete core file.
CTEST(core, core02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
          "<dc:title>This is an example spreadsheet</dc:title>"
          "<dc:subject>With document properties</dc:subject>"
          "<dc:creator>A Person</dc:creator>"
          "<cp:keywords>Sample, Example, Properties</cp:keywords>"
          "<dc:description>Created with libxlsxwriter</dc:description>"
          "<cp:lastModifiedBy>A Person</cp:lastModifiedBy>"
          "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2011-04-06T19:45:15Z</dcterms:created>"
          "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2011-04-06T19:45:15Z</dcterms:modified>"
          "<cp:category>Example spreadsheets</cp:category>"
          "<cp:contentStatus>Quo</cp:contentStatus>"
        "</cp:coreProperties>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_core *core = lxlsx_core_new();
    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);

    core->file = testfile;
    core->properties = workbook->properties;

    // Add data to the core->properties.
    struct tm tmp_tm;
    tmp_tm.tm_year  = 111;
    tmp_tm.tm_mon   = 3;
    tmp_tm.tm_mday  = 6;
    tmp_tm.tm_hour  = 19;
    tmp_tm.tm_min   = 45;
    tmp_tm.tm_sec   = 15;
    tmp_tm.tm_isdst = -1;

    core->properties->created  = timegm(&tmp_tm);
    core->properties->title    = strdup("This is an example spreadsheet");
    core->properties->subject  = strdup("With document properties");
    core->properties->author   = strdup("A Person");
    core->properties->keywords = strdup("Sample, Example, Properties");
    core->properties->comments = strdup("Created with libxlsxwriter");
    core->properties->category = strdup("Example spreadsheets");
    core->properties->status   = strdup("Quo");

    lxlsx_core_assemble_xml_file(core);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_workbook_free(workbook);
    lxlsx_core_free(core);
}
