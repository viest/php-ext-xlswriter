/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/app.h"

// Test assembling a complete App file.
CTEST(app, app01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
          "<Application>Microsoft Excel</Application>"
          "<DocSecurity>0</DocSecurity>"
          "<ScaleCrop>false</ScaleCrop>"
          "<HeadingPairs>"
            "<vt:vector size=\"2\" baseType=\"variant\">"
              "<vt:variant>"
                "<vt:lpstr>Worksheets</vt:lpstr>"
              "</vt:variant>"
              "<vt:variant>"
                "<vt:i4>1</vt:i4>"
              "</vt:variant>"
            "</vt:vector>"
          "</HeadingPairs>"
          "<TitlesOfParts>"
            "<vt:vector size=\"1\" baseType=\"lpstr\">"
              "<vt:lpstr>Sheet1</vt:lpstr>"
            "</vt:vector>"
          "</TitlesOfParts>"
          "<Company>"
          "</Company>"
          "<LinksUpToDate>false</LinksUpToDate>"
          "<SharedDoc>false</SharedDoc>"
          "<HyperlinksChanged>false</HyperlinksChanged>"
          "<AppVersion>12.0000</AppVersion>"
        "</Properties>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_app *app = lxlsx_app_new();
    app->file = testfile;

    lxlsx_app_add_part_name(app,"Sheet1");
    lxlsx_app_add_heading_pair(app, "Worksheets", "1");

    lxlsx_app_assemble_xml_file(app);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_app_free(app);
}

// Test assembling a complete App file.
CTEST(app, app02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
          "<Application>Microsoft Excel</Application>"
          "<DocSecurity>0</DocSecurity>"
          "<ScaleCrop>false</ScaleCrop>"
          "<HeadingPairs>"
            "<vt:vector size=\"2\" baseType=\"variant\">"
              "<vt:variant>"
                "<vt:lpstr>Worksheets</vt:lpstr>"
              "</vt:variant>"
              "<vt:variant>"
                "<vt:i4>2</vt:i4>"
              "</vt:variant>"
            "</vt:vector>"
          "</HeadingPairs>"
          "<TitlesOfParts>"
            "<vt:vector size=\"2\" baseType=\"lpstr\">"
              "<vt:lpstr>Sheet1</vt:lpstr>"
              "<vt:lpstr>Sheet2</vt:lpstr>"
            "</vt:vector>"
          "</TitlesOfParts>"
          "<Company>"
          "</Company>"
          "<LinksUpToDate>false</LinksUpToDate>"
          "<SharedDoc>false</SharedDoc>"
          "<HyperlinksChanged>false</HyperlinksChanged>"
          "<AppVersion>12.0000</AppVersion>"
        "</Properties>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_app *app = lxlsx_app_new();
    app->file = testfile;

    lxlsx_app_add_part_name(app,"Sheet1");
    lxlsx_app_add_part_name(app,"Sheet2");
    lxlsx_app_add_heading_pair(app, "Worksheets", "2");

    lxlsx_app_assemble_xml_file(app);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_app_free(app);
}


// Test assembling a complete App file.
CTEST(app, app03) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
          "<Application>Microsoft Excel</Application>"
          "<DocSecurity>0</DocSecurity>"
          "<ScaleCrop>false</ScaleCrop>"
          "<HeadingPairs>"
            "<vt:vector size=\"4\" baseType=\"variant\">"
              "<vt:variant>"
                "<vt:lpstr>Worksheets</vt:lpstr>"
              "</vt:variant>"
              "<vt:variant>"
                "<vt:i4>1</vt:i4>"
              "</vt:variant>"
              "<vt:variant>"
                "<vt:lpstr>Named Ranges</vt:lpstr>"
              "</vt:variant>"
              "<vt:variant>"
                "<vt:i4>1</vt:i4>"
              "</vt:variant>"
            "</vt:vector>"
          "</HeadingPairs>"
          "<TitlesOfParts>"
            "<vt:vector size=\"2\" baseType=\"lpstr\">"
              "<vt:lpstr>Sheet1</vt:lpstr>"
              "<vt:lpstr>Sheet1!Print_Titles</vt:lpstr>"
            "</vt:vector>"
          "</TitlesOfParts>"
          "<Company>"
          "</Company>"
          "<LinksUpToDate>false</LinksUpToDate>"
          "<SharedDoc>false</SharedDoc>"
          "<HyperlinksChanged>false</HyperlinksChanged>"
          "<AppVersion>12.0000</AppVersion>"
        "</Properties>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_app *app = lxlsx_app_new();
    app->file = testfile;

    lxlsx_app_add_part_name(app,"Sheet1");
    lxlsx_app_add_part_name(app,"Sheet1!Print_Titles");
    lxlsx_app_add_heading_pair(app, "Worksheets", "1");
    lxlsx_app_add_heading_pair(app, "Named Ranges", "1");

    lxlsx_app_assemble_xml_file(app);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_app_free(app);
}

