/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/workbook.h"
#include "../../../include/lxlsx/shared_strings.h"

// Test assembling a complete Workbook file.
CTEST(workbook, workbook01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
          "<workbookPr defaultThemeVersion=\"124226\"/>"
          "<bookViews>"
            "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
          "</bookViews>"
          "<sheets>"
            "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"
          "</sheets>"
          "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        "</workbook>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_add_worksheet(workbook, NULL);

    lxw_workbook_assemble_xml_file(workbook);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_workbook_free(workbook);
}

// Test assembling a complete Workbook file.
CTEST(workbook, workbook02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
          "<workbookPr defaultThemeVersion=\"124226\"/>"
          "<bookViews>"
            "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
          "</bookViews>"
          "<sheets>"
            "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"
            "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>"
          "</sheets>"
          "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        "</workbook>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_add_worksheet(workbook, NULL);
    workbook_add_worksheet(workbook, NULL);

    lxw_workbook_assemble_xml_file(workbook);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_workbook_free(workbook);
}

// Test assembling a complete Workbook file.
CTEST(workbook, workbook03) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>"
          "<workbookPr defaultThemeVersion=\"124226\"/>"
          "<bookViews>"
            "<workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>"
          "</bookViews>"
          "<sheets>"
            "<sheet name=\"Non Default Name\" sheetId=\"1\" r:id=\"rId1\"/>"
            "<sheet name=\"Another Name\" sheetId=\"2\" r:id=\"rId2\"/>"
          "</sheets>"
          "<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/>"
        "</workbook>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_workbook *workbook = workbook_new(NULL);
    workbook->file = testfile;

    workbook_add_worksheet(workbook, "Non Default Name");
    workbook_add_worksheet(workbook, "Another Name");

    lxw_workbook_assemble_xml_file(workbook);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_workbook_free(workbook);
}
