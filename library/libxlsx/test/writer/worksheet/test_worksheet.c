/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/worksheet.h"
#include "../../../include/lxlsx/shared_strings.h"

// Test assembling a complete Worksheet file.
CTEST(worksheet, worksheet01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<dimension ref=\"A1\"/>"
          "<sheetViews>"
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
          "</sheetViews>"
          "<sheetFormatPr defaultRowHeight=\"15\"/>"
          "<sheetData/>"
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
}

// Test assembling a complete Worksheet file.
CTEST(worksheet, worksheet02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<dimension ref=\"A1\"/>"
          "<sheetViews>"
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
          "</sheetViews>"
          "<sheetFormatPr defaultRowHeight=\"15\"/>"
          "<sheetData>"
            "<row r=\"1\" spans=\"1:1\">"
              "<c r=\"A1\">"
                "<v>123</v>"
              "</c>"
            "</row>"
          "</sheetData>"
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_write_number(worksheet, 0, 0, 123, NULL);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
}

// Test assembling a complete Worksheet file.
CTEST(worksheet, worksheet03) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<dimension ref=\"A1:E9\"/>"
          "<sheetViews>"
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
          "</sheetViews>"
          "<sheetFormatPr defaultRowHeight=\"15\"/>"
          "<sheetData>"
            "<row r=\"1\" spans=\"1:5\">"
              "<c r=\"A1\" t=\"s\">"
                "<v>0</v>"
              "</c>"
            "</row>"
            "<row r=\"2\" spans=\"1:5\">"
              "<c r=\"C2\">"
                "<v>123</v>"
              "</c>"
            "</row>"
            "<row r=\"4\" spans=\"1:5\">"
              "<c r=\"B4\" t=\"s\">"
                "<v>1</v>"
              "</c>"
            "</row>"
            "<row r=\"9\" spans=\"1:5\">"
              "<c r=\"E9\">"
                "<v>890</v>"
              "</c>"
            "</row>"
          "</sheetData>"
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet->sst = lxw_sst_new();
    worksheet_select(worksheet);

    worksheet_write_string(worksheet, 0, 0, "Foo", NULL);
    worksheet_write_number(worksheet, 1, 2, 123, NULL);
    worksheet_write_string(worksheet, 3, 1, "Bar", NULL);
    worksheet_write_number(worksheet, 8, 4, 890, NULL);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_sst_free(worksheet->sst);
    lxw_worksheet_free(worksheet);
}


