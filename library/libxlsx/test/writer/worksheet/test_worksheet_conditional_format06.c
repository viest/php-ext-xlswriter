/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/worksheet.h"
#include "../../../include/libxlsx/shared_strings.h"

// Test assembling a complete Worksheet file.
CTEST(worksheet, lxlsx_worksheet_condtional_format06) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
              "<dimension ref=\"A1:A4\"/>"
              "<sheetViews>"
                "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
              "</sheetViews>"
              "<sheetFormatPr defaultRowHeight=\"15\"/>"
              "<sheetData>"
                "<row r=\"1\" spans=\"1:1\">"
                  "<c r=\"A1\">"
                    "<v>10</v>"
                  "</c>"
                "</row>"
                "<row r=\"2\" spans=\"1:1\">"
                  "<c r=\"A2\">"
                    "<v>20</v>"
                  "</c>"
                "</row>"
                "<row r=\"3\" spans=\"1:1\">"
                  "<c r=\"A3\">"
                    "<v>30</v>"
                  "</c>"
                "</row>"
                "<row r=\"4\" spans=\"1:1\">"
                  "<c r=\"A4\">"
                    "<v>40</v>"
                  "</c>"
                "</row>"
              "</sheetData>"
              "<conditionalFormatting sqref=\"A1:A4\">"
                "<cfRule type=\"top10\" priority=\"1\" rank=\"15\"/>"
                "<cfRule type=\"top10\" priority=\"2\" bottom=\"1\" rank=\"16\"/>"
                "<cfRule type=\"top10\" priority=\"3\" percent=\"1\" rank=\"17\"/>"
                "<cfRule type=\"top10\" priority=\"4\" percent=\"1\" bottom=\"1\" rank=\"18\"/>"
              "</conditionalFormatting>"
              "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "</worksheet>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    lxlsx_worksheet_write_number(worksheet, CELL("A1"), 10, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A2"), 20, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A3"), 30, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A4"), 40, NULL);

    lxlsx_conditional_format *conditional_format = calloc(1, sizeof(lxlsx_conditional_format));

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_TOP;
    conditional_format->value    = 15;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_BOTTOM;
    conditional_format->value    = 16;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_TOP;
    conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT;
    conditional_format->value    = 17;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_BOTTOM;
    conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT;
    conditional_format->value    = 18;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);


    free(conditional_format);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
}

// Test int truncation.
CTEST(worksheet, lxlsx_worksheet_condtional_format06b) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
              "<dimension ref=\"A1:A4\"/>"
              "<sheetViews>"
                "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
              "</sheetViews>"
              "<sheetFormatPr defaultRowHeight=\"15\"/>"
              "<sheetData>"
                "<row r=\"1\" spans=\"1:1\">"
                  "<c r=\"A1\">"
                    "<v>10</v>"
                  "</c>"
                "</row>"
                "<row r=\"2\" spans=\"1:1\">"
                  "<c r=\"A2\">"
                    "<v>20</v>"
                  "</c>"
                "</row>"
                "<row r=\"3\" spans=\"1:1\">"
                  "<c r=\"A3\">"
                    "<v>30</v>"
                  "</c>"
                "</row>"
                "<row r=\"4\" spans=\"1:1\">"
                  "<c r=\"A4\">"
                    "<v>40</v>"
                  "</c>"
                "</row>"
              "</sheetData>"
              "<conditionalFormatting sqref=\"A1:A4\">"
                "<cfRule type=\"top10\" priority=\"1\" rank=\"10\"/>"
                "<cfRule type=\"top10\" priority=\"2\" bottom=\"1\" rank=\"10\"/>"
                "<cfRule type=\"top10\" priority=\"3\" percent=\"1\" rank=\"10\"/>"
                "<cfRule type=\"top10\" priority=\"4\" percent=\"1\" bottom=\"1\" rank=\"10\"/>"
              "</conditionalFormatting>"
              "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "</worksheet>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    lxlsx_worksheet_write_number(worksheet, CELL("A1"), 10, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A2"), 20, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A3"), 30, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A4"), 40, NULL);

    lxlsx_conditional_format *conditional_format = calloc(1, sizeof(lxlsx_conditional_format));

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_TOP;
    conditional_format->value    = 10.1;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_BOTTOM;
    conditional_format->value    = 10.2;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_TOP;
    conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT;
    conditional_format->value    = 10.3;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type     = LXLSX_CONDITIONAL_TYPE_BOTTOM;
    conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT;
    conditional_format->value    = 10.4;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);


    free(conditional_format);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
}

