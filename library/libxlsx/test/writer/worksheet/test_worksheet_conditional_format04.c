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
CTEST(worksheet, lxlsx_worksheet_condtional_format04) {

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
                "<cfRule type=\"duplicateValues\" priority=\"1\"/>"
                "<cfRule type=\"uniqueValues\" priority=\"2\"/>"
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

    conditional_format->type = LXLSX_CONDITIONAL_TYPE_DUPLICATE;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type = LXLSX_CONDITIONAL_TYPE_UNIQUE;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);


    free(conditional_format);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
}


