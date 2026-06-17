/*
 * Tests for the libxlsxwriter library.
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
CTEST(worksheet, worksheet_condtional_format14) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
              "<dimension ref=\"A1:A12\"/>"
              "<sheetViews>"
                "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
              "</sheetViews>"
              "<sheetFormatPr defaultRowHeight=\"15\"/>"
              "<sheetData>"
                "<row r=\"1\" spans=\"1:1\">"
                  "<c r=\"A1\">"
                    "<v>1</v>"
                  "</c>"
                "</row>"
                "<row r=\"2\" spans=\"1:1\">"
                  "<c r=\"A2\">"
                    "<v>2</v>"
                  "</c>"
                "</row>"
                "<row r=\"3\" spans=\"1:1\">"
                  "<c r=\"A3\">"
                    "<v>3</v>"
                  "</c>"
                "</row>"
                "<row r=\"4\" spans=\"1:1\">"
                  "<c r=\"A4\">"
                    "<v>4</v>"
                  "</c>"
                "</row>"
                "<row r=\"5\" spans=\"1:1\">"
                  "<c r=\"A5\">"
                    "<v>5</v>"
                  "</c>"
                "</row>"
                "<row r=\"6\" spans=\"1:1\">"
                  "<c r=\"A6\">"
                    "<v>6</v>"
                  "</c>"
                "</row>"
                "<row r=\"7\" spans=\"1:1\">"
                  "<c r=\"A7\">"
                    "<v>7</v>"
                  "</c>"
                "</row>"
                "<row r=\"8\" spans=\"1:1\">"
                  "<c r=\"A8\">"
                    "<v>8</v>"
                  "</c>"
                "</row>"
                "<row r=\"9\" spans=\"1:1\">"
                  "<c r=\"A9\">"
                    "<v>9</v>"
                  "</c>"
                "</row>"
                "<row r=\"10\" spans=\"1:1\">"
                  "<c r=\"A10\">"
                    "<v>10</v>"
                  "</c>"
                "</row>"
                "<row r=\"11\" spans=\"1:1\">"
                  "<c r=\"A11\">"
                    "<v>11</v>"
                  "</c>"
                "</row>"
                "<row r=\"12\" spans=\"1:1\">"
                  "<c r=\"A12\">"
                    "<v>12</v>"
                  "</c>"
                "</row>"
              "</sheetData>"
              "<conditionalFormatting sqref=\"A1:A12\">"
                "<cfRule type=\"dataBar\" priority=\"1\">"
                  "<dataBar>"
                    "<cfvo type=\"min\" val=\"0\"/>"
                    "<cfvo type=\"max\" val=\"0\"/>"
                    "<color rgb=\"FF638EC6\"/>"
                  "</dataBar>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet_select(worksheet);

    worksheet_write_number(worksheet, CELL("A1"),   1, NULL);
    worksheet_write_number(worksheet, CELL("A2"),   2, NULL);
    worksheet_write_number(worksheet, CELL("A3"),   3, NULL);
    worksheet_write_number(worksheet, CELL("A4"),   4, NULL);
    worksheet_write_number(worksheet, CELL("A5"),   5, NULL);
    worksheet_write_number(worksheet, CELL("A6"),   6, NULL);
    worksheet_write_number(worksheet, CELL("A7"),   7, NULL);
    worksheet_write_number(worksheet, CELL("A8"),   8, NULL);
    worksheet_write_number(worksheet, CELL("A9"),   9, NULL);
    worksheet_write_number(worksheet, CELL("A10"), 10, NULL);
    worksheet_write_number(worksheet, CELL("A11"), 11, NULL);
    worksheet_write_number(worksheet, CELL("A12"), 12, NULL);

    lxw_conditional_format *conditional_format = calloc(1, sizeof(lxw_conditional_format));

    conditional_format->type     = LXW_CONDITIONAL_DATA_BAR;
    worksheet_conditional_format_range(worksheet, RANGE("A1:A12"), conditional_format);


    free(conditional_format);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
}
