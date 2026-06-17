/*
 * Tests for the lib_xlsx_writer library.
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
CTEST(merged_range, array_formula01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<dimension ref=\"A1:C7\"/>"
          "<sheetViews>"
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
          "</sheetViews>"
          "<sheetFormatPr defaultRowHeight=\"15\"/>"
          "<sheetData>"
            "<row r=\"1\" spans=\"1:3\">"
              "<c r=\"A1\">"
                "<f t=\"array\" ref=\"A1\">SUM(B1:C1*B2:C2)</f>"
                "<v>9500</v>"
              "</c>"
              "<c r=\"B1\">"
                "<v>500</v>"
              "</c>"
              "<c r=\"C1\">"
                "<v>300</v>"
              "</c>"
            "</row>"
            "<row r=\"2\" spans=\"1:3\">"
              "<c r=\"A2\">"
                "<f t=\"array\" ref=\"A2\">SUM(B1:C1*B2:C2)</f>"
                "<v>9500</v>"
              "</c>"
              "<c r=\"B2\">"
                "<v>10</v>"
              "</c>"
              "<c r=\"C2\">"
                "<v>15</v>"
              "</c>"
            "</row>"
            "<row r=\"5\" spans=\"1:3\">"
              "<c r=\"A5\">"
                "<f t=\"array\" ref=\"A5:A7\">TREND(C5:C7,B5:B7)</f>"
                "<v>22196</v>"
              "</c>"
              "<c r=\"B5\">"
                "<v>1</v>"
              "</c>"
              "<c r=\"C5\">"
                "<v>20234</v>"
              "</c>"
            "</row>"
            "<row r=\"6\" spans=\"1:3\">"
              "<c r=\"A6\">"
                "<v>0</v>"
              "</c>"
              "<c r=\"B6\">"
                "<v>2</v>"
              "</c>"
              "<c r=\"C6\">"
                "<v>21003</v>"
              "</c>"
            "</row>"
            "<row r=\"7\" spans=\"1:3\">"
              "<c r=\"A7\">"
                "<v>0</v>"
              "</c>"
              "<c r=\"B7\">"
                "<v>3</v>"
              "</c>"
              "<c r=\"C7\">"
                "<v>10000</v>"
              "</c>"
            "</row>"
          "</sheetData>"
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "</worksheet>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    //worksheet->sst = _new_sst();
    lxlsx_worksheet_select(worksheet);

    lxlsx_format *format = lxlsx_format_new();
    format->xf_index = 1;

    lxlsx_worksheet_write_array_formula_num(worksheet, 0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", NULL, 9500);
    lxlsx_worksheet_write_array_formula_num(worksheet, 1, 0, 1, 0, "{=SUM(B1:C1*B2:C2)}", NULL, 9500);
    lxlsx_worksheet_write_array_formula_num(worksheet, 4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", NULL, 22196);

    lxlsx_worksheet_write_number(worksheet, 0, 1, 500, NULL);
    lxlsx_worksheet_write_number(worksheet, 1, 1, 10, NULL);
    lxlsx_worksheet_write_number(worksheet, 4, 1, 1, NULL);
    lxlsx_worksheet_write_number(worksheet, 5, 1, 2, NULL);
    lxlsx_worksheet_write_number(worksheet, 6, 1, 3, NULL);

    lxlsx_worksheet_write_number(worksheet, 0, 2, 300, NULL);
    lxlsx_worksheet_write_number(worksheet, 1, 2, 15, NULL);
    lxlsx_worksheet_write_number(worksheet, 4, 2, 20234, NULL);
    lxlsx_worksheet_write_number(worksheet, 5, 2, 21003, NULL);
    lxlsx_worksheet_write_number(worksheet, 6, 2, 10000, NULL);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    //_free_sst(worksheet->sst);
    lxlsx_worksheet_free(worksheet);
}
