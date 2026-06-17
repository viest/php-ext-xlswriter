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
CTEST(worksheet, lxlsx_worksheet_condtional_format22) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
              "<dimension ref=\"A1:A9\"/>"
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
              "</sheetData>"
              "<conditionalFormatting sqref=\"A1\">"
                "<cfRule type=\"iconSet\" priority=\"1\">"
                  "<iconSet iconSet=\"3Arrows\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"33\"/>"
                    "<cfvo type=\"percent\" val=\"67\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A2\">"
                "<cfRule type=\"iconSet\" priority=\"2\">"
                  "<iconSet iconSet=\"3Flags\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"33\"/>"
                    "<cfvo type=\"percent\" val=\"67\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A3\">"
                "<cfRule type=\"iconSet\" priority=\"3\">"
                  "<iconSet iconSet=\"3TrafficLights2\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"33\"/>"
                    "<cfvo type=\"percent\" val=\"67\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A4\">"
                "<cfRule type=\"iconSet\" priority=\"4\">"
                  "<iconSet iconSet=\"3Symbols\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"33\"/>"
                    "<cfvo type=\"percent\" val=\"67\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A5\">"
                "<cfRule type=\"iconSet\" priority=\"5\">"
                  "<iconSet iconSet=\"4Arrows\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"25\"/>"
                    "<cfvo type=\"percent\" val=\"50\"/>"
                    "<cfvo type=\"percent\" val=\"75\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A6\">"
                "<cfRule type=\"iconSet\" priority=\"6\">"
                  "<iconSet iconSet=\"4RedToBlack\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"25\"/>"
                    "<cfvo type=\"percent\" val=\"50\"/>"
                    "<cfvo type=\"percent\" val=\"75\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A7\">"
                "<cfRule type=\"iconSet\" priority=\"7\">"
                  "<iconSet iconSet=\"4TrafficLights\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"25\"/>"
                    "<cfvo type=\"percent\" val=\"50\"/>"
                    "<cfvo type=\"percent\" val=\"75\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A8\">"
                "<cfRule type=\"iconSet\" priority=\"8\">"
                  "<iconSet iconSet=\"5ArrowsGray\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"20\"/>"
                    "<cfvo type=\"percent\" val=\"40\"/>"
                    "<cfvo type=\"percent\" val=\"60\"/>"
                    "<cfvo type=\"percent\" val=\"80\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<conditionalFormatting sqref=\"A9\">"
                "<cfRule type=\"iconSet\" priority=\"9\">"
                  "<iconSet iconSet=\"5Quarters\">"
                    "<cfvo type=\"percent\" val=\"0\"/>"
                    "<cfvo type=\"percent\" val=\"20\"/>"
                    "<cfvo type=\"percent\" val=\"40\"/>"
                    "<cfvo type=\"percent\" val=\"60\"/>"
                    "<cfvo type=\"percent\" val=\"80\"/>"
                  "</iconSet>"
                "</cfRule>"
              "</conditionalFormatting>"
              "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "</worksheet>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;
    lxlsx_worksheet_select(worksheet);

    lxlsx_worksheet_write_number(worksheet, CELL("A1"),   1, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A2"),   2, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A3"),   3, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A4"),   4, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A5"),   5, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A6"),   6, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A7"),   7, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A8"),   8, NULL);
    lxlsx_worksheet_write_number(worksheet, CELL("A9"),   9, NULL);

    lxlsx_conditional_format *conditional_format = calloc(1, sizeof(lxlsx_conditional_format));

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_3_ARROWS_COLORED;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A1"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_3_FLAGS;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A2"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A3"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A4"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_4_ARROWS_COLORED;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A5"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_4_RED_TO_BLACK;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A6"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A7"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_5_ARROWS_GRAY;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A8"), conditional_format);

    conditional_format->type       = LXLSX_CONDITIONAL_TYPE_ICON_SETS;
    conditional_format->icon_style = LXLSX_CONDITIONAL_ICONS_5_QUARTERS;
    lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A9"), conditional_format);

    free(conditional_format);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
}


