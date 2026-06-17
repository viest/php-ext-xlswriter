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
CTEST(worksheet, lxlsx_worksheet_condtional_format08) {

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
                "<cfRule type=\"timePeriod\" priority=\"1\" timePeriod=\"yesterday\">"
                  "<formula>FLOOR(A1,1)=TODAY()-1</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"2\" timePeriod=\"today\">"
                  "<formula>FLOOR(A1,1)=TODAY()</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"3\" timePeriod=\"tomorrow\">"
                  "<formula>FLOOR(A1,1)=TODAY()+1</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"4\" timePeriod=\"last7Days\">"
                  "<formula>AND(TODAY()-FLOOR(A1,1)&lt;=6,FLOOR(A1,1)&lt;=TODAY())</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"5\" timePeriod=\"lastWeek\">"
                  "<formula>AND(TODAY()-ROUNDDOWN(A1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)&lt;(WEEKDAY(TODAY())+7))</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"6\" timePeriod=\"thisWeek\">"
                  "<formula>AND(TODAY()-ROUNDDOWN(A1,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"7\" timePeriod=\"nextWeek\">"
                  "<formula>AND(ROUNDDOWN(A1,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"8\" timePeriod=\"lastMonth\">"
                  "<formula>AND(MONTH(A1)=MONTH(TODAY())-1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=1,YEAR(A1)=YEAR(TODAY())-1)))</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"9\" timePeriod=\"thisMonth\">"
                  "<formula>AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))</formula>"
                "</cfRule>"
                "<cfRule type=\"timePeriod\" priority=\"10\" timePeriod=\"nextMonth\">"
                  "<formula>AND(MONTH(A1)=MONTH(TODAY())+1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=12,YEAR(A1)=YEAR(TODAY())+1)))</formula>"
                "</cfRule>"
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

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);

    conditional_format->type         = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD;
    conditional_format->criteria     = LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH;
    lxlsx_worksheet_conditional_format_range(worksheet, RANGE("A1:A4"), conditional_format);


    free(conditional_format);

    lxlsx_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
}
