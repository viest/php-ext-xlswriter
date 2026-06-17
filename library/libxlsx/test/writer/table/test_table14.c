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
#include "../../../include/lxlsx/table.h"

// Test assembling a complete Table file.
CTEST(worksheet, worksheet_table14) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"B3:G7\" totalsRowShown=\"0\">"
              "<autoFilter ref=\"B3:G7\"/>"
              "<tableColumns count=\"6\">"
                "<tableColumn id=\"1\" name=\"Product\"/>"
                "<tableColumn id=\"2\" name=\"Quarter 1\"/>"
                "<tableColumn id=\"3\" name=\"Quarter 2\"/>"
                "<tableColumn id=\"4\" name=\"Quarter 3\"/>"
                "<tableColumn id=\"5\" name=\"Quarter 4\"/>"
                "<tableColumn id=\"6\" name=\"Year\">"
                  "<calculatedColumnFormula>SUM(Table1[[#This Row],[Quarter 1]:[Quarter 4]])</calculatedColumnFormula>"
                "</tableColumn>"
              "</tableColumns>"
              "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>"
            "</table>";

    FILE* testfile = lxw_tmpfile(NULL);
    FILE* testfile2 = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile2;
    worksheet->sst = lxw_sst_new();

    lxw_table *table = lxw_table_new();
    table->file = testfile;

    /* Set the table options. */
    lxw_table_column col8_1 = {.header = "Product"};
    lxw_table_column col8_2 = {.header = "Quarter 1"};
    lxw_table_column col8_3 = {.header = "Quarter 2"};
    lxw_table_column col8_4 = {.header = "Quarter 3"};
    lxw_table_column col8_5 = {.header = "Quarter 4"};
    lxw_table_column col8_6 = {.header = "Year",
                               .formula = "=SUM(Table1[@[Quarter 1]:[Quarter 4]])"};

    lxw_table_column *columns8[] = {&col8_1, &col8_2, &col8_3, &col8_4, &col8_5, &col8_6, NULL};

    lxw_table_options options8 = {.columns = columns8};

    /* Add a table to the worksheet. */
    worksheet_add_table(worksheet, RANGE("B3:G7"), &options8);

    table->table_obj = STAILQ_FIRST(worksheet->table_objs);
    table->table_obj->id = 1;

    lxw_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
    lxw_table_free(table);
}


