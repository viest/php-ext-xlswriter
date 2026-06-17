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
CTEST(worksheet, worksheet_table09) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"B2:K8\" totalsRowCount=\"1\">"
              "<autoFilter ref=\"B2:K7\"/>"
              "<tableColumns count=\"10\">"
                "<tableColumn id=\"1\" name=\"Column1\" totalsRowLabel=\"Total\"/>"
                "<tableColumn id=\"2\" name=\"Column2\"/>"
                "<tableColumn id=\"3\" name=\"Column3\" totalsRowFunction=\"average\"/>"
                "<tableColumn id=\"4\" name=\"Column4\" totalsRowFunction=\"count\"/>"
                "<tableColumn id=\"5\" name=\"Column5\" totalsRowFunction=\"countNums\"/>"
                "<tableColumn id=\"6\" name=\"Column6\" totalsRowFunction=\"max\"/>"
                "<tableColumn id=\"7\" name=\"Column7\" totalsRowFunction=\"min\"/>"
                "<tableColumn id=\"8\" name=\"Column8\" totalsRowFunction=\"sum\"/>"
                "<tableColumn id=\"9\" name=\"Column9\" totalsRowFunction=\"stdDev\"/>"
                "<tableColumn id=\"10\" name=\"Column10\" totalsRowFunction=\"var\"/>"
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

    lxw_table_column col1  = {.total_string = "Total"};
    lxw_table_column col2  = {0};
    lxw_table_column col3  = {.total_function = LXW_TABLE_FUNCTION_AVERAGE};
    lxw_table_column col4  = {.total_function = LXW_TABLE_FUNCTION_COUNT};
    lxw_table_column col5  = {.total_function = LXW_TABLE_FUNCTION_COUNT_NUMS};
    lxw_table_column col6  = {.total_function = LXW_TABLE_FUNCTION_MAX};
    lxw_table_column col7  = {.total_function = LXW_TABLE_FUNCTION_MIN};
    lxw_table_column col8  = {.total_function = LXW_TABLE_FUNCTION_SUM};
    lxw_table_column col9  = {.total_function = LXW_TABLE_FUNCTION_STD_DEV};
    lxw_table_column col10 = {.total_function = LXW_TABLE_FUNCTION_VAR};

    lxw_table_column *columns[] = {&col1, &col2, &col3, &col4, &col5, &col6, &col7, &col8, &col9, &col10, NULL};

    lxw_table_options options = {.total_row = LXW_TRUE, .columns = columns};

    worksheet_add_table(worksheet, RANGE("B2:K8"), &options);

    table->table_obj = STAILQ_FIRST(worksheet->table_objs);
    table->table_obj->id = 1;

    lxw_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
    lxw_table_free(table);
}


