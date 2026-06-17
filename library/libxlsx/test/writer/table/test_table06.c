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
CTEST(worksheet, lxlsx_worksheet_table06) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"C3:F13\" totalsRowShown=\"0\">"
              "<autoFilter ref=\"C3:F13\"/>"
              "<tableColumns count=\"4\">"
                "<tableColumn id=\"1\" name=\"Foo\"/>"
                "<tableColumn id=\"2\" name=\"Column2\"/>"
                "<tableColumn id=\"3\" name=\"Column3\"/>"
                "<tableColumn id=\"4\" name=\"Baz\"/>"
              "</tableColumns>"
              "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>"
            "</table>";

    FILE* testfile = lxlsx_tmpfile(NULL);
    FILE* testfile2 = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile2;
    worksheet->sst = lxlsx_sst_new();

    lxlsx_table *table = lxlsx_table_new();
    table->file = testfile;

    lxlsx_table_column col1 = {.header = "Foo"};
    lxlsx_table_column col2 = {0};
    lxlsx_table_column col3 = {0};
    lxlsx_table_column col4 = {.header = "Baz"};
    lxlsx_table_column *columns[] = {&col1, &col2, &col3, &col4, NULL};

    lxlsx_table_options options = {.columns = columns};

    lxlsx_worksheet_add_table(worksheet, RANGE("C3:F13"), &options);

    table->lxlsx_table_obj = STAILQ_FIRST(worksheet->lxlsx_table_objs);
    table->lxlsx_table_obj->id = 1;

    lxlsx_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
    lxlsx_table_free(table);
}


