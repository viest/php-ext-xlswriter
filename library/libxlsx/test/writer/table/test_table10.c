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
CTEST(worksheet, worksheet_table10) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"MyTable\" displayName=\"MyTable\" ref=\"C2:F13\" totalsRowShown=\"0\">"
              "<autoFilter ref=\"C2:F13\"/>"
              "<tableColumns count=\"4\">"
                "<tableColumn id=\"1\" name=\"Column1\"/>"
                "<tableColumn id=\"2\" name=\"Column2\"/>"
                "<tableColumn id=\"3\" name=\"Column3\"/>"
                "<tableColumn id=\"4\" name=\"Column4\"/>"
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

    lxw_table_options options = {.name = "MyTable"};

    worksheet_add_table(worksheet, RANGE("C2:F13"), &options);

    table->table_obj = STAILQ_FIRST(worksheet->table_objs);
    table->table_obj->id = 1;

    lxw_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_worksheet_free(worksheet);
    lxw_table_free(table);
}


