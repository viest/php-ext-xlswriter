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
CTEST(worksheet, lxlsx_worksheet_table02) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"D4:I15\" totalsRowShown=\"0\">"
              "<autoFilter ref=\"D4:I15\"/>"
              "<tableColumns count=\"6\">"
                "<tableColumn id=\"1\" name=\"Column1\"/>"
                "<tableColumn id=\"2\" name=\"Column2\"/>"
                "<tableColumn id=\"3\" name=\"Column3\"/>"
                "<tableColumn id=\"4\" name=\"Column4\"/>"
                "<tableColumn id=\"5\" name=\"Column5\"/>"
                "<tableColumn id=\"6\" name=\"Column6\"/>"
              "</tableColumns>"
              "<tableStyleInfo name=\"TableStyleLight17\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>"
            "</table>";

    FILE* testfile = lxlsx_tmpfile(NULL);
    FILE* testfile2 = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile2;
    worksheet->sst = lxlsx_sst_new();

    lxlsx_table *table = lxlsx_table_new();
    table->file = testfile;

    lxlsx_table_options options = {.style_type = LXLSX_TABLE_STYLE_TYPE_LIGHT, .style_type_number = 17};

    lxlsx_worksheet_add_table(worksheet, RANGE("D4:I15"), &options);

    table->lxlsx_table_obj = STAILQ_FIRST(worksheet->lxlsx_table_objs);
    table->lxlsx_table_obj->id = 1;

    lxlsx_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
    lxlsx_table_free(table);
}


