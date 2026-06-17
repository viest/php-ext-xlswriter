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
CTEST(worksheet, lxlsx_worksheet_table03) {

    char* got;
    char exp[] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"C5:D16\" totalsRowShown=\"0\">"
              "<autoFilter ref=\"C5:D16\"/>"
              "<tableColumns count=\"2\">"
                "<tableColumn id=\"1\" name=\"Column1\"/>"
                "<tableColumn id=\"2\" name=\"Column2\"/>"
              "</tableColumns>"
              "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"1\" showLastColumn=\"1\" showRowStripes=\"0\" showColumnStripes=\"1\"/>"
            "</table>";

    FILE* testfile = lxlsx_tmpfile(NULL);
    FILE* testfile2 = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile2;
    worksheet->sst = lxlsx_sst_new();

    lxlsx_table *table = lxlsx_table_new();
    table->file = testfile;

    lxlsx_table_options options = {.first_column = LXLSX_TRUE, .last_column = LXLSX_TRUE, .no_banded_rows = LXLSX_TRUE, .banded_columns = LXLSX_TRUE};

    lxlsx_worksheet_add_table(worksheet, RANGE("C5:D16"), &options);

    table->lxlsx_table_obj = STAILQ_FIRST(worksheet->lxlsx_table_objs);
    table->lxlsx_table_obj->id = 1;

    lxlsx_table_assemble_xml_file(table);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_worksheet_free(worksheet);
    lxlsx_table_free(table);
}


