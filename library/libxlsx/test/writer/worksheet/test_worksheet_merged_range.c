/*
 * Tests for the lib_xlsx_writer library.
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
CTEST(merged_range, merged_range01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<dimension ref=\"B3:C3\"/>"
          "<sheetViews>"
            "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
          "</sheetViews>"
          "<sheetFormatPr defaultRowHeight=\"15\"/>"
          "<sheetData>"
            "<row r=\"3\" spans=\"2:3\">"
              "<c r=\"B3\" s=\"1\" t=\"s\">"
                "<v>0</v>"
              "</c>"
              "<c r=\"C3\" s=\"1\"/>"
            "</row>"
          "</sheetData>"
          "<mergeCells count=\"1\">"
            "<mergeCell ref=\"B3:C3\"/>"
          "</mergeCells>"
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "</worksheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;
    worksheet->sst = lxw_sst_new();
    worksheet_select(worksheet);

    lxw_format *format = lxw_format_new();
    format->xf_index = 1;

    worksheet_merge_range(worksheet, 2, 1, 2, 2, "Foo", format);

    lxw_worksheet_assemble_xml_file(worksheet);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_sst_free(worksheet->sst);
    lxw_worksheet_free(worksheet);
}
