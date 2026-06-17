/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/styles.h"

// Test assembling a complete Styles file.
CTEST(styles, styles01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
          "<fonts count=\"1\">"
            "<font>"
              "<sz val=\"11\"/>"
              "<color theme=\"1\"/>"
              "<name val=\"Calibri\"/>"
              "<family val=\"2\"/>"
              "<scheme val=\"minor\"/>"
            "</font>"
          "</fonts>"
          "<fills count=\"2\">"
            "<fill>"
              "<patternFill patternType=\"none\"/>"
            "</fill>"
            "<fill>"
              "<patternFill patternType=\"gray125\"/>"
            "</fill>"
          "</fills>"
          "<borders count=\"1\">"
            "<border>"
              "<left/>"
              "<right/>"
              "<top/>"
              "<bottom/>"
              "<diagonal/>"
            "</border>"
          "</borders>"
          "<cellStyleXfs count=\"1\">"
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
          "</cellStyleXfs>"
          "<cellXfs count=\"1\">"
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
          "</cellXfs>"
          "<cellStyles count=\"1\">"
            "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
          "</cellStyles>"
          "<dxfs count=\"0\"/>"
          "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        "</styleSheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    lxw_format *format = lxw_format_new();

    format->has_font   = 1;
    format->has_border = 1;

    styles->file = testfile;
    styles->font_count = 1;
    styles->border_count = 1;
    styles->fill_count = 2;
    styles->xf_count = 1;

    STAILQ_INSERT_TAIL(styles->xf_formats, format, list_pointers);

    lxw_styles_assemble_xml_file(styles);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_styles_free(styles);
}

// Test assembling a complete Styles file.
CTEST(styles, styles02) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
          "<fonts count=\"4\">"
            "<font>"
              "<sz val=\"11\"/>"
              "<color theme=\"1\"/>"
              "<name val=\"Calibri\"/>"
              "<family val=\"2\"/>"
              "<scheme val=\"minor\"/>"
            "</font>"
            "<font>"
              "<b/>"
              "<sz val=\"11\"/>"
              "<color theme=\"1\"/>"
              "<name val=\"Calibri\"/>"
              "<family val=\"2\"/>"
              "<scheme val=\"minor\"/>"
            "</font>"
            "<font>"
              "<i/>"
              "<sz val=\"11\"/>"
              "<color theme=\"1\"/>"
              "<name val=\"Calibri\"/>"
              "<family val=\"2\"/>"
              "<scheme val=\"minor\"/>"
            "</font>"
            "<font>"
              "<b/>"
              "<i/>"
              "<sz val=\"11\"/>"
              "<color theme=\"1\"/>"
              "<name val=\"Calibri\"/>"
              "<family val=\"2\"/>"
              "<scheme val=\"minor\"/>"
            "</font>"
          "</fonts>"
          "<fills count=\"2\">"
            "<fill>"
              "<patternFill patternType=\"none\"/>"
            "</fill>"
            "<fill>"
              "<patternFill patternType=\"gray125\"/>"
            "</fill>"
          "</fills>"
          "<borders count=\"1\">"
            "<border>"
              "<left/>"
              "<right/>"
              "<top/>"
              "<bottom/>"
              "<diagonal/>"
            "</border>"
          "</borders>"
          "<cellStyleXfs count=\"1\">"
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
          "</cellStyleXfs>"
          "<cellXfs count=\"4\">"
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
            "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
            "<xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
          "</cellXfs>"
          "<cellStyles count=\"1\">"
            "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
          "</cellStyles>"
          "<dxfs count=\"0\"/>"
          "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        "</styleSheet>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_styles *styles = lxw_styles_new();
    lxw_format *format1 = lxw_format_new();
    lxw_format *format2 = lxw_format_new();
    lxw_format *format3 = lxw_format_new();
    lxw_format *format4 = lxw_format_new();


    format1->has_font = 1;
    format1->has_border = 1;

    format2->bold = 1;
    format2->font_index = 1;
    format2->has_font = 1;

    format3->italic = 1;
    format3->font_index = 2;
    format3->has_font = 1;

    format4->bold = 1;
    format4->italic = 1;
    format4->font_index = 3;
    format4->has_font = 1;

    styles->file = testfile;
    styles->font_count = 4;
    styles->border_count = 1;
    styles->fill_count = 2;
    styles->xf_count = 4;

    STAILQ_INSERT_TAIL(styles->xf_formats, format1, list_pointers);
    STAILQ_INSERT_TAIL(styles->xf_formats, format2, list_pointers);
    STAILQ_INSERT_TAIL(styles->xf_formats, format3, list_pointers);
    STAILQ_INSERT_TAIL(styles->xf_formats, format4, list_pointers);

    lxw_styles_assemble_xml_file(styles);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_styles_free(styles);
}
