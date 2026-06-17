/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */
#ifdef _WIN32
#define strdup _strdup
#endif

#include <string.h>
#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/drawing.h"
#include "../../../include/lxlsx/worksheet.h"

// Test assembling a complete Drawing file.
CTEST(drawing, drawing_image01) {

    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
          "<xdr:twoCellAnchor editAs=\"oneCell\">"
            "<xdr:from>"
              "<xdr:col>2</xdr:col>"
              "<xdr:colOff>0</xdr:colOff>"
              "<xdr:row>1</xdr:row>"
              "<xdr:rowOff>0</xdr:rowOff>"
            "</xdr:from>"
            "<xdr:to>"
              "<xdr:col>3</xdr:col>"
              "<xdr:colOff>533257</xdr:colOff>"
              "<xdr:row>6</xdr:row>"
              "<xdr:rowOff>190357</xdr:rowOff>"
            "</xdr:to>"
            "<xdr:pic>"
              "<xdr:nvPicPr>"
                "<xdr:cNvPr id=\"2\" name=\"Picture 1\" descr=\"republic.png\"/>"
                "<xdr:cNvPicPr>"
                  "<a:picLocks noChangeAspect=\"1\"/>"
                "</xdr:cNvPicPr>"
              "</xdr:nvPicPr>"
              "<xdr:blipFill>"
                "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId1\"/>"
                "<a:stretch>"
                  "<a:fillRect/>"
                "</a:stretch>"
              "</xdr:blipFill>"
              "<xdr:spPr>"
                "<a:xfrm>"
                  "<a:off x=\"1219200\" y=\"190500\"/>"
                  "<a:ext cx=\"1142857\" cy=\"1142857\"/>"
                "</a:xfrm>"
                "<a:prstGeom prst=\"rect\">"
                  "<a:avLst/>"
                "</a:prstGeom>"
              "</xdr:spPr>"
            "</xdr:pic>"
            "<xdr:clientData/>"
          "</xdr:twoCellAnchor>"
        "</xdr:wsDr>";

    FILE* testfile = lxw_tmpfile(NULL);

    lxw_drawing *drawing = lxw_drawing_new();
    drawing->file = testfile;

    drawing->embedded = LXW_TRUE;

    lxw_drawing_object *drawing_object = calloc(1, sizeof(lxw_drawing_object));

    drawing_object->type = LXW_DRAWING_IMAGE;
    drawing_object->anchor = LXW_OBJECT_MOVE_DONT_SIZE;

    drawing_object->from.col = 2;
    drawing_object->from.col_offset = 0;
    drawing_object->from.row = 1;
    drawing_object->from.row_offset = 0;

    drawing_object->to.col = 3;
    drawing_object->to.col_offset = 533257;
    drawing_object->to.row = 6;
    drawing_object->to.row_offset = 190357;

    drawing_object->description = strdup("republic.png");

    drawing_object->col_absolute = 1219200;
    drawing_object->row_absolute = 190500;

    drawing_object->width  = 1142857;
    drawing_object->height = 1142857;

    drawing_object->rel_index = 1;

    lxw_add_drawing_object(drawing, drawing_object);

    lxw_drawing_assemble_xml_file(drawing);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxw_drawing_free(drawing);
}

