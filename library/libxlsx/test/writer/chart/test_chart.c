/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/chart.h"

// Test assembling a complete Chart file.
CTEST(chart, chart01) {

    lxlsx_chart_series *series1;
    lxlsx_chart_series *series2;

    uint8_t data[5][3] = {
        {1, 2,  3},
        {2, 4,  6},
        {3, 6,  9},
        {4, 8,  12},
        {5, 10, 15}
    };



    char* got;
    char exp[] =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          "<c:lang val=\"en-US\"/>"
          "<c:chart>"
            "<c:plotArea>"
              "<c:layout/>"
              "<c:barChart>"
                "<c:barDir val=\"bar\"/>"
                "<c:grouping val=\"clustered\"/>"
                "<c:ser>"
                  "<c:idx val=\"0\"/>"
                  "<c:order val=\"0\"/>"
                  "<c:val>"
                    "<c:numRef>"
                      "<c:f>Sheet1!$A$1:$A$5</c:f>"
                      "<c:numCache>"
                        "<c:formatCode>General</c:formatCode>"
                        "<c:ptCount val=\"5\"/>"
                        "<c:pt idx=\"0\">"
                          "<c:v>1</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"1\">"
                          "<c:v>2</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"2\">"
                          "<c:v>3</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"3\">"
                          "<c:v>4</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"4\">"
                          "<c:v>5</c:v>"
                        "</c:pt>"
                      "</c:numCache>"
                    "</c:numRef>"
                  "</c:val>"
                "</c:ser>"
                "<c:ser>"
                  "<c:idx val=\"1\"/>"
                  "<c:order val=\"1\"/>"
                  "<c:val>"
                    "<c:numRef>"
                      "<c:f>Sheet1!$B$1:$B$5</c:f>"
                      "<c:numCache>"
                        "<c:formatCode>General</c:formatCode>"
                        "<c:ptCount val=\"5\"/>"
                        "<c:pt idx=\"0\">"
                          "<c:v>2</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"1\">"
                          "<c:v>4</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"2\">"
                          "<c:v>6</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"3\">"
                          "<c:v>8</c:v>"
                        "</c:pt>"
                        "<c:pt idx=\"4\">"
                          "<c:v>10</c:v>"
                        "</c:pt>"
                      "</c:numCache>"
                    "</c:numRef>"
                  "</c:val>"
                "</c:ser>"
                "<c:axId val=\"50010001\"/>"
                "<c:axId val=\"50010002\"/>"
              "</c:barChart>"
              "<c:catAx>"
                "<c:axId val=\"50010001\"/>"
                "<c:scaling>"
                  "<c:orientation val=\"minMax\"/>"
                "</c:scaling>"
                "<c:axPos val=\"l\"/>"
                "<c:tickLblPos val=\"nextTo\"/>"
                "<c:crossAx val=\"50010002\"/>"
                "<c:crosses val=\"autoZero\"/>"
                "<c:auto val=\"1\"/>"
                "<c:lblAlgn val=\"ctr\"/>"
                "<c:lblOffset val=\"100\"/>"
              "</c:catAx>"
              "<c:valAx>"
                "<c:axId val=\"50010002\"/>"
                "<c:scaling>"
                  "<c:orientation val=\"minMax\"/>"
                "</c:scaling>"
                "<c:axPos val=\"b\"/>"
                "<c:majorGridlines/>"
                "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>"
                "<c:tickLblPos val=\"nextTo\"/>"
                "<c:crossAx val=\"50010001\"/>"
                "<c:crosses val=\"autoZero\"/>"
                "<c:crossBetween val=\"between\"/>"
              "</c:valAx>"
            "</c:plotArea>"
            "<c:legend>"
              "<c:legendPos val=\"r\"/>"
              "<c:layout/>"
            "</c:legend>"
            "<c:plotVisOnly val=\"1\"/>"
          "</c:chart>"
          "<c:printSettings>"
            "<c:headerFooter/>"
            "<c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
            "<c:pageSetup/>"
          "</c:printSettings>"
        "</c:chartSpace>";

    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_chart *chart = lxlsx_chart_new(LXLSX_CHART_BAR);
    chart->file = testfile;

    series1 = lxlsx_chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
    series2 = lxlsx_chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5");

    lxlsx_chart_add_data_cache(series1->values, data[0], 5, 3, 0);
    lxlsx_chart_add_data_cache(series2->values, data[0], 5, 3, 1);

    lxlsx_chart_assemble_xml_file(chart);

    RUN_XLSX_STREQ_SHORT(exp, got);

    lxlsx_chart_free(chart);
}

