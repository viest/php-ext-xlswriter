/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/chart.h"

// Test the _write_legend() function.
CTEST(chart, write_legend01) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}

CTEST(chart, write_legend02) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_RIGHT);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}

CTEST(chart, write_legend03) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"t\"/><c:layout/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_TOP);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}

CTEST(chart, write_legend04) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"l\"/><c:layout/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_LEFT);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}

CTEST(chart, write_legend05) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"b\"/><c:layout/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}


CTEST(chart, write_legend06) {

    char* got;
    char exp[] = "";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}


CTEST(chart, write_legend07) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"1\"/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_OVERLAY_RIGHT);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}


CTEST(chart, write_legend08) {

    char* got;
    char exp[] = "<c:legend><c:legendPos val=\"l\"/><c:layout/><c:overlay val=\"1\"/></c:legend>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chart *chart = lxw_chart_new(LXW_CHART_AREA);
    chart->file = testfile;

    chart_legend_set_position(chart, LXW_CHART_LEGEND_OVERLAY_LEFT);
    _chart_write_legend(chart);

    RUN_XLSX_STREQ(exp, got);

    lxw_chart_free(chart);
}
