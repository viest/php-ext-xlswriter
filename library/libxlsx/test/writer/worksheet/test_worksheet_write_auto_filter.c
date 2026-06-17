/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "lxlsx/worksheet.h"


/* Test the _write_auto_filter() method with no filter. */
CTEST(worksheet, write_write_auto_filter01) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);
    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter02) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><filters><filter val=\"East\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "East"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter03) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><filters><filter val=\"East\"/><filter val=\"North\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "East"};

    lxw_filter_rule filter_rule2 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "North"};

    worksheet_filter_column2(worksheet, 0, &filter_rule1, &filter_rule2, LXW_FILTER_OR);


    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter04) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters and=\"1\"><customFilter val=\"East\"/><customFilter val=\"North\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "East"};

    lxw_filter_rule filter_rule2 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "North"};

    worksheet_filter_column2(worksheet, 0, &filter_rule1, &filter_rule2, LXW_FILTER_AND);


    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter05) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" val=\"East\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                                    .value_string = "East"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);
    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter06) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter val=\"S*\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "S*"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter07) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" val=\"S*\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                                    .value_string = "S*"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter08) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter val=\"*h\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "*h"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter09) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" val=\"*h\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                                    .value_string = "*h"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter10) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter val=\"*o*\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value_string = "*o*"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter11) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" val=\"*r*\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                                    .value_string = "*r*"};

    worksheet_filter_column(worksheet, 0, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter12) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><filters><filter val=\"1000\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_EQUAL_TO,
                                    .value        = 1000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter13) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters><customFilter operator=\"notEqual\" val=\"2000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                                    .value        = 2000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

     _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter14) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters><customFilter operator=\"greaterThan\" val=\"3000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_GREATER_THAN,
                                    .value        = 3000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

     _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter15) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters><customFilter operator=\"greaterThanOrEqual\" val=\"4000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                                    .value        = 4000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter16) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters><customFilter operator=\"lessThan\" val=\"5000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_LESS_THAN,
                                    .value        = 5000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter17) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters><customFilter operator=\"lessThanOrEqual\" val=\"6000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
                                    .value        = 6000};

    worksheet_filter_column(worksheet, 2, &filter_rule1);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter() method with a filter. */
CTEST(worksheet, write_write_auto_filter18) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"2\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"1000\"/><customFilter operator=\"lessThanOrEqual\" val=\"2000\"/></customFilters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    lxw_filter_rule filter_rule1 = {.criteria     = LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                                    .value        = 1000};

    lxw_filter_rule filter_rule2 = {.criteria     = LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
                                    .value        = 2000};

    worksheet_filter_column2(worksheet, 2, &filter_rule1, &filter_rule2, LXW_FILTER_AND);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter_list() method with a filter. */
CTEST(worksheet, write_write_auto_filter19) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><filters><filter val=\"East\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    const char* list[] = {"East", NULL};

    worksheet_filter_list(worksheet, 0, list);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter_list() method with a filter. */
CTEST(worksheet, write_write_auto_filter20) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"0\"><filters><filter val=\"East\"/><filter val=\"North\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    const char* list[] = {"East", "North", NULL};

    worksheet_filter_list(worksheet, 0, list);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* Test the _write_auto_filter_list() method with a filter. */
CTEST(worksheet, write_write_auto_filter21) {
    char* got;
    char exp[] = "<autoFilter ref=\"A1:D51\"><filterColumn colId=\"3\"><filters><filter val=\"February\"/><filter val=\"January\"/><filter val=\"July\"/><filter val=\"June\"/></filters></filterColumn></autoFilter>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_autofilter(worksheet, 0, 0, 50, 3);

    const char* list[] = {"February", "January", "July", "June", NULL};

    worksheet_filter_list(worksheet, 3, list);

    _worksheet_write_auto_filter(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}
