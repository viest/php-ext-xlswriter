/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/chartsheet.h"


/* 1. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection01) {
    char* got;
    char exp[] = "<sheetProtection content=\"1\" objects=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    chartsheet_protect(chartsheet, NULL, NULL);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}


/* 2. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection02) {
    char* got;
    char exp[] = "<sheetProtection password=\"83AF\" content=\"1\" objects=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    chartsheet_protect(chartsheet, "password", NULL);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}

/* 3. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection03) {
    char* got;
    char exp[] = "<sheetProtection content=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    lxw_protection options = {.no_objects = 1};

    chartsheet_protect(chartsheet, NULL, &options);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}


/* 4. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection04) {
    char* got;
    char exp[] = "<sheetProtection objects=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    lxw_protection options = {.no_content = 1};

    chartsheet_protect(chartsheet, NULL, &options);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}


/* 5. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection05) {
    char* got;
    char exp[] = "";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    lxw_protection options = {.no_content = 1, .no_objects = 1};

    chartsheet_protect(chartsheet, NULL, &options);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}


/* 6. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection06) {
    char* got;
    char exp[] = "<sheetProtection password=\"83AF\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    lxw_protection options = {.no_content = 1, .no_objects = 1};

    chartsheet_protect(chartsheet, "password", &options);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}


/* 7. Test the _write_sheet_protection() method. */
CTEST(chartsheet, write_write_sheet_protection07) {
    char* got;
    char exp[] = "<sheetProtection password=\"83AF\" content=\"1\" objects=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_chartsheet *chartsheet = lxw_chartsheet_new(NULL);
    chartsheet->file = testfile;
    chartsheet->worksheet->file = testfile;

    lxw_protection options = {
        .objects                  = 1,
        .scenarios                = 1,
        .format_cells             = 1,
        .format_columns           = 1,
        .format_rows              = 1,
        .insert_columns           = 1,
        .insert_rows              = 1,
        .insert_hyperlinks        = 1,
        .delete_columns           = 1,
        .delete_rows              = 1,
        .no_select_locked_cells   = 1,
        .sort                     = 1,
        .autofilter               = 1,
        .pivot_tables             = 1,
        .no_select_unlocked_cells = 1,
    };

    chartsheet_protect(chartsheet, "password", &options);
    _chartsheet_write_sheet_protection(chartsheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_chartsheet_free(chartsheet);
}
