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


/* 1. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection01) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_protect(worksheet, NULL, NULL);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 2. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection02) {
    char* got;
    char exp[] = "<sheetProtection password=\"83AF\" sheet=\"1\" objects=\"1\" scenarios=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_protect(worksheet, "password", NULL);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}

/* 3. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection03) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" selectLockedCells=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.no_select_locked_cells = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 4. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection04) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" formatCells=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.format_cells = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 5. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection05) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" formatColumns=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.format_columns = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 6. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection06) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" formatRows=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.format_rows = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 7. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection07) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" insertColumns=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.insert_columns = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 8. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection08) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" insertRows=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.insert_rows = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 9. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection09) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" insertHyperlinks=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.insert_hyperlinks = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 10. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection10) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" deleteColumns=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.delete_columns = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 11. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection11) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" deleteRows=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.delete_rows = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 12. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection12) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" sort=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.sort = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 13. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection13) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" autoFilter=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.autofilter = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 14. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection14) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" pivotTables=\"0\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.pivot_tables = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 15. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection15) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" scenarios=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.objects = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 16. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection16) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {.scenarios = 1};

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 17. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection17) {
    char* got;
    char exp[] = "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\" formatCells=\"0\" selectLockedCells=\"1\" selectUnlockedCells=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    lxw_protection options = {
        .format_cells             = 1,
        .no_select_locked_cells   = 1,
        .no_select_unlocked_cells = 1,
    };

    worksheet_protect(worksheet, NULL, &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}


/* 18. Test the _write_sheet_protection() method. */
CTEST(worksheet, write_write_sheet_protection18) {
    char* got;
    char exp[] = "<sheetProtection password=\"996B\" sheet=\"1\" formatCells=\"0\" formatColumns=\"0\" formatRows=\"0\" insertColumns=\"0\" insertRows=\"0\" insertHyperlinks=\"0\" deleteColumns=\"0\" deleteRows=\"0\" selectLockedCells=\"1\" sort=\"0\" autoFilter=\"0\" pivotTables=\"0\" selectUnlockedCells=\"1\"/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

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

    worksheet_protect(worksheet, "drowssap", &options);
    _worksheet_write_sheet_protection(worksheet, &worksheet->protection);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
}
