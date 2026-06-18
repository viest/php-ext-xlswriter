/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/workbook.h"

/* Test the _write_defined_names() method. */
CTEST(workbook, write_defined_names) {


    char* got;
    char exp[] = "<definedNames><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Sheet1!$1:$1</definedName></definedNames>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    lxlsx_workbook_add_worksheet(workbook, NULL);

    _store_defined_name(workbook, "_xlnm.Print_Titles", "", "Sheet1!$1:$1", 0, 0);

    _write_defined_names(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}



/* Test the _write_defined_name() method. */
CTEST(workbook, write_defined_names_sorted) {
    char* got;
    char exp[] = "<definedNames><definedName name=\"_Egg\">Sheet1!$A$1</definedName><definedName name=\"_Fog\">Sheet1!$A$1</definedName><definedName name=\"aaa\" localSheetId=\"1\">Sheet2!$A$1</definedName><definedName name=\"Abc\">Sheet1!$A$1</definedName><definedName name=\"Bar\" localSheetId=\"2\">'Sheet 3'!$A$1</definedName><definedName name=\"Bar\" localSheetId=\"0\">Sheet1!$A$1</definedName><definedName name=\"Bar\" localSheetId=\"1\">Sheet2!$A$1</definedName><definedName name=\"Baz\">0.98</definedName><definedName name=\"car\" localSheetId=\"2\">\"Saab 900\"</definedName></definedNames>";
    FILE* testfile = lxlsx_tmpfile(NULL);


    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    lxlsx_workbook_add_worksheet(workbook, NULL);
    lxlsx_workbook_add_worksheet(workbook, NULL);
    lxlsx_workbook_add_worksheet(workbook, "Sheet 3");


    lxlsx_workbook_define_name(workbook, "'Sheet 3'!Bar", "='Sheet 3'!$A$1");
    lxlsx_workbook_define_name(workbook, "Abc",           "=Sheet1!$A$1"   );
    lxlsx_workbook_define_name(workbook, "Baz",           "=0.98"          );
    lxlsx_workbook_define_name(workbook, "Sheet1!Bar",    "=Sheet1!$A$1"   );
    lxlsx_workbook_define_name(workbook, "Sheet2!Bar",    "=Sheet2!$A$1"   );
    lxlsx_workbook_define_name(workbook, "Sheet2!aaa",    "=Sheet2!$A$1"   );
    lxlsx_workbook_define_name(workbook, "'Sheet 3'!car", "=\"Saab 900\""  );
    lxlsx_workbook_define_name(workbook, "_Egg",          "=Sheet1!$A$1"   );
    lxlsx_workbook_define_name(workbook, "_Fog",          "=Sheet1!$A$1"   );

    _write_defined_names(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}

/* Test invalid names formats. */
CTEST(workbook, write_defined_names_invalid) {
    char* got;
    char exp[] = "";
    FILE* testfile = lxlsx_tmpfile(NULL);


    lxlsx_workbook *workbook = lxlsx_workbook_new(NULL);
    workbook->file = testfile;

    lxlsx_workbook_add_worksheet(workbook, NULL);

    lxlsx_workbook_define_name(workbook, "", "=123");
    lxlsx_workbook_define_name(workbook, "Foo", "");
    lxlsx_workbook_define_name(workbook, "Sheet1!", "=123");
    lxlsx_workbook_define_name(workbook, "!", "=123");

    _write_defined_names(workbook);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_workbook_free(workbook);
}

