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

// Test the header and footer functions.
CTEST(worksheet, write_odd_header) {

    char* got;
    char exp[] = "<oddHeader>Page &amp;P of &amp;N</oddHeader>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_header(worksheet, "Page &P of &N");

    _worksheet_write_odd_header(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}

// Test the header and footer functions.
CTEST(worksheet, write_odd_footer) {

    char* got;
    char exp[] = "<oddFooter>&amp;F</oddFooter>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_footer(worksheet, "&F");

    _worksheet_write_odd_footer(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}


// Test the header and footer functions.
CTEST(worksheet, _worksheet_write_header_footer1) {

    char* got;
    char exp[] = "<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader></headerFooter>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_header(worksheet, "Page &P of &N");

    _worksheet_write_header_footer(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}

// Test the header and footer functions.
CTEST(worksheet, _worksheet_write_header_footer2) {

    char* got;
    char exp[] = "<headerFooter><oddFooter>&amp;F</oddFooter></headerFooter>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_footer(worksheet, "&F");

    _worksheet_write_header_footer(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}

// Test the header and footer functions.
CTEST(worksheet, _worksheet_write_header_footer3) {

    char* got;
    char exp[] = "<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader><oddFooter>&amp;F</oddFooter></headerFooter>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_worksheet *worksheet = lxlsx_worksheet_new(NULL);
    worksheet->file = testfile;

    lxlsx_worksheet_set_header(worksheet, "Page &P of &N");
    lxlsx_worksheet_set_footer(worksheet, "&F");

    _worksheet_write_header_footer(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_worksheet_free(worksheet);
}
