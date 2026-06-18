/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/styles.h"

// Test the _write_font() function.
CTEST(styles, write_font01) {


    char* got;
    char exp[] = "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font02) {


    char* got;
    char exp[] = "<font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_bold(format);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font03) {


    char* got;
    char exp[] = "<font><i/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_italic(format);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font04) {


    char* got;
    char exp[] = "<font><u/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_SINGLE);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font05) {


    char* got;
    char exp[] = "<font><strike/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_strikeout(format);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font06) {


    char* got;
    char exp[] = "<font><vertAlign val=\"superscript\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_script(format, LXLSX_FONT_SUPERSCRIPT);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font07) {


    char* got;
    char exp[] = "<font><vertAlign val=\"subscript\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_script(format, LXLSX_FONT_SUBSCRIPT);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font08) {


    char* got;
    char exp[] = "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Arial\"/><family val=\"2\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_name(format, "Arial");

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font09) {


    char* got;
    char exp[] = "<font><sz val=\"12\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_size(format, 12);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font10) {


    char* got;
    char exp[] = "<font><outline/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_outline(format);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font11) {


    char* got;
    char exp[] = "<font><shadow/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_shadow(format);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font12) {


    char* got;
    char exp[] = "<font><sz val=\"11\"/><color rgb=\"FFFF0000\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_font_color(format, LXLSX_COLOR_RED);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font13) {


    char* got;
    char exp[] = "<font><b/><i/><strike/><outline/><shadow/><u/><vertAlign val=\"superscript\"/><sz val=\"12\"/><color rgb=\"FFFF0000\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_bold(format);
    lxlsx_format_set_italic(format);
    lxlsx_format_set_font_size(format, 12);
    lxlsx_format_set_font_color(format, LXLSX_COLOR_RED);
    lxlsx_format_set_font_strikeout(format);
    lxlsx_format_set_font_outline(format);
    lxlsx_format_set_font_shadow(format);
    lxlsx_format_set_font_script(format, LXLSX_FONT_SUPERSCRIPT);
    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_SINGLE);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font14) {


    char* got;
    char exp[] = "<font><u val=\"double\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_DOUBLE);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font15) {


    char* got;
    char exp[] = "<font><u val=\"singleAccounting\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_SINGLE_ACCOUNTING);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}

// Test the _write_font() function.
CTEST(styles, write_font16) {


    char* got;
    char exp[] = "<font><u val=\"doubleAccounting\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_DOUBLE_ACCOUNTING);

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}



// Test the _write_font() function.
CTEST(styles, write_font17) {


    char* got;
    char exp[] = "<font><u/><sz val=\"11\"/><color theme=\"10\"/><name val=\"Calibri\"/><family val=\"2\"/></font>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_underline(format, LXLSX_UNDERLINE_SINGLE);
    lxlsx_format_set_theme(format, 10);
    format->hyperlink = 1;

    styles->file = testfile;

    _write_font(styles, format, LXLSX_FALSE, LXLSX_FALSE);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}
