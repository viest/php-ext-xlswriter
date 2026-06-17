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
#include "../../../include/lxlsx/format.h"

// Test the _write_xf() method. Default properties.
CTEST(styles, write_xf01) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Has font but is first XF.
CTEST(styles, write_xf02) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    format->has_font = 1;

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Has font but isn't first XF.
CTEST(styles, write_xf03) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    format->has_font = 1;
    format->font_index = 1;

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Uses built-in number format.
CTEST(styles, write_xf04) {

    char* got;
    char exp[] = "<xf numFmtId=\"2\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_num_format_index(format, 2);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Uses built-in number format + font.
CTEST(styles, write_xf05) {

    char* got;
    char exp[] = "<xf numFmtId=\"2\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_num_format_index(format, 2);
    format->has_font = 1;
    format->font_index = 1;

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Vertical alignment = top.
CTEST(styles, write_xf06) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_TOP);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Vertical alignment = centre.
CTEST(styles, write_xf07) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"center\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_CENTER);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Vertical alignment = bottom.
CTEST(styles, write_xf08) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_BOTTOM);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Vertical alignment = justify.
CTEST(styles, write_xf09) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"justify\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_JUSTIFY);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Vertical alignment = distributed.
CTEST(styles, write_xf10) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"distributed\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_DISTRIBUTED);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = left.
CTEST(styles, write_xf11) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"left\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_LEFT);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = center.
CTEST(styles, write_xf12) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_CENTER);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = right.
CTEST(styles, write_xf13) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"right\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_RIGHT);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = left + indent.
CTEST(styles, write_xf14) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"left\" indent=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_LEFT);
    format->indent = 1;

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = right + indent.
CTEST(styles, write_xf15) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"right\" indent=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_RIGHT);
    lxlsx_format_set_indent(format, 1);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = fill.
CTEST(styles, write_xf16) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"fill\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_FILL);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = justify.
CTEST(styles, write_xf17) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"justify\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_JUSTIFY);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = center across.
CTEST(styles, write_xf18) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"centerContinuous\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_CENTER_ACROSS);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = distributed.
CTEST(styles, write_xf19) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"distributed\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_DISTRIBUTED);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = distributed + indent.
CTEST(styles, write_xf20) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"distributed\" indent=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_DISTRIBUTED);
    lxlsx_format_set_indent(format, 1);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = justify distributed.
CTEST(styles, write_xf21) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"distributed\" justifyLastLine=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_DISTRIBUTED);
    format->just_distrib = 1;

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = indent only.
// This should default to left alignment.
CTEST(styles, write_xf22) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"left\" indent=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_indent(format, 1);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Horizontal alignment = distributed + indent.
// The justify_distributed should drop back to plain distributed if there
// is an indent.
CTEST(styles, write_xf23) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"distributed\" indent=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_DISTRIBUTED);
    lxlsx_format_set_indent(format, 1);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = text wrap
CTEST(styles, write_xf24) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_text_wrap(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = shrink to fit
CTEST(styles, write_xf25) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment shrinkToFit=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_shrink(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = reading order
CTEST(styles, write_xf26) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment readingOrder=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_reading_order(format, 1);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = reading order
CTEST(styles, write_xf27) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment readingOrder=\"2\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_reading_order(format, 2);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = rotation
CTEST(styles, write_xf28) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment textRotation=\"45\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_rotation(format, 45);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = rotation
CTEST(styles, write_xf29) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment textRotation=\"135\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_rotation(format, -45);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = rotation
CTEST(styles, write_xf30) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment textRotation=\"255\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_rotation(format, 270);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = rotation
CTEST(styles, write_xf31) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment textRotation=\"90\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_rotation(format, 90);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. Alignment = rotation
CTEST(styles, write_xf32) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment textRotation=\"180\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_rotation(format, -90);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. With cell protection.
CTEST(styles, write_xf33) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyProtection=\"1\"><protection locked=\"0\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_unlocked(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. With cell protection.
CTEST(styles, write_xf34) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyProtection=\"1\"><protection hidden=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_hidden(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. With cell protection.
CTEST(styles, write_xf35) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyProtection=\"1\"><protection locked=\"0\" hidden=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_unlocked(format);
    lxlsx_format_set_hidden(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}


// Test the _write_xf() method. With cell protection + align.
CTEST(styles, write_xf36) {

    char* got;
    char exp[] = "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyProtection=\"1\"><alignment horizontal=\"right\"/><protection locked=\"0\" hidden=\"1\"/></xf>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_styles *styles = lxlsx_styles_new();
    lxlsx_format *format = lxlsx_format_new();

    lxlsx_format_set_align(format, LXLSX_ALIGN_RIGHT);
    lxlsx_format_set_unlocked(format);
    lxlsx_format_set_hidden(format);

    styles->file = testfile;

    _write_xf(styles, format);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_styles_free(styles);
    lxlsx_format_free(format);
}
