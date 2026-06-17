/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/drawing.h"

// Test _xml_declaration().
CTEST(drawing, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_drawing *drawing = lxlsx_drawing_new();
    drawing->file = testfile;

    _drawing_xml_declaration(drawing);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_drawing_free(drawing);
}
