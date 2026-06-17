/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/comment.h"

// Test _xml_declaration().
CTEST(comment, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = tmpfile();

    lxlsx_comment *comment = lxlsx_comment_new();
    comment->file = testfile;

    _comment_xml_declaration(comment);

    RUN_XLSX_STREQ(exp, got);

    lxlsx_comment_free(comment);
}
