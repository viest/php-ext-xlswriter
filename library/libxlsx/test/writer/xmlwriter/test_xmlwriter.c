/*
 * Tests for xmlwriter.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/xmlwriter.h"

// Test _xml_declaration().
CTEST(xmlwriter, lxlsx_xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_xml_declaration(testfile);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_start_tag() with no attributes.
CTEST(xmlwriter, lxlsx_xml_start_tag) {

    char* got;
    char exp[] = "<foo>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_xml_start_tag(testfile, "foo", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_start_tag() with attributes.
CTEST(xmlwriter, lxlsx_xml_start_tag_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\" baz=\"7\">";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "8");
    LXLSX_PUSH_ATTRIBUTES_STR("baz",  "7");

    lxlsx_xml_start_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_start_tag() with attributes requiring escaping.
CTEST(xmlwriter, lxlsx_xml_start_tag_with_attributes_to_escape) {

    char* got;
    char exp[] = "<foo span=\"&amp;&lt;&gt;&quot;\">";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxlsx_xml_start_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_start_tag_unencoded() with attributes.
CTEST(xmlwriter, lxlsx_xml_start_tag_unencoded) {

    char* got;
    char exp[] = "<foo span=\"&<>\"\">";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxlsx_xml_start_tag_unencoded(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_end_tag().
CTEST(xmlwriter, lxlsx_xml_end_tag) {

    char* got;
    char exp[] = "</foo>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_xml_end_tag(testfile, "foo");

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_empty_tag() with no attributes.
CTEST(xmlwriter, lxlsx_xml_empty_tag) {

    char* got;
    char exp[] = "<foo/>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_xml_empty_tag(testfile, "foo", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_empty_tag() with attributes.
CTEST(xmlwriter, lxlsx_xml_empty_tag_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\" baz=\"7\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "8");
    LXLSX_PUSH_ATTRIBUTES_STR("baz",  "7");

    lxlsx_xml_empty_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag() with attributes requiring escaping.
CTEST(xmlwriter, lxlsx_xml_empty_tag_with_attributes_to_escape) {

    char* got;
    char exp[] = "<foo span=\"&amp;&lt;&gt;&quot;\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxlsx_xml_empty_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag_unencoded() with attributes.
CTEST(xmlwriter, lxlsx_xml_empty_tag_unencoded) {

    char* got;
    char exp[] = "<foo span=\"&<>\"\"/>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxlsx_xml_empty_tag_unencoded(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag() with no attributes.
CTEST(xmlwriter, lxlsx_xml_data_element) {

    char* got;
    char exp[] = "<foo>bar</foo>";
    FILE* testfile = lxlsx_tmpfile(NULL);

    lxlsx_xml_data_element(testfile, "foo", "bar", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_data_element() with attributes.
CTEST(xmlwriter, lxlsx_xml_data_element_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\">bar</foo>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "8");

    lxlsx_xml_data_element(testfile, "foo", "bar", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

// Test _xml_data_element() with data requiring escaping.
CTEST(xmlwriter, lxlsx_xml_data_element_with_escapes) {

    char* got;
    char exp[] = "<foo span=\"8\">&amp;&lt;&gt;\"</foo>";
    FILE* testfile = lxlsx_tmpfile(NULL);
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("span", "8");

    lxlsx_xml_data_element(testfile, "foo", "&<>\"", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXLSX_FREE_ATTRIBUTES();
}

