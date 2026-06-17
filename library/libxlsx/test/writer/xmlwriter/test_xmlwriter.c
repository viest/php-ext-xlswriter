/*
 * Tests for xmlwriter.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/xmlwriter.h"

// Test _xml_declaration().
CTEST(xmlwriter, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_xml_declaration(testfile);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_start_tag() with no attributes.
CTEST(xmlwriter, xml_start_tag) {

    char* got;
    char exp[] = "<foo>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_xml_start_tag(testfile, "foo", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_start_tag() with attributes.
CTEST(xmlwriter, xml_start_tag_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\" baz=\"7\">";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "8");
    LXW_PUSH_ATTRIBUTES_STR("baz",  "7");

    lxw_xml_start_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_start_tag() with attributes requiring escaping.
CTEST(xmlwriter, xml_start_tag_with_attributes_to_escape) {

    char* got;
    char exp[] = "<foo span=\"&amp;&lt;&gt;&quot;\">";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxw_xml_start_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_start_tag_unencoded() with attributes.
CTEST(xmlwriter, xml_start_tag_unencoded) {

    char* got;
    char exp[] = "<foo span=\"&<>\"\">";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxw_xml_start_tag_unencoded(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_end_tag().
CTEST(xmlwriter, xml_end_tag) {

    char* got;
    char exp[] = "</foo>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_xml_end_tag(testfile, "foo");

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_empty_tag() with no attributes.
CTEST(xmlwriter, xml_empty_tag) {

    char* got;
    char exp[] = "<foo/>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_xml_empty_tag(testfile, "foo", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_empty_tag() with attributes.
CTEST(xmlwriter, xml_empty_tag_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\" baz=\"7\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "8");
    LXW_PUSH_ATTRIBUTES_STR("baz",  "7");

    lxw_xml_empty_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag() with attributes requiring escaping.
CTEST(xmlwriter, xml_empty_tag_with_attributes_to_escape) {

    char* got;
    char exp[] = "<foo span=\"&amp;&lt;&gt;&quot;\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxw_xml_empty_tag(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag_unencoded() with attributes.
CTEST(xmlwriter, xml_empty_tag_unencoded) {

    char* got;
    char exp[] = "<foo span=\"&<>\"\"/>";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "&<>\"");

    lxw_xml_empty_tag_unencoded(testfile, "foo", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_empty_tag() with no attributes.
CTEST(xmlwriter, xml_data_element) {

    char* got;
    char exp[] = "<foo>bar</foo>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_xml_data_element(testfile, "foo", "bar", NULL);

    RUN_XLSX_STREQ(exp, got);
}

// Test _xml_data_element() with attributes.
CTEST(xmlwriter, xml_data_element_with_attributes) {

    char* got;
    char exp[] = "<foo span=\"8\">bar</foo>";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "8");

    lxw_xml_data_element(testfile, "foo", "bar", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

// Test _xml_data_element() with data requiring escaping.
CTEST(xmlwriter, xml_data_element_with_escapes) {

    char* got;
    char exp[] = "<foo span=\"8\">&amp;&lt;&gt;\"</foo>";
    FILE* testfile = lxw_tmpfile(NULL);
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("span", "8");

    lxw_xml_data_element(testfile, "foo", "&<>\"", &attributes);

    RUN_XLSX_STREQ(exp, got);

    LXW_FREE_ATTRIBUTES();
}

