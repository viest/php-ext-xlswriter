/*****************************************************************************
 * xmlwriter - A base library for libxlsxwriter libraries.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include "xlsxwriter/xmlwriter.h"

#define LXW_AMP  "&amp;"
#define LXW_LT   "&lt;"
#define LXW_GT   "&gt;"
#define LXW_QUOT "&quot;"

/* Defines. */
#define LXW_MAX_ENCODED_ATTRIBUTE_LENGTH (LXW_MAX_ATTRIBUTE_LENGTH*6)

/* Forward declarations. */
STATIC char *_escape_attributes(struct xml_attribute *attribute);

char *lxw_escape_data(const char *data);

STATIC void _fprint_escaped_attributes(FILE * xmlfile,
                                       struct xml_attribute_list *attributes);

STATIC void _fprint_escaped_data(FILE * xmlfile, const char *data);

/*
 * Write the XML declaration.
 */
void
lxw_xml_declaration(FILE * xmlfile)
{
    fprintf(xmlfile, "<?xml version=\"1.0\" "
            "encoding=\"UTF-8\" standalone=\"yes\"?>\n");
}

/*
 * Write an XML start tag with optional attributes.
 */
void
lxw_xml_start_tag(FILE * xmlfile,
                  const char *tag, struct xml_attribute_list *attributes)
{
    fprintf(xmlfile, "<%s", tag);

    _fprint_escaped_attributes(xmlfile, attributes);

    fprintf(xmlfile, ">");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void
lxw_xml_start_tag_unencoded(FILE * xmlfile,
                            const char *tag,
                            struct xml_attribute_list *attributes)
{
    struct xml_attribute *attribute;

    fprintf(xmlfile, "<%s", tag);

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(xmlfile, " %s=\"%s\"", attribute->key, attribute->value);
        }
    }

    fprintf(xmlfile, ">");
}

/*
 * Write an XML end tag.
 */
void
lxw_xml_end_tag(FILE * xmlfile, const char *tag)
{
    fprintf(xmlfile, "</%s>", tag);
}

/*
 * Write an empty XML tag with optional attributes.
 */
void
lxw_xml_empty_tag(FILE * xmlfile,
                  const char *tag, struct xml_attribute_list *attributes)
{
    fprintf(xmlfile, "<%s", tag);

    _fprint_escaped_attributes(xmlfile, attributes);

    fprintf(xmlfile, "/>");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void
lxw_xml_empty_tag_unencoded(FILE * xmlfile,
                            const char *tag,
                            struct xml_attribute_list *attributes)
{
    struct xml_attribute *attribute;

    fprintf(xmlfile, "<%s", tag);

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(xmlfile, " %s=\"%s\"", attribute->key, attribute->value);
        }
    }

    fprintf(xmlfile, "/>");
}

/*
 * Write an XML element containing data with optional attributes.
 */
void
lxw_xml_data_element(FILE * xmlfile,
                     const char *tag,
                     const char *data, struct xml_attribute_list *attributes)
{
    fprintf(xmlfile, "<%s", tag);

    _fprint_escaped_attributes(xmlfile, attributes);

    fprintf(xmlfile, ">");

    _fprint_escaped_data(xmlfile, data);

    fprintf(xmlfile, "</%s>", tag);
}

/*
 * Escape XML characters in attributes.
 */
STATIC char *
_escape_attributes(struct xml_attribute *attribute)
{
    char *encoded = (char *) calloc(LXW_MAX_ENCODED_ATTRIBUTE_LENGTH, 1);
    char *p_encoded = encoded;
    char *p_attr = attribute->value;

    while (*p_attr) {
        switch (*p_attr) {
            case '&':
                strncat(p_encoded, LXW_AMP, sizeof(LXW_AMP) - 1);
                p_encoded += sizeof(LXW_AMP) - 1;
                break;
            case '<':
                strncat(p_encoded, LXW_LT, sizeof(LXW_LT) - 1);
                p_encoded += sizeof(LXW_LT) - 1;
                break;
            case '>':
                strncat(p_encoded, LXW_GT, sizeof(LXW_GT) - 1);
                p_encoded += sizeof(LXW_GT) - 1;
                break;
            case '"':
                strncat(p_encoded, LXW_QUOT, sizeof(LXW_QUOT) - 1);
                p_encoded += sizeof(LXW_QUOT) - 1;
                break;
            default:
                *p_encoded = *p_attr;
                p_encoded++;
                break;
        }
        p_attr++;
    }

    return encoded;
}

/*
 * Escape XML characters in data sections of tags.
 * Note, this is different from _escape_attributes()
 * in that double quotes are not escaped by Excel.
 */
char *
lxw_escape_data(const char *data)
{
    size_t encoded_len = (strlen(data) * 5 + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*data) {
        switch (*data) {
            case '&':
                strncat(p_encoded, LXW_AMP, sizeof(LXW_AMP) - 1);
                p_encoded += sizeof(LXW_AMP) - 1;
                break;
            case '<':
                strncat(p_encoded, LXW_LT, sizeof(LXW_LT) - 1);
                p_encoded += sizeof(LXW_LT) - 1;
                break;
            case '>':
                strncat(p_encoded, LXW_GT, sizeof(LXW_GT) - 1);
                p_encoded += sizeof(LXW_GT) - 1;
                break;
            default:
                *p_encoded = *data;
                p_encoded++;
                break;
        }
        data++;
    }

    return encoded;
}

/*
 * Escape control characters in strings with with _xHHHH_.
 */
char *
lxw_escape_control_characters(const char *string)
{
    size_t escape_len = sizeof("_xHHHH_") - 1;
    size_t encoded_len = (strlen(string) * escape_len + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*string) {
        switch (*string) {
            case '\x01':
            case '\x02':
            case '\x03':
            case '\x04':
            case '\x05':
            case '\x06':
            case '\x07':
            case '\x08':
            case '\x0B':
            case '\x0C':
            case '\x0D':
            case '\x0E':
            case '\x0F':
            case '\x10':
            case '\x11':
            case '\x12':
            case '\x13':
            case '\x14':
            case '\x15':
            case '\x16':
            case '\x17':
            case '\x18':
            case '\x19':
            case '\x1A':
            case '\x1B':
            case '\x1C':
            case '\x1D':
            case '\x1E':
            case '\x1F':
                lxw_snprintf(p_encoded, escape_len + 1, "_x%04X_", *string);
                p_encoded += escape_len;
                break;
            default:
                *p_encoded = *string;
                p_encoded++;
                break;
        }
        string++;
    }

    return encoded;
}

/* Write out escaped attributes. */
STATIC void
_fprint_escaped_attributes(FILE * xmlfile,
                           struct xml_attribute_list *attributes)
{
    struct xml_attribute *attribute;

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(xmlfile, " %s=", attribute->key);

            if (!strpbrk(attribute->value, "&<>\"")) {
                fprintf(xmlfile, "\"%s\"", attribute->value);
            }
            else {
                char *encoded = _escape_attributes(attribute);

                if (encoded) {
                    fprintf(xmlfile, "\"%s\"", encoded);

                    free(encoded);
                }
            }
        }
    }
}

/* Write out escaped XML data. */
STATIC void
_fprint_escaped_data(FILE * xmlfile, const char *data)
{
    /* Escape the data section of the XML element. */
    if (!strpbrk(data, "&<>")) {
        fprintf(xmlfile, "%s", data);
    }
    else {
        char *encoded = lxw_escape_data(data);
        if (encoded) {
            fprintf(xmlfile, "%s", encoded);
            free(encoded);
        }
    }
}

/* Create a new string XML attribute. */
struct xml_attribute *
lxw_new_attribute_str(const char *key, const char *value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    LXW_ATTRIBUTE_COPY(attribute->value, value);

    return attribute;
}

/* Create a new integer XML attribute. */
struct xml_attribute *
lxw_new_attribute_int(const char *key, uint32_t value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    lxw_snprintf(attribute->value, LXW_MAX_ATTRIBUTE_LENGTH, "%d", value);

    return attribute;
}

/* Create a new double XML attribute. */
struct xml_attribute *
lxw_new_attribute_dbl(const char *key, double value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    lxw_sprintf_dbl(attribute->value, value);

    return attribute;
}
