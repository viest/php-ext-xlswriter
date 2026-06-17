/*****************************************************************************
 * xmlwriter - A base library for libxlsxwriter libraries.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include "lxlsx/xmlwriter.h"

#define LXLSX_AMP  "&amp;"
#define LXLSX_LT   "&lt;"
#define LXLSX_GT   "&gt;"
#define LXLSX_QUOT "&quot;"
#define LXLSX_NL   "&#xA;"

/* Defines. */
#define LXLSX_MAX_ENCODED_ATTRIBUTE_LENGTH (LXLSX_MAX_ATTRIBUTE_LENGTH*6)

/* Forward declarations. */
STATIC char *_escape_attributes(struct lxlsx_xml_attribute *attribute);

char *lxlsx_escape_data(const char *data);

STATIC void _fprint_escaped_attributes(FILE *xmlfile,
                                       struct lxlsx_xml_attribute_list *attributes);

STATIC void _fprint_escaped_data(FILE *xmlfile, const char *data);

/*
 * Write the XML declaration.
 */
void
lxlsx_xml_declaration(FILE *xmlfile)
{
    fprintf(xmlfile, "<?xml version=\"1.0\" "
            "encoding=\"UTF-8\" standalone=\"yes\"?>\n");
}

/*
 * Write an XML start tag with optional attributes.
 */
void
lxlsx_xml_start_tag(FILE *xmlfile,
                  const char *tag, struct lxlsx_xml_attribute_list *attributes)
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
lxlsx_xml_start_tag_unencoded(FILE *xmlfile,
                            const char *tag,
                            struct lxlsx_xml_attribute_list *attributes)
{
    struct lxlsx_xml_attribute *attribute;

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
lxlsx_xml_end_tag(FILE *xmlfile, const char *tag)
{
    fprintf(xmlfile, "</%s>", tag);
}

/*
 * Write an empty XML tag with optional attributes.
 */
void
lxlsx_xml_empty_tag(FILE *xmlfile,
                  const char *tag, struct lxlsx_xml_attribute_list *attributes)
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
lxlsx_xml_empty_tag_unencoded(FILE *xmlfile,
                            const char *tag,
                            struct lxlsx_xml_attribute_list *attributes)
{
    struct lxlsx_xml_attribute *attribute;

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
lxlsx_xml_data_element(FILE *xmlfile,
                     const char *tag,
                     const char *data, struct lxlsx_xml_attribute_list *attributes)
{
    fprintf(xmlfile, "<%s", tag);

    _fprint_escaped_attributes(xmlfile, attributes);

    fprintf(xmlfile, ">");

    _fprint_escaped_data(xmlfile, data);

    fprintf(xmlfile, "</%s>", tag);
}

/*
 * Write an XML <si> element for rich strings, without encoding.
 */
void
lxlsx_xml_rich_si_element(FILE *xmlfile, const char *string)
{
    fprintf(xmlfile, "<si>%s</si>", string);
}

/*
 * Escape XML characters in attributes.
 */
STATIC char *
_escape_attributes(struct lxlsx_xml_attribute *attribute)
{
    char *encoded = (char *) calloc(LXLSX_MAX_ENCODED_ATTRIBUTE_LENGTH, 1);
    char *p_encoded = encoded;
    char *p_attr = attribute->value;

    while (*p_attr) {
        switch (*p_attr) {
            case '&':
                memcpy(p_encoded, LXLSX_AMP, sizeof(LXLSX_AMP) - 1);
                p_encoded += sizeof(LXLSX_AMP) - 1;
                break;
            case '<':
                memcpy(p_encoded, LXLSX_LT, sizeof(LXLSX_LT) - 1);
                p_encoded += sizeof(LXLSX_LT) - 1;
                break;
            case '>':
                memcpy(p_encoded, LXLSX_GT, sizeof(LXLSX_GT) - 1);
                p_encoded += sizeof(LXLSX_GT) - 1;
                break;
            case '"':
                memcpy(p_encoded, LXLSX_QUOT, sizeof(LXLSX_QUOT) - 1);
                p_encoded += sizeof(LXLSX_QUOT) - 1;
                break;
            case '\n':
                memcpy(p_encoded, LXLSX_NL, sizeof(LXLSX_NL) - 1);
                p_encoded += sizeof(LXLSX_NL) - 1;
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
lxlsx_escape_data(const char *data)
{
    size_t encoded_len = (strlen(data) * 5 + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*data) {
        switch (*data) {
            case '&':
                memcpy(p_encoded, LXLSX_AMP, sizeof(LXLSX_AMP) - 1);
                p_encoded += sizeof(LXLSX_AMP) - 1;
                break;
            case '<':
                memcpy(p_encoded, LXLSX_LT, sizeof(LXLSX_LT) - 1);
                p_encoded += sizeof(LXLSX_LT) - 1;
                break;
            case '>':
                memcpy(p_encoded, LXLSX_GT, sizeof(LXLSX_GT) - 1);
                p_encoded += sizeof(LXLSX_GT) - 1;
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
 * Check for control characters in strings.
 */
uint8_t
lxlsx_has_control_characters(const char *string)
{
    while (*string) {
        /* 0xE0 == 0b11100000 masks values > 0x19 == 0b00011111. */
        if (!(*string & 0xE0) && *string != 0x0A && *string != 0x09)
            return LXLSX_TRUE;

        string++;
    }

    return LXLSX_FALSE;
}

/*
 * Escape control characters in strings with _xHHHH_.
 */
char *
lxlsx_escape_control_characters(const char *string)
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
                lxlsx_snprintf(p_encoded, escape_len + 1, "_x%04X_", *string);
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

/*
 * Escape special characters in URL strings with with %XX.
 */
char *
lxlsx_escape_url_characters(const char *string, uint8_t escape_hash)
{

    size_t escape_len = sizeof("%XX") - 1;
    size_t encoded_len = (strlen(string) * escape_len + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*string) {
        switch (*string) {
            case ' ':
            case '"':
            case '<':
            case '>':
            case '[':
            case ']':
            case '`':
            case '^':
            case '{':
            case '}':
                lxlsx_snprintf(p_encoded, escape_len + 1, "%%%2x", *string);
                p_encoded += escape_len;
                break;
            case '#':
                /* This is only escaped for "external:" style links. */
                if (escape_hash) {
                    lxlsx_snprintf(p_encoded, escape_len + 1, "%%%2x", *string);
                    p_encoded += escape_len;
                }
                else {
                    *p_encoded = *string;
                    p_encoded++;
                }
                break;
            case '%':
                /* Only escape % if it isn't already an escape. */
                if (!isxdigit(*(string + 1)) || !isxdigit(*(string + 2))) {
                    lxlsx_snprintf(p_encoded, escape_len + 1, "%%%2x", *string);
                    p_encoded += escape_len;
                }
                else {
                    *p_encoded = *string;
                    p_encoded++;
                }
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
_fprint_escaped_attributes(FILE *xmlfile,
                           struct lxlsx_xml_attribute_list *attributes)
{
    struct lxlsx_xml_attribute *attribute;

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(xmlfile, " %s=", attribute->key);

            if (!strpbrk(attribute->value, "&<>\"\n")) {
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
_fprint_escaped_data(FILE *xmlfile, const char *data)
{
    /* Escape the data section of the XML element. */
    if (!strpbrk(data, "&<>")) {
        fprintf(xmlfile, "%s", data);
    }
    else {
        char *encoded = lxlsx_escape_data(data);
        if (encoded) {
            fprintf(xmlfile, "%s", encoded);
            free(encoded);
        }
    }
}

/* Create a new string XML attribute. */
struct lxlsx_xml_attribute *
lxlsx_new_attribute_str(const char *key, const char *value)
{
    struct lxlsx_xml_attribute *attribute = malloc(sizeof(struct lxlsx_xml_attribute));

    LXLSX_ATTRIBUTE_COPY(attribute->key, key);
    LXLSX_ATTRIBUTE_COPY(attribute->value, value);

    return attribute;
}

/* Create a new integer XML attribute. */
struct lxlsx_xml_attribute *
lxlsx_new_attribute_int(const char *key, int32_t value)
{
    struct lxlsx_xml_attribute *attribute = malloc(sizeof(struct lxlsx_xml_attribute));

    LXLSX_ATTRIBUTE_COPY(attribute->key, key);
    lxlsx_snprintf(attribute->value, LXLSX_MAX_ATTRIBUTE_LENGTH, "%d", value);

    return attribute;
}

/* Create a new double XML attribute. */
struct lxlsx_xml_attribute *
lxlsx_new_attribute_dbl(const char *key, double value)
{
    struct lxlsx_xml_attribute *attribute = malloc(sizeof(struct lxlsx_xml_attribute));

    LXLSX_ATTRIBUTE_COPY(attribute->key, key);
    lxlsx_sprintf_dbl(attribute->value, value);

    return attribute;
}
