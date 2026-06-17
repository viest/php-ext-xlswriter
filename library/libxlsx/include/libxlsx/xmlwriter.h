/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * xmlwriter - A libxlsxwriter library for creating Excel XLSX
 *             XML files.
 *
 * The xmlwriter library is used to create the XML sub-components files
 * in the Excel XLSX file format.
 *
 * This library is used in preference to a more generic XML library to allow
 * for customization and optimization for the XLSX file format.
 *
 * The xmlwriter functions are only used internally and do not need to be
 * called directly by the end user.
 *
 */
#ifndef __XMLWRITER_H__
#define __XMLWRITER_H__

#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include "utility.h"

#define LXLSX_MAX_ATTRIBUTE_LENGTH 2080   /* Max URL length. */
#define LXLSX_ATTR_32              32

#define LXLSX_ATTRIBUTE_COPY(dst, src)                    \
    do{                                                 \
        strncpy(dst, src, LXLSX_MAX_ATTRIBUTE_LENGTH -1); \
        dst[LXLSX_MAX_ATTRIBUTE_LENGTH - 1] = '\0';       \
    } while (0)


 /* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/* Attribute used in XML elements. */
struct lxlsx_xml_attribute {
    char key[LXLSX_MAX_ATTRIBUTE_LENGTH];
    char value[LXLSX_MAX_ATTRIBUTE_LENGTH];

    /* Make the struct a queue.h list element. */
    STAILQ_ENTRY (lxlsx_xml_attribute) list_entries;
};

/* Use queue.h macros to define the lxlsx_xml_attribute_list type. */
STAILQ_HEAD(lxlsx_xml_attribute_list, lxlsx_xml_attribute);

/* Create a new attribute struct to add to a lxlsx_xml_attribute_list. */
struct lxlsx_xml_attribute *lxlsx_new_attribute_str(const char *key,
                                            const char *value);
struct lxlsx_xml_attribute *lxlsx_new_attribute_int(const char *key, int32_t value);
struct lxlsx_xml_attribute *lxlsx_new_attribute_dbl(const char *key, double value);

/* Macro to initialize the lxlsx_xml_attribute_list pointers. */
#define LXLSX_INIT_ATTRIBUTES()                                 \
    STAILQ_INIT(&attributes)

/* Macro to add attribute string elements to lxlsx_xml_attribute_list. */
#define LXLSX_PUSH_ATTRIBUTES_STR(key, value)                   \
    do {                                                      \
    attribute = lxlsx_new_attribute_str((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to add attribute int values to lxlsx_xml_attribute_list. */
#define LXLSX_PUSH_ATTRIBUTES_INT(key, value)                   \
    do {                                                      \
    attribute = lxlsx_new_attribute_int((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to add attribute double values to lxlsx_xml_attribute_list. */
#define LXLSX_PUSH_ATTRIBUTES_DBL(key, value)                   \
    do {                                                      \
    attribute = lxlsx_new_attribute_dbl((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to free lxlsx_xml_attribute_list and attribute. */
#define LXLSX_FREE_ATTRIBUTES()                                 \
    do {                                                      \
        while (!STAILQ_EMPTY(&attributes)) {                  \
            attribute = STAILQ_FIRST(&attributes);            \
            STAILQ_REMOVE_HEAD(&attributes, list_entries);    \
            free(attribute);                                  \
        }                                                     \
    } while (0)

/**
 * Create the XML declaration in an XML file.
 *
 * @param xmlfile A FILE pointer to the output XML file.
 */
void lxlsx_xml_declaration(FILE *xmlfile);

/**
 * Write an XML start tag with optional attributes.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
void lxlsx_xml_start_tag(FILE *xmlfile,
                       const char *tag,
                       struct lxlsx_xml_attribute_list *attributes);

/**
 * Write an XML start tag with optional un-encoded attributes.
 * This is a minor optimization for attributes that don't need encoding.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
void lxlsx_xml_start_tag_unencoded(FILE *xmlfile,
                                 const char *tag,
                                 struct lxlsx_xml_attribute_list *attributes);

/**
 * Write an XML end tag.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 */
void lxlsx_xml_end_tag(FILE *xmlfile, const char *tag);

/**
 * Write an XML empty tag with optional attributes.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
void lxlsx_xml_empty_tag(FILE *xmlfile,
                       const char *tag,
                       struct lxlsx_xml_attribute_list *attributes);

/**
 * Write an XML empty tag with optional un-encoded attributes.
 * This is a minor optimization for attributes that don't need encoding.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
void lxlsx_xml_empty_tag_unencoded(FILE *xmlfile,
                                 const char *tag,
                                 struct lxlsx_xml_attribute_list *attributes);

/**
 * Write an XML element containing data and optional attributes.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param data       The data section of the XML element.
 * @param attributes An optional list of attributes to add to the tag.
 */
void lxlsx_xml_data_element(FILE *xmlfile,
                          const char *tag,
                          const char *data,
                          struct lxlsx_xml_attribute_list *attributes);

void lxlsx_xml_rich_si_element(FILE *xmlfile, const char *string);

uint8_t lxlsx_has_control_characters(const char *string);
char *lxlsx_escape_control_characters(const char *string);
char *lxlsx_escape_url_characters(const char *string, uint8_t escape_hash);

char *lxlsx_escape_data(const char *data);

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __XMLWRITER_H__ */
