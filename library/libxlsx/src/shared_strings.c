/*****************************************************************************
 * shared_strings - A library for creating Excel XLSX sst files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/shared_strings.h"
#include "libxlsx/utility.h"
#include <ctype.h>

/*
 * Forward declarations.
 */

STATIC int _element_cmp(struct lxlsx_sst_element *element1,
                        struct lxlsx_sst_element *element2);

#ifndef __clang_analyzer__
LXLSX_RB_GENERATE_ELEMENT(lxlsx_sst_rb_tree, lxlsx_sst_element, lxlsx_sst_tree_pointers,
                        _element_cmp);
#endif

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new SST SharedString object.
 */
lxlsx_sst *
lxlsx_sst_new(void)
{
    /* Create the new shared string table. */
    lxlsx_sst *sst = calloc(1, sizeof(lxlsx_sst));
    RETURN_ON_MEM_ERROR(sst, NULL);

    /* Add the sst RB tree. */
    sst->rb_tree = calloc(1, sizeof(struct lxlsx_sst_rb_tree));
    GOTO_LABEL_ON_MEM_ERROR(sst->rb_tree, mem_error);

    /* Add a list for tracking the insertion order. */
    sst->order_list = calloc(1, sizeof(struct lxlsx_sst_order_list));
    GOTO_LABEL_ON_MEM_ERROR(sst->order_list, mem_error);

    /* Initialize the order list. */
    STAILQ_INIT(sst->order_list);

    /* Initialize the RB tree. */
    RB_INIT(sst->rb_tree);

    return sst;

mem_error:
    lxlsx_sst_free(sst);
    return NULL;
}

/*
 * Free a SST SharedString table object.
 */
void
lxlsx_sst_free(lxlsx_sst *sst)
{
    struct lxlsx_sst_element *lxlsx_sst_element;
    struct lxlsx_sst_element *lxlsx_sst_element_temp;

    if (!sst)
        return;

    /* Free the lxlsx_sst_elements and their data using the ordered linked list. */
    if (sst->order_list) {
        STAILQ_FOREACH_SAFE(lxlsx_sst_element, sst->order_list, lxlsx_sst_order_pointers,
                            lxlsx_sst_element_temp) {

            if (lxlsx_sst_element && lxlsx_sst_element->string)
                free(lxlsx_sst_element->string);
            if (lxlsx_sst_element)
                free(lxlsx_sst_element);
        }
    }

    free(sst->order_list);
    free(sst->rb_tree);
    free(sst);
}

/*
 * Comparator for the element structure
 */
STATIC int
_element_cmp(struct lxlsx_sst_element *element1, struct lxlsx_sst_element *element2)
{
    return strcmp(element1->string, element2->string);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/
/*
 * Write the XML declaration.
 */
LXLSX_DEFINE_XML_DECLARATION(_sst_xml_declaration, lxlsx_sst)

/*
 * Write the <t> element.
 */
STATIC void
_write_t(lxlsx_sst *self, char *string)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Add attribute to preserve leading or trailing whitespace. */
    if (isspace((unsigned char) string[0])
        || isspace((unsigned char) string[strlen(string) - 1]))
        LXLSX_PUSH_ATTRIBUTES_STR("xml:space", "preserve");

    lxlsx_xml_data_element(self->file, "t", string, &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <si> element.
 */
STATIC void
_write_si(lxlsx_sst *self, char *string)
{
    uint8_t escaped_string = LXLSX_FALSE;

    lxlsx_xml_start_tag(self->file, "si", NULL);

    /* Look for and escape control chars in the string. */
    if (lxlsx_has_control_characters(string)) {
        string = lxlsx_escape_control_characters(string);
        escaped_string = LXLSX_TRUE;
    }

    /* Write the t element. */
    _write_t(self, string);

    lxlsx_xml_end_tag(self->file, "si");

    if (escaped_string)
        free(string);
}

/*
 * Write the <si> element for rich strings.
 */
STATIC void
_write_rich_si(lxlsx_sst *self, char *string)
{
    lxlsx_xml_rich_si_element(self->file, string);
}

/*
 * Write the <sst> element.
 */
STATIC void
_write_sst(lxlsx_sst *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->string_count);
    LXLSX_PUSH_ATTRIBUTES_INT("uniqueCount", self->unique_count);

    lxlsx_xml_start_tag(self->file, "sst", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
STATIC void
_write_sst_strings(lxlsx_sst *self)
{
    struct lxlsx_sst_element *lxlsx_sst_element;

    STAILQ_FOREACH(lxlsx_sst_element, self->order_list, lxlsx_sst_order_pointers) {
        /* Write the si element. */
        if (lxlsx_sst_element->is_rich_string)
            _write_rich_si(self, lxlsx_sst_element->string);
        else
            _write_si(self, lxlsx_sst_element->string);

    }
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_sst_assemble_xml_file(lxlsx_sst *self)
{
    /* Write the XML declaration. */
    _sst_xml_declaration(self);

    /* Write the sst element. */
    _write_sst(self);

    /* Write the sst strings. */
    _write_sst_strings(self);

    /* Close the sst tag. */
    lxlsx_xml_end_tag(self->file, "sst");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Add to or find a string in the SST SharedString table and return it's index.
 */
struct lxlsx_sst_element *
lxlsx_get_sst_index(lxlsx_sst *sst, const char *string, uint8_t is_rich_string)
{
    struct lxlsx_sst_element key;
    struct lxlsx_sst_element *element;
    struct lxlsx_sst_element *existing_element;

    key.string = (char *)string;
    existing_element = RB_FIND(lxlsx_sst_rb_tree, sst->rb_tree, &key);
    if (existing_element) {
        sst->string_count++;
        return existing_element;
    }

    /* Create an sst element to potentially add to the table. */
    element = calloc(1, sizeof(struct lxlsx_sst_element));
    if (!element)
        return NULL;

    /* Create potential new element with the string and its index. */
    element->index = sst->unique_count;
    element->string = lxlsx_strdup(string);
    element->is_rich_string = is_rich_string;

    /* Try to insert it and see whether we already have that string. */
    existing_element = RB_INSERT(lxlsx_sst_rb_tree, sst->rb_tree, element);

    /* If existing_element is not NULL, then it already existed. */
    /* Free new created element. */
    if (existing_element) {
        free(element->string);
        free(element);
        sst->string_count++;
        return existing_element;
    }

    /* If it didn't exist, also add it to the insertion order linked list. */
    STAILQ_INSERT_TAIL(sst->order_list, element, lxlsx_sst_order_pointers);

    /* Update SST string counts. */
    sst->string_count++;
    sst->unique_count++;
    return element;
}
