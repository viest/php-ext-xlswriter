/*****************************************************************************
 * lxlsx_rich_value_rel - A library for creating Excel XLSX lxlsx_rich_value_rel files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/rich_value_rel.h"
#include "lxlsx/utility.h"

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new lxlsx_rich_value_rel object.
 */
lxlsx_rich_value_rel *
lxlsx_rich_value_rel_new(void)
{
    lxlsx_rich_value_rel *lxlsx_rich_value_rel =
        calloc(1, sizeof(*lxlsx_rich_value_rel));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_rich_value_rel, mem_error);

    return lxlsx_rich_value_rel;

mem_error:
    lxlsx_rich_value_rel_free(lxlsx_rich_value_rel);
    return NULL;
}

/*
 * Free a lxlsx_rich_value_rel object.
 */
void
lxlsx_rich_value_rel_free(lxlsx_rich_value_rel *lxlsx_rich_value_rel)
{
    if (!lxlsx_rich_value_rel)
        return;

    free(lxlsx_rich_value_rel);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
STATIC void
_rich_value_rel_xml_declaration(lxlsx_rich_value_rel *self)
{
    lxlsx_xml_declaration(self->file);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <rel> element.
 */
STATIC void
_rich_value_rel_write_rel(lxlsx_rich_value_rel *self, uint32_t rel_index)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", rel_index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "rel", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <richValueRels> element.
 */
STATIC void
_rich_value_rel_write_rich_value_rels(lxlsx_rich_value_rel *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    char xmlns_r[] =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxlsx_xml_start_tag(self->file, "richValueRels", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_rich_value_rel_assemble_xml_file(lxlsx_rich_value_rel *self)
{
    uint32_t i;

    /* Write the XML declaration. */
    _rich_value_rel_xml_declaration(self);

    /* Write the richValueRels element. */
    _rich_value_rel_write_rich_value_rels(self);

    for (i = 1; i <= self->num_embedded_images; i++) {
        /* Write the rel element. */
        _rich_value_rel_write_rel(self, i);
    }

    lxlsx_xml_end_tag(self->file, "richValueRels");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
