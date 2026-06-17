/*****************************************************************************
 * rich_value_rel - A library for creating Excel XLSX rich_value_rel files.
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
 * Create a new rich_value_rel object.
 */
lxw_rich_value_rel *
lxw_rich_value_rel_new(void)
{
    lxw_rich_value_rel *rich_value_rel =
        calloc(1, sizeof(lxw_rich_value_rel));
    GOTO_LABEL_ON_MEM_ERROR(rich_value_rel, mem_error);

    return rich_value_rel;

mem_error:
    lxw_rich_value_rel_free(rich_value_rel);
    return NULL;
}

/*
 * Free a rich_value_rel object.
 */
void
lxw_rich_value_rel_free(lxw_rich_value_rel *rich_value_rel)
{
    if (!rich_value_rel)
        return;

    free(rich_value_rel);
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
_rich_value_rel_xml_declaration(lxw_rich_value_rel *self)
{
    lxw_xml_declaration(self->file);
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
_rich_value_rel_write_rel(lxw_rich_value_rel *self, uint32_t rel_index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", rel_index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "rel", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <richValueRels> element.
 */
STATIC void
_rich_value_rel_write_rich_value_rels(lxw_rich_value_rel *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    char xmlns_r[] =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxw_xml_start_tag(self->file, "richValueRels", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxw_rich_value_rel_assemble_xml_file(lxw_rich_value_rel *self)
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

    lxw_xml_end_tag(self->file, "richValueRels");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
