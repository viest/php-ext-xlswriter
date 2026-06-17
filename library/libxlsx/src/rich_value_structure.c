/*****************************************************************************
 * rich_value_structure - A library for creating Excel XLSX rich_value_structure files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/rich_value_structure.h"
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
 * Create a new rich_value_structure object.
 */
lxw_rich_value_structure *
lxw_rich_value_structure_new(void)
{
    lxw_rich_value_structure *rich_value_structure =
        calloc(1, sizeof(lxw_rich_value_structure));
    GOTO_LABEL_ON_MEM_ERROR(rich_value_structure, mem_error);

    return rich_value_structure;

mem_error:
    lxw_rich_value_structure_free(rich_value_structure);
    return NULL;
}

/*
 * Free a rich_value_structure object.
 */
void
lxw_rich_value_structure_free(lxw_rich_value_structure *rich_value_structure)
{
    if (!rich_value_structure)
        return;

    free(rich_value_structure);
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
_rich_value_structure_xml_declaration(lxw_rich_value_structure *self)
{
    lxw_xml_declaration(self->file);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <k> element.
 */
STATIC void
_rich_value_structure_write_k(lxw_rich_value_structure *self, char *name,
                              char *type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("n", name);
    LXW_PUSH_ATTRIBUTES_STR("t", type);

    lxw_xml_empty_tag(self->file, "k", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <s> element.
 */
STATIC void
_rich_value_structure_write_s(lxw_rich_value_structure *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("t", "_localImage");

    lxw_xml_start_tag(self->file, "s", &attributes);

    /* Write the k element. */
    _rich_value_structure_write_k(self, "_rvRel:LocalImageIdentifier", "i");
    _rich_value_structure_write_k(self, "CalcOrigin", "i");

    if (self->has_embedded_image_descriptions)
        _rich_value_structure_write_k(self, "Text", "s");

    lxw_xml_end_tag(self->file, "s");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <rvStructures> element.
 */
STATIC void
_rich_value_structure_write_rv_structures(lxw_rich_value_structure *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("count", "1");

    lxw_xml_start_tag(self->file, "rvStructures", &attributes);

    /* Write the s element. */
    _rich_value_structure_write_s(self);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxw_rich_value_structure_assemble_xml_file(lxw_rich_value_structure *self)
{
    /* Write the XML declaration. */
    _rich_value_structure_xml_declaration(self);

    /* Write the rvStructures element. */
    _rich_value_structure_write_rv_structures(self);

    lxw_xml_end_tag(self->file, "rvStructures");

}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
