/*****************************************************************************
 * lxlsx_rich_value_structure - A library for creating Excel XLSX lxlsx_rich_value_structure files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/rich_value_structure.h"
#include "libxlsx/utility.h"

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new lxlsx_rich_value_structure object.
 */
lxlsx_rich_value_structure *
lxlsx_rich_value_structure_new(void)
{
    lxlsx_rich_value_structure *lxlsx_rich_value_structure =
        calloc(1, sizeof(*lxlsx_rich_value_structure));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_rich_value_structure, mem_error);

    return lxlsx_rich_value_structure;

mem_error:
    lxlsx_rich_value_structure_free(lxlsx_rich_value_structure);
    return NULL;
}

/*
 * Free a lxlsx_rich_value_structure object.
 */
void
lxlsx_rich_value_structure_free(lxlsx_rich_value_structure *lxlsx_rich_value_structure)
{
    if (!lxlsx_rich_value_structure)
        return;

    free(lxlsx_rich_value_structure);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
LXLSX_DEFINE_XML_DECLARATION(_rich_value_structure_xml_declaration, lxlsx_rich_value_structure)

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <k> element.
 */
STATIC void
_rich_value_structure_write_k(lxlsx_rich_value_structure *self, char *name,
                              char *type)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("n", name);
    LXLSX_PUSH_ATTRIBUTES_STR("t", type);

    lxlsx_xml_empty_tag(self->file, "k", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <s> element.
 */
STATIC void
_rich_value_structure_write_s(lxlsx_rich_value_structure *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("t", "_localImage");

    lxlsx_xml_start_tag(self->file, "s", &attributes);

    /* Write the k element. */
    _rich_value_structure_write_k(self, "_rvRel:LocalImageIdentifier", "i");
    _rich_value_structure_write_k(self, "CalcOrigin", "i");

    if (self->has_embedded_image_descriptions)
        _rich_value_structure_write_k(self, "Text", "s");

    lxlsx_xml_end_tag(self->file, "s");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <rvStructures> element.
 */
STATIC void
_rich_value_structure_write_rv_structures(lxlsx_rich_value_structure *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("count", "1");

    lxlsx_xml_start_tag(self->file, "rvStructures", &attributes);

    /* Write the s element. */
    _rich_value_structure_write_s(self);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_rich_value_structure_assemble_xml_file(lxlsx_rich_value_structure *self)
{
    /* Write the XML declaration. */
    _rich_value_structure_xml_declaration(self);

    /* Write the rvStructures element. */
    _rich_value_structure_write_rv_structures(self);

    lxlsx_xml_end_tag(self->file, "rvStructures");

}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
