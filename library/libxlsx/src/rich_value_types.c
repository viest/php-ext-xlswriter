/*****************************************************************************
 * lxlsx_rich_value_types - A library for creating Excel XLSX lxlsx_rich_value_types files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/rich_value_types.h"
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
 * Create a new lxlsx_rich_value_types object.
 */
lxlsx_rich_value_types *
lxlsx_rich_value_types_new(void)
{
    lxlsx_rich_value_types *lxlsx_rich_value_types =
        calloc(1, sizeof(*lxlsx_rich_value_types));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_rich_value_types, mem_error);

    return lxlsx_rich_value_types;

mem_error:
    lxlsx_rich_value_types_free(lxlsx_rich_value_types);
    return NULL;
}

/*
 * Free a lxlsx_rich_value_types object.
 */
void
lxlsx_rich_value_types_free(lxlsx_rich_value_types *lxlsx_rich_value_types)
{
    if (!lxlsx_rich_value_types)
        return;

    free(lxlsx_rich_value_types);
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
_rich_value_types_xml_declaration(lxlsx_rich_value_types *self)
{
    lxlsx_xml_declaration(self->file);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <rvTypesInfo> element.
 */
STATIC void
_rich_value_types_write_rv_types_info(lxlsx_rich_value_types *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
    char xmlns_mc[] =
        "http://schemas.openxmlformats.org/markup-compatibility/2006";
    char xmlns_x[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    char mc_ignorable[] = "x";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:mc", xmlns_mc);
    LXLSX_PUSH_ATTRIBUTES_STR("mc:Ignorable", mc_ignorable);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:x", xmlns_x);

    lxlsx_xml_start_tag(self->file, "rvTypesInfo", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <flag> element.
 */
STATIC void
_rich_value_types_write_flag(lxlsx_rich_value_types *self, char *name)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("name", name);
    LXLSX_PUSH_ATTRIBUTES_STR("value", "1");

    lxlsx_xml_empty_tag(self->file, "flag", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <key> element.
 */
STATIC void
_rich_value_types_write_key(lxlsx_rich_value_types *self, char *name)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("name", name);

    lxlsx_xml_start_tag(self->file, "key", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <keyFlags> element.
 */
STATIC void
_rich_value_types_write_key_flags(lxlsx_rich_value_types *self)
{
    int i;
    char *key_flags[10][3] = {
        {"_Self", "ExcludeFromFile", "ExcludeFromCalcComparison"},
        {"_DisplayString", "ExcludeFromCalcComparison", ""},
        {"_Flags", "ExcludeFromCalcComparison", ""},
        {"_Format", "ExcludeFromCalcComparison", ""},
        {"_SubLabel", "ExcludeFromCalcComparison", ""},
        {"_Attribution", "ExcludeFromCalcComparison", ""},
        {"_Icon", "ExcludeFromCalcComparison", ""},
        {"_Display", "ExcludeFromCalcComparison", ""},
        {"_CanonicalPropertyNames", "ExcludeFromCalcComparison", ""},
        {"_ClassificationId", "ExcludeFromCalcComparison", ""},
    };

    lxlsx_xml_start_tag(self->file, "global", NULL);
    lxlsx_xml_start_tag(self->file, "keyFlags", NULL);

    for (i = 0; i < 10; i++) {
        char **flags = key_flags[i];

        _rich_value_types_write_key(self, flags[0]);
        _rich_value_types_write_flag(self, flags[1]);

        if (*flags[2]) {
            _rich_value_types_write_flag(self, flags[2]);

        }

        lxlsx_xml_end_tag(self->file, "key");
    }

    lxlsx_xml_end_tag(self->file, "keyFlags");
    lxlsx_xml_end_tag(self->file, "global");
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_rich_value_types_assemble_xml_file(lxlsx_rich_value_types *self)
{
    /* Write the XML declaration. */
    _rich_value_types_xml_declaration(self);

    /* Write the rvTypesInfo element. */
    _rich_value_types_write_rv_types_info(self);

    /* Write the keyFlags element. */
    _rich_value_types_write_key_flags(self);

    lxlsx_xml_end_tag(self->file, "rvTypesInfo");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
