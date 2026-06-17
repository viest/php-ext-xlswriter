/*****************************************************************************
 * rich_value - A library for creating Excel XLSX rich_value files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/rich_value.h"
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
 * Create a new rich_value object.
 */
lxw_rich_value *
lxw_rich_value_new(void)
{
    lxw_rich_value *rich_value = calloc(1, sizeof(lxw_rich_value));
    GOTO_LABEL_ON_MEM_ERROR(rich_value, mem_error);

    return rich_value;

mem_error:
    lxw_rich_value_free(rich_value);
    return NULL;
}

/*
 * Free a rich_value object.
 */
void
lxw_rich_value_free(lxw_rich_value *rich_value)
{
    if (!rich_value)
        return;

    free(rich_value);
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
_rich_value_xml_declaration(lxw_rich_value *self)
{
    lxw_xml_declaration(self->file);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write the <v> element.
 */
STATIC void
_rich_value_write_v(lxw_rich_value *self, char *value)
{
    lxw_xml_data_element(self->file, "v", value, NULL);

}

/*
 * Write the <rv> element.
 */
STATIC void
_rich_value_write_rv(lxw_rich_value *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("s", "0");

    lxw_xml_start_tag(self->file, "rv", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the metadata for each embedded image.
 */
void
lxw_rich_value_write_images(lxw_rich_value *self)
{

    lxw_workbook *workbook = self->workbook;
    lxw_sheet *sheet;
    lxw_worksheet *worksheet;
    lxw_object_properties *object_props;
    char value[LXW_UINT32_T_LENGTH];
    uint32_t index = 0;
    uint8_t type = 5;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
                       list_pointers) {

            if (object_props->is_duplicate)
                continue;

            if (object_props->decorative)
                type = 6;

            /* Write the rv element. */
            _rich_value_write_rv(self);

            /* Write the v element. */
            lxw_snprintf(value, LXW_UINT32_T_LENGTH, "%u", index);
            _rich_value_write_v(self, value);

            lxw_snprintf(value, LXW_UINT32_T_LENGTH, "%u", type);
            _rich_value_write_v(self, value);

            if (object_props->description && *object_props->description)
                _rich_value_write_v(self, object_props->description);

            lxw_xml_end_tag(self->file, "rv");

            index++;
        }
    }
}

/*
 * Write the <rvData> element.
 */
STATIC void
_rich_value_write_rv_data(lxw_rich_value *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_INT("count", self->workbook->num_embedded_images);

    lxw_xml_start_tag(self->file, "rvData", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxw_rich_value_assemble_xml_file(lxw_rich_value *self)
{
    /* Write the XML declaration. */
    _rich_value_xml_declaration(self);

    /* Write the rvData element. */
    _rich_value_write_rv_data(self);

    lxw_rich_value_write_images(self);

    lxw_xml_end_tag(self->file, "rvData");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
