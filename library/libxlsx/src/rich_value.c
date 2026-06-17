/*****************************************************************************
 * rich_value - A library for creating Excel XLSX rich_value files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/rich_value.h"
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
 * Create a new rich_value object.
 */
lxlsx_rich_value *
lxlsx_rich_value_new(void)
{
    lxlsx_rich_value *rich_value = calloc(1, sizeof(lxlsx_rich_value));
    GOTO_LABEL_ON_MEM_ERROR(rich_value, mem_error);

    return rich_value;

mem_error:
    lxlsx_rich_value_free(rich_value);
    return NULL;
}

/*
 * Free a rich_value object.
 */
void
lxlsx_rich_value_free(lxlsx_rich_value *rich_value)
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
_rich_value_xml_declaration(lxlsx_rich_value *self)
{
    lxlsx_xml_declaration(self->file);
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
_rich_value_write_v(lxlsx_rich_value *self, char *value)
{
    lxlsx_xml_data_element(self->file, "v", value, NULL);

}

/*
 * Write the <rv> element.
 */
STATIC void
_rich_value_write_rv(lxlsx_rich_value *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("s", "0");

    lxlsx_xml_start_tag(self->file, "rv", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the metadata for each embedded image.
 */
void
lxlsx_rich_value_write_images(lxlsx_rich_value *self)
{

    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_object_properties *object_props;
    char value[LXLSX_UINT32_T_LENGTH];
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
            lxlsx_snprintf(value, LXLSX_UINT32_T_LENGTH, "%u", index);
            _rich_value_write_v(self, value);

            lxlsx_snprintf(value, LXLSX_UINT32_T_LENGTH, "%u", type);
            _rich_value_write_v(self, value);

            if (object_props->description && *object_props->description)
                _rich_value_write_v(self, object_props->description);

            lxlsx_xml_end_tag(self->file, "rv");

            index++;
        }
    }
}

/*
 * Write the <rvData> element.
 */
STATIC void
_rich_value_write_rv_data(lxlsx_rich_value *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->workbook->num_embedded_images);

    lxlsx_xml_start_tag(self->file, "rvData", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_rich_value_assemble_xml_file(lxlsx_rich_value *self)
{
    /* Write the XML declaration. */
    _rich_value_xml_declaration(self);

    /* Write the rvData element. */
    _rich_value_write_rv_data(self);

    lxlsx_rich_value_write_images(self);

    lxlsx_xml_end_tag(self->file, "rvData");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
