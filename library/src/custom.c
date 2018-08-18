/*****************************************************************************
 * custom - A library for creating Excel custom property files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/custom.h"
#include "xlsxwriter/utility.h"

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new custom object.
 */
lxw_custom *
lxw_custom_new()
{
    lxw_custom *custom = calloc(1, sizeof(lxw_custom));
    GOTO_LABEL_ON_MEM_ERROR(custom, mem_error);

    return custom;

mem_error:
    lxw_custom_free(custom);
    return NULL;
}

/*
 * Free a custom object.
 */
void
lxw_custom_free(lxw_custom *custom)
{
    if (!custom)
        return;

    free(custom);
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
_custom_xml_declaration(lxw_custom *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <vt:lpwstr> element.
 */
STATIC void
_chart_write_vt_lpwstr(lxw_custom *self, char *value)
{
    lxw_xml_data_element(self->file, "vt:lpwstr", value, NULL);
}

/*
 * Write the <vt:r8> element.
 */
STATIC void
_chart_write_vt_r_8(lxw_custom *self, double value)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, value);

    lxw_xml_data_element(self->file, "vt:r8", data, NULL);
}

/*
 * Write the <vt:i4> element.
 */
STATIC void
_custom_write_vt_i_4(lxw_custom *self, int32_t value)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%d", value);

    lxw_xml_data_element(self->file, "vt:i4", data, NULL);
}

/*
 * Write the <vt:bool> element.
 */
STATIC void
_custom_write_vt_bool(lxw_custom *self, uint8_t value)
{
    if (value)
        lxw_xml_data_element(self->file, "vt:bool", "true", NULL);
    else
        lxw_xml_data_element(self->file, "vt:bool", "false", NULL);
}

/*
 * Write the <vt:filetime> element.
 */
STATIC void
_custom_write_vt_filetime(lxw_custom *self, lxw_datetime *datetime)
{
    char data[LXW_DATETIME_LENGTH];

    lxw_snprintf(data, LXW_DATETIME_LENGTH, "%4d-%02d-%02dT%02d:%02d:%02dZ",
                 datetime->year, datetime->month, datetime->day,
                 datetime->hour, datetime->min, (int) datetime->sec);

    lxw_xml_data_element(self->file, "vt:filetime", data, NULL);
}

/*
 * Write the <property> element.
 */
STATIC void
_chart_write_custom_property(lxw_custom *self,
                             lxw_custom_property *custom_property)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char fmtid[] = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    self->pid++;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("fmtid", fmtid);
    LXW_PUSH_ATTRIBUTES_INT("pid", self->pid + 1);
    LXW_PUSH_ATTRIBUTES_STR("name", custom_property->name);

    lxw_xml_start_tag(self->file, "property", &attributes);

    if (custom_property->type == LXW_CUSTOM_STRING) {
        /* Write the vt:lpwstr element. */
        _chart_write_vt_lpwstr(self, custom_property->u.string);
    }
    else if (custom_property->type == LXW_CUSTOM_DOUBLE) {
        /* Write the vt:r8 element. */
        _chart_write_vt_r_8(self, custom_property->u.number);
    }
    else if (custom_property->type == LXW_CUSTOM_INTEGER) {
        /* Write the vt:i4 element. */
        _custom_write_vt_i_4(self, custom_property->u.integer);
    }
    else if (custom_property->type == LXW_CUSTOM_BOOLEAN) {
        /* Write the vt:bool element. */
        _custom_write_vt_bool(self, custom_property->u.boolean);
    }
    else if (custom_property->type == LXW_CUSTOM_DATETIME) {
        /* Write the vt:filetime element. */
        _custom_write_vt_filetime(self, &custom_property->u.datetime);
    }

    lxw_xml_end_tag(self->file, "property");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Properties> element.
 */
STATIC void
_write_custom_properties(lxw_custom *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] = LXW_SCHEMA_OFFICEDOC "/custom-properties";
    char xmlns_vt[] = LXW_SCHEMA_OFFICEDOC "/docPropsVTypes";
    lxw_custom_property *custom_property;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:vt", xmlns_vt);

    lxw_xml_start_tag(self->file, "Properties", &attributes);

    STAILQ_FOREACH(custom_property, self->custom_properties, list_pointers) {
        _chart_write_custom_property(self, custom_property);
    }

    LXW_FREE_ATTRIBUTES();
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void
lxw_custom_assemble_xml_file(lxw_custom *self)
{
    /* Write the XML declaration. */
    _custom_xml_declaration(self);

    _write_custom_properties(self);

    lxw_xml_end_tag(self->file, "Properties");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
