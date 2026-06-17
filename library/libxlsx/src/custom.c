/*****************************************************************************
 * custom - A library for creating Excel custom property files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/custom.h"
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
 * Create a new custom object.
 */
lxlsx_custom *
lxlsx_custom_new(void)
{
    lxlsx_custom *custom = calloc(1, sizeof(lxlsx_custom));
    GOTO_LABEL_ON_MEM_ERROR(custom, mem_error);

    return custom;

mem_error:
    lxlsx_custom_free(custom);
    return NULL;
}

/*
 * Free a custom object.
 */
void
lxlsx_custom_free(lxlsx_custom *custom)
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
_custom_xml_declaration(lxlsx_custom *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <vt:lpwstr> element.
 */
STATIC void
_chart_write_vt_lpwstr(lxlsx_custom *self, char *value)
{
    lxlsx_xml_data_element(self->file, "vt:lpwstr", value, NULL);
}

/*
 * Write the <vt:r8> element.
 */
STATIC void
_chart_write_vt_r_8(lxlsx_custom *self, double value)
{
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, value);

    lxlsx_xml_data_element(self->file, "vt:r8", data, NULL);
}

/*
 * Write the <vt:i4> element.
 */
STATIC void
_custom_write_vt_i_4(lxlsx_custom *self, int32_t value)
{
    char data[LXLSX_ATTR_32];

    lxlsx_snprintf(data, LXLSX_ATTR_32, "%d", value);

    lxlsx_xml_data_element(self->file, "vt:i4", data, NULL);
}

/*
 * Write the <vt:bool> element.
 */
STATIC void
_custom_write_vt_bool(lxlsx_custom *self, uint8_t value)
{
    if (value)
        lxlsx_xml_data_element(self->file, "vt:bool", "true", NULL);
    else
        lxlsx_xml_data_element(self->file, "vt:bool", "false", NULL);
}

/*
 * Write the <vt:filetime> element.
 */
STATIC void
_custom_write_vt_filetime(lxlsx_custom *self, lxlsx_datetime *datetime)
{
    char data[LXLSX_DATETIME_LENGTH];

    lxlsx_snprintf(data, LXLSX_DATETIME_LENGTH, "%4d-%02d-%02dT%02d:%02d:%02dZ",
                 datetime->year, datetime->month, datetime->day,
                 datetime->hour, datetime->min, (int) datetime->sec);

    lxlsx_xml_data_element(self->file, "vt:filetime", data, NULL);
}

/*
 * Write the <property> element.
 */
STATIC void
_chart_write_custom_property(lxlsx_custom *self,
                             lxlsx_custom_property *lxlsx_custom_property)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char fmtid[] = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    self->pid++;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("fmtid", fmtid);
    LXLSX_PUSH_ATTRIBUTES_INT("pid", self->pid + 1);
    LXLSX_PUSH_ATTRIBUTES_STR("name", lxlsx_custom_property->name);

    lxlsx_xml_start_tag(self->file, "property", &attributes);

    if (lxlsx_custom_property->type == LXLSX_CUSTOM_STRING) {
        /* Write the vt:lpwstr element. */
        _chart_write_vt_lpwstr(self, lxlsx_custom_property->u.string);
    }
    else if (lxlsx_custom_property->type == LXLSX_CUSTOM_DOUBLE) {
        /* Write the vt:r8 element. */
        _chart_write_vt_r_8(self, lxlsx_custom_property->u.number);
    }
    else if (lxlsx_custom_property->type == LXLSX_CUSTOM_INTEGER) {
        /* Write the vt:i4 element. */
        _custom_write_vt_i_4(self, lxlsx_custom_property->u.integer);
    }
    else if (lxlsx_custom_property->type == LXLSX_CUSTOM_BOOLEAN) {
        /* Write the vt:bool element. */
        _custom_write_vt_bool(self, lxlsx_custom_property->u.boolean);
    }
    else if (lxlsx_custom_property->type == LXLSX_CUSTOM_DATETIME) {
        /* Write the vt:filetime element. */
        _custom_write_vt_filetime(self, &lxlsx_custom_property->u.datetime);
    }

    lxlsx_xml_end_tag(self->file, "property");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <Properties> element.
 */
STATIC void
_write_custom_properties(lxlsx_custom *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = LXLSX_SCHEMA_OFFICEDOC "/custom-properties";
    char xmlns_vt[] = LXLSX_SCHEMA_OFFICEDOC "/docPropsVTypes";
    lxlsx_custom_property *lxlsx_custom_property;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:vt", xmlns_vt);

    lxlsx_xml_start_tag(self->file, "Properties", &attributes);

    STAILQ_FOREACH(lxlsx_custom_property, self->lxlsx_custom_properties, list_pointers) {
        _chart_write_custom_property(self, lxlsx_custom_property);
    }

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
void
lxlsx_custom_assemble_xml_file(lxlsx_custom *self)
{
    /* Write the XML declaration. */
    _custom_xml_declaration(self);

    _write_custom_properties(self);

    lxlsx_xml_end_tag(self->file, "Properties");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
