/*****************************************************************************
 * app - A library for creating Excel XLSX app files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/app.h"
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
 * Create a new app object.
 */
lxlsx_app *
lxlsx_app_new(void)
{
    lxlsx_app *app = calloc(1, sizeof(lxlsx_app));
    GOTO_LABEL_ON_MEM_ERROR(app, mem_error);

    app->heading_pairs = calloc(1, sizeof(struct lxlsx_heading_pairs));
    GOTO_LABEL_ON_MEM_ERROR(app->heading_pairs, mem_error);
    STAILQ_INIT(app->heading_pairs);

    app->part_names = calloc(1, sizeof(struct lxlsx_part_names));
    GOTO_LABEL_ON_MEM_ERROR(app->part_names, mem_error);
    STAILQ_INIT(app->part_names);

    return app;

mem_error:
    lxlsx_app_free(app);
    return NULL;
}

/*
 * Free a app object.
 */
void
lxlsx_app_free(lxlsx_app *app)
{
    lxlsx_heading_pair *heading_pair;
    lxlsx_part_name *part_name;

    if (!app)
        return;

    /* Free the lists in the App object. */
    if (app->heading_pairs) {
        while (!STAILQ_EMPTY(app->heading_pairs)) {
            heading_pair = STAILQ_FIRST(app->heading_pairs);
            STAILQ_REMOVE_HEAD(app->heading_pairs, list_pointers);
            free(heading_pair->key);
            free(heading_pair->value);
            free(heading_pair);
        }
        free(app->heading_pairs);
    }

    if (app->part_names) {
        while (!STAILQ_EMPTY(app->part_names)) {
            part_name = STAILQ_FIRST(app->part_names);
            STAILQ_REMOVE_HEAD(app->part_names, list_pointers);
            free(part_name->name);
            free(part_name);
        }
        free(app->part_names);
    }

    free(app);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
LXLSX_DEFINE_XML_DECLARATION(_app_xml_declaration, lxlsx_app)

/*
 * Write the <Properties> element.
 */
STATIC void
_write_properties(lxlsx_app *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = LXLSX_SCHEMA_OFFICEDOC "/extended-properties";
    char xmlns_vt[] = LXLSX_SCHEMA_OFFICEDOC "/docPropsVTypes";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:vt", xmlns_vt);

    lxlsx_xml_start_tag(self->file, "Properties", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <Application> element.
 */
STATIC void
_write_application(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "Application", "Microsoft Excel", NULL);
}

/*
 * Write the <DocSecurity> element.
 */
STATIC void
_write_doc_security(lxlsx_app *self)
{
    if (self->doc_security == 2)
        lxlsx_xml_data_element(self->file, "DocSecurity", "2", NULL);
    else
        lxlsx_xml_data_element(self->file, "DocSecurity", "0", NULL);
}

/*
 * Write the <ScaleCrop> element.
 */
STATIC void
_write_scale_crop(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "ScaleCrop", "false", NULL);
}

/*
 * Write the <vt:lpstr> element.
 */
STATIC void
_write_vt_lpstr(lxlsx_app *self, const char *str)
{
    lxlsx_xml_data_element(self->file, "vt:lpstr", str, NULL);
}

/*
 * Write the <vt:i4> element.
 */
STATIC void
_write_vt_i4(lxlsx_app *self, const char *value)
{
    lxlsx_xml_data_element(self->file, "vt:i4", value, NULL);
}

/*
 * Write the <vt:variant> element.
 */
STATIC void
_write_vt_variant(lxlsx_app *self, const char *key, const char *value)
{
    /* Write the vt:lpstr element. */
    lxlsx_xml_start_tag(self->file, "vt:variant", NULL);
    _write_vt_lpstr(self, key);
    lxlsx_xml_end_tag(self->file, "vt:variant");

    /* Write the vt:i4 element. */
    lxlsx_xml_start_tag(self->file, "vt:variant", NULL);
    _write_vt_i4(self, value);
    lxlsx_xml_end_tag(self->file, "vt:variant");
}

/*
 * Write the <vt:vector> element for the heading pairs.
 */
STATIC void
_write_vt_vector_heading_pairs(lxlsx_app *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_heading_pair *heading_pair;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("size", self->num_heading_pairs * 2);
    LXLSX_PUSH_ATTRIBUTES_STR("baseType", "variant");

    lxlsx_xml_start_tag(self->file, "vt:vector", &attributes);

    STAILQ_FOREACH(heading_pair, self->heading_pairs, list_pointers) {
        _write_vt_variant(self, heading_pair->key, heading_pair->value);
    }

    lxlsx_xml_end_tag(self->file, "vt:vector");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <vt:vector> element for the named parts.
 */
STATIC void
_write_vt_vector_lpstr_named_parts(lxlsx_app *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_part_name *part_name;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("size", self->num_part_names);
    LXLSX_PUSH_ATTRIBUTES_STR("baseType", "lpstr");

    lxlsx_xml_start_tag(self->file, "vt:vector", &attributes);

    STAILQ_FOREACH(part_name, self->part_names, list_pointers) {
        _write_vt_lpstr(self, part_name->name);
    }

    lxlsx_xml_end_tag(self->file, "vt:vector");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <HeadingPairs> element.
 */
STATIC void
_write_heading_pairs(lxlsx_app *self)
{
    lxlsx_xml_start_tag(self->file, "HeadingPairs", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_heading_pairs(self);

    lxlsx_xml_end_tag(self->file, "HeadingPairs");
}

/*
 * Write the <TitlesOfParts> element.
 */
STATIC void
_write_titles_of_parts(lxlsx_app *self)
{
    lxlsx_xml_start_tag(self->file, "TitlesOfParts", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_lpstr_named_parts(self);

    lxlsx_xml_end_tag(self->file, "TitlesOfParts");
}

/*
 * Write the <Manager> element.
 */
STATIC void
_write_manager(lxlsx_app *self)
{
    lxlsx_doc_properties *properties = self->properties;

    if (!properties)
        return;

    if (properties->manager)
        lxlsx_xml_data_element(self->file, "Manager", properties->manager,
                             NULL);
}

/*
 * Write the <Company> element.
 */
STATIC void
_write_company(lxlsx_app *self)
{
    lxlsx_doc_properties *properties = self->properties;

    if (properties && properties->company)
        lxlsx_xml_data_element(self->file, "Company", properties->company,
                             NULL);
    else
        lxlsx_xml_data_element(self->file, "Company", "", NULL);
}

/*
 * Write the <LinksUpToDate> element.
 */
STATIC void
_write_links_up_to_date(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "LinksUpToDate", "false", NULL);
}

/*
 * Write the <SharedDoc> element.
 */
STATIC void
_write_shared_doc(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "SharedDoc", "false", NULL);
}

/*
 * Write the <HyperlinkBase> element.
 */
STATIC void
_write_hyperlink_base(lxlsx_app *self)
{
    lxlsx_doc_properties *properties = self->properties;

    if (!properties)
        return;

    if (properties->hyperlink_base)
        lxlsx_xml_data_element(self->file, "HyperlinkBase",
                             properties->hyperlink_base, NULL);
}

/*
 * Write the <HyperlinksChanged> element.
 */
STATIC void
_write_hyperlinks_changed(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "HyperlinksChanged", "false", NULL);
}

/*
 * Write the <AppVersion> element.
 */
STATIC void
_write_app_version(lxlsx_app *self)
{
    lxlsx_xml_data_element(self->file, "AppVersion", "12.0000", NULL);
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
lxlsx_app_assemble_xml_file(lxlsx_app *self)
{

    /* Write the XML declaration. */
    _app_xml_declaration(self);

    _write_properties(self);
    _write_application(self);
    _write_doc_security(self);
    _write_scale_crop(self);
    _write_heading_pairs(self);
    _write_titles_of_parts(self);
    _write_manager(self);
    _write_company(self);
    _write_links_up_to_date(self);
    _write_shared_doc(self);
    _write_hyperlink_base(self);
    _write_hyperlinks_changed(self);
    _write_app_version(self);

    lxlsx_xml_end_tag(self->file, "Properties");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add the name of a workbook Part such as 'Sheet1' or 'Print_Titles'.
 */
void
lxlsx_app_add_part_name(lxlsx_app *self, const char *name)
{
    lxlsx_part_name *part_name;

    if (!name)
        return;

    part_name = calloc(1, sizeof(lxlsx_part_name));
    GOTO_LABEL_ON_MEM_ERROR(part_name, mem_error);

    part_name->name = lxlsx_strdup(name);
    GOTO_LABEL_ON_MEM_ERROR(part_name->name, mem_error);

    STAILQ_INSERT_TAIL(self->part_names, part_name, list_pointers);
    self->num_part_names++;

    return;

mem_error:
    if (part_name) {
        free(part_name->name);
        free(part_name);
    }
}

/*
 * Add the name of a workbook Heading Pair such as 'Worksheets', 'Charts' or
 * 'Named Ranges'.
 */
void
lxlsx_app_add_heading_pair(lxlsx_app *self, const char *key, const char *value)
{
    lxlsx_heading_pair *heading_pair;

    if (!key || !value)
        return;

    heading_pair = calloc(1, sizeof(lxlsx_heading_pair));
    GOTO_LABEL_ON_MEM_ERROR(heading_pair, mem_error);

    heading_pair->key = lxlsx_strdup(key);
    GOTO_LABEL_ON_MEM_ERROR(heading_pair->key, mem_error);

    heading_pair->value = lxlsx_strdup(value);
    GOTO_LABEL_ON_MEM_ERROR(heading_pair->value, mem_error);

    STAILQ_INSERT_TAIL(self->heading_pairs, heading_pair, list_pointers);
    self->num_heading_pairs++;

    return;

mem_error:
    if (heading_pair) {
        free(heading_pair->key);
        free(heading_pair->value);
        free(heading_pair);
    }
}
