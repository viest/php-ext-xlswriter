/*****************************************************************************
 * app - A library for creating Excel XLSX app files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/app.h"
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
 * Create a new app object.
 */
lxw_app *
lxw_app_new()
{
    lxw_app *app = calloc(1, sizeof(lxw_app));
    GOTO_LABEL_ON_MEM_ERROR(app, mem_error);

    app->heading_pairs = calloc(1, sizeof(struct lxw_heading_pairs));
    GOTO_LABEL_ON_MEM_ERROR(app->heading_pairs, mem_error);
    STAILQ_INIT(app->heading_pairs);

    app->part_names = calloc(1, sizeof(struct lxw_part_names));
    GOTO_LABEL_ON_MEM_ERROR(app->part_names, mem_error);
    STAILQ_INIT(app->part_names);

    return app;

mem_error:
    lxw_app_free(app);
    return NULL;
}

/*
 * Free a app object.
 */
void
lxw_app_free(lxw_app *app)
{
    lxw_heading_pair *heading_pair;
    lxw_part_name *part_name;

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
STATIC void
_app_xml_declaration(lxw_app *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <Properties> element.
 */
STATIC void
_write_properties(lxw_app *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] = LXW_SCHEMA_OFFICEDOC "/extended-properties";
    char xmlns_vt[] = LXW_SCHEMA_OFFICEDOC "/docPropsVTypes";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:vt", xmlns_vt);

    lxw_xml_start_tag(self->file, "Properties", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Application> element.
 */
STATIC void
_write_application(lxw_app *self)
{
    lxw_xml_data_element(self->file, "Application", "Microsoft Excel", NULL);
}

/*
 * Write the <DocSecurity> element.
 */
STATIC void
_write_doc_security(lxw_app *self)
{
    lxw_xml_data_element(self->file, "DocSecurity", "0", NULL);
}

/*
 * Write the <ScaleCrop> element.
 */
STATIC void
_write_scale_crop(lxw_app *self)
{
    lxw_xml_data_element(self->file, "ScaleCrop", "false", NULL);
}

/*
 * Write the <vt:lpstr> element.
 */
STATIC void
_write_vt_lpstr(lxw_app *self, const char *str)
{
    lxw_xml_data_element(self->file, "vt:lpstr", str, NULL);
}

/*
 * Write the <vt:i4> element.
 */
STATIC void
_write_vt_i4(lxw_app *self, const char *value)
{
    lxw_xml_data_element(self->file, "vt:i4", value, NULL);
}

/*
 * Write the <vt:variant> element.
 */
STATIC void
_write_vt_variant(lxw_app *self, const char *key, const char *value)
{
    /* Write the vt:lpstr element. */
    lxw_xml_start_tag(self->file, "vt:variant", NULL);
    _write_vt_lpstr(self, key);
    lxw_xml_end_tag(self->file, "vt:variant");

    /* Write the vt:i4 element. */
    lxw_xml_start_tag(self->file, "vt:variant", NULL);
    _write_vt_i4(self, value);
    lxw_xml_end_tag(self->file, "vt:variant");
}

/*
 * Write the <vt:vector> element for the heading pairs.
 */
STATIC void
_write_vt_vector_heading_pairs(lxw_app *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_heading_pair *heading_pair;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("size", self->num_heading_pairs * 2);
    LXW_PUSH_ATTRIBUTES_STR("baseType", "variant");

    lxw_xml_start_tag(self->file, "vt:vector", &attributes);

    STAILQ_FOREACH(heading_pair, self->heading_pairs, list_pointers) {
        _write_vt_variant(self, heading_pair->key, heading_pair->value);
    }

    lxw_xml_end_tag(self->file, "vt:vector");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <vt:vector> element for the named parts.
 */
STATIC void
_write_vt_vector_lpstr_named_parts(lxw_app *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_part_name *part_name;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("size", self->num_part_names);
    LXW_PUSH_ATTRIBUTES_STR("baseType", "lpstr");

    lxw_xml_start_tag(self->file, "vt:vector", &attributes);

    STAILQ_FOREACH(part_name, self->part_names, list_pointers) {
        _write_vt_lpstr(self, part_name->name);
    }

    lxw_xml_end_tag(self->file, "vt:vector");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <HeadingPairs> element.
 */
STATIC void
_write_heading_pairs(lxw_app *self)
{
    lxw_xml_start_tag(self->file, "HeadingPairs", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_heading_pairs(self);

    lxw_xml_end_tag(self->file, "HeadingPairs");
}

/*
 * Write the <TitlesOfParts> element.
 */
STATIC void
_write_titles_of_parts(lxw_app *self)
{
    lxw_xml_start_tag(self->file, "TitlesOfParts", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_lpstr_named_parts(self);

    lxw_xml_end_tag(self->file, "TitlesOfParts");
}

/*
 * Write the <Manager> element.
 */
STATIC void
_write_manager(lxw_app *self)
{
    lxw_doc_properties *properties = self->properties;

    if (!properties)
        return;

    if (properties->manager)
        lxw_xml_data_element(self->file, "Manager", properties->manager,
                             NULL);
}

/*
 * Write the <Company> element.
 */
STATIC void
_write_company(lxw_app *self)
{
    lxw_doc_properties *properties = self->properties;

    if (properties && properties->company)
        lxw_xml_data_element(self->file, "Company", properties->company,
                             NULL);
    else
        lxw_xml_data_element(self->file, "Company", "", NULL);
}

/*
 * Write the <LinksUpToDate> element.
 */
STATIC void
_write_links_up_to_date(lxw_app *self)
{
    lxw_xml_data_element(self->file, "LinksUpToDate", "false", NULL);
}

/*
 * Write the <SharedDoc> element.
 */
STATIC void
_write_shared_doc(lxw_app *self)
{
    lxw_xml_data_element(self->file, "SharedDoc", "false", NULL);
}

/*
 * Write the <HyperlinkBase> element.
 */
STATIC void
_write_hyperlink_base(lxw_app *self)
{
    lxw_doc_properties *properties = self->properties;

    if (!properties)
        return;

    if (properties->hyperlink_base)
        lxw_xml_data_element(self->file, "HyperlinkBase",
                             properties->hyperlink_base, NULL);
}

/*
 * Write the <HyperlinksChanged> element.
 */
STATIC void
_write_hyperlinks_changed(lxw_app *self)
{
    lxw_xml_data_element(self->file, "HyperlinksChanged", "false", NULL);
}

/*
 * Write the <AppVersion> element.
 */
STATIC void
_write_app_version(lxw_app *self)
{
    lxw_xml_data_element(self->file, "AppVersion", "12.0000", NULL);
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
lxw_app_assemble_xml_file(lxw_app *self)
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

    lxw_xml_end_tag(self->file, "Properties");
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
lxw_app_add_part_name(lxw_app *self, const char *name)
{
    lxw_part_name *part_name;

    if (!name)
        return;

    part_name = calloc(1, sizeof(lxw_part_name));
    GOTO_LABEL_ON_MEM_ERROR(part_name, mem_error);

    part_name->name = lxw_strdup(name);
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
lxw_app_add_heading_pair(lxw_app *self, const char *key, const char *value)
{
    lxw_heading_pair *heading_pair;

    if (!key || !value)
        return;

    heading_pair = calloc(1, sizeof(lxw_heading_pair));
    GOTO_LABEL_ON_MEM_ERROR(heading_pair, mem_error);

    heading_pair->key = lxw_strdup(key);
    GOTO_LABEL_ON_MEM_ERROR(heading_pair->key, mem_error);

    heading_pair->value = lxw_strdup(value);
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
