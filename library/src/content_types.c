/*****************************************************************************
 * content_types - A library for creating Excel XLSX content_types files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/content_types.h"
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
 * Create a new content_types object.
 */
lxw_content_types *
lxw_content_types_new()
{
    lxw_content_types *content_types = calloc(1, sizeof(lxw_content_types));
    GOTO_LABEL_ON_MEM_ERROR(content_types, mem_error);

    content_types->default_types = calloc(1, sizeof(struct lxw_tuples));
    GOTO_LABEL_ON_MEM_ERROR(content_types->default_types, mem_error);
    STAILQ_INIT(content_types->default_types);

    content_types->overrides = calloc(1, sizeof(struct lxw_tuples));
    GOTO_LABEL_ON_MEM_ERROR(content_types->overrides, mem_error);
    STAILQ_INIT(content_types->overrides);

    lxw_ct_add_default(content_types, "rels",
                       LXW_APP_PACKAGE "relationships+xml");
    lxw_ct_add_default(content_types, "xml", "application/xml");

    lxw_ct_add_override(content_types, "/docProps/app.xml",
                        LXW_APP_DOCUMENT "extended-properties+xml");
    lxw_ct_add_override(content_types, "/docProps/core.xml",
                        LXW_APP_PACKAGE "core-properties+xml");
    lxw_ct_add_override(content_types, "/xl/styles.xml",
                        LXW_APP_DOCUMENT "spreadsheetml.styles+xml");
    lxw_ct_add_override(content_types, "/xl/theme/theme1.xml",
                        LXW_APP_DOCUMENT "theme+xml");
    lxw_ct_add_override(content_types, "/xl/workbook.xml",
                        LXW_APP_DOCUMENT "spreadsheetml.sheet.main+xml");

    return content_types;

mem_error:
    lxw_content_types_free(content_types);
    return NULL;
}

/*
 * Free a content_types object.
 */
void
lxw_content_types_free(lxw_content_types *content_types)
{
    lxw_tuple *default_type;
    lxw_tuple *override;

    if (!content_types)
        return;

    if (content_types->default_types) {
        while (!STAILQ_EMPTY(content_types->default_types)) {
            default_type = STAILQ_FIRST(content_types->default_types);
            STAILQ_REMOVE_HEAD(content_types->default_types, list_pointers);
            free(default_type->key);
            free(default_type->value);
            free(default_type);
        }
        free(content_types->default_types);
    }

    if (content_types->overrides) {
        while (!STAILQ_EMPTY(content_types->overrides)) {
            override = STAILQ_FIRST(content_types->overrides);
            STAILQ_REMOVE_HEAD(content_types->overrides, list_pointers);
            free(override->key);
            free(override->value);
            free(override);
        }
        free(content_types->overrides);
    }

    free(content_types);
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
_content_types_xml_declaration(lxw_content_types *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <Types> element.
 */
STATIC void
_write_types(lxw_content_types *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", LXW_SCHEMA_CONTENT);

    lxw_xml_start_tag(self->file, "Types", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Default> element.
 */
STATIC void
_write_default(lxw_content_types *self, const char *ext, const char *type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("Extension", ext);
    LXW_PUSH_ATTRIBUTES_STR("ContentType", type);

    lxw_xml_empty_tag(self->file, "Default", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Override> element.
 */
STATIC void
_write_override(lxw_content_types *self, const char *part_name,
                const char *type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("PartName", part_name);
    LXW_PUSH_ATTRIBUTES_STR("ContentType", type);

    lxw_xml_empty_tag(self->file, "Override", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write out all of the <Default> types.
 */
STATIC void
_write_defaults(lxw_content_types *self)
{
    lxw_tuple *tuple;

    STAILQ_FOREACH(tuple, self->default_types, list_pointers) {
        _write_default(self, tuple->key, tuple->value);
    }
}

/*
 * Write out all of the <Override> types.
 */
STATIC void
_write_overrides(lxw_content_types *self)
{
    lxw_tuple *tuple;

    STAILQ_FOREACH(tuple, self->overrides, list_pointers) {
        _write_override(self, tuple->key, tuple->value);
    }
}

/*
 * Assemble and write the XML file.
 */
void
lxw_content_types_assemble_xml_file(lxw_content_types *self)
{
    /* Write the XML declaration. */
    _content_types_xml_declaration(self);

    _write_types(self);
    _write_defaults(self);
    _write_overrides(self);

    /* Close the content_types tag. */
    lxw_xml_end_tag(self->file, "Types");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Add elements to the ContentTypes defaults.
 */
void
lxw_ct_add_default(lxw_content_types *self, const char *key,
                   const char *value)
{
    lxw_tuple *tuple;

    if (!key || !value)
        return;

    tuple = calloc(1, sizeof(lxw_tuple));
    GOTO_LABEL_ON_MEM_ERROR(tuple, mem_error);

    tuple->key = lxw_strdup(key);
    GOTO_LABEL_ON_MEM_ERROR(tuple->key, mem_error);

    tuple->value = lxw_strdup(value);
    GOTO_LABEL_ON_MEM_ERROR(tuple->value, mem_error);

    STAILQ_INSERT_TAIL(self->default_types, tuple, list_pointers);

    return;

mem_error:
    if (tuple) {
        free(tuple->key);
        free(tuple->value);
        free(tuple);
    }
}

/*
 * Add elements to the ContentTypes overrides.
 */
void
lxw_ct_add_override(lxw_content_types *self, const char *key,
                    const char *value)
{
    lxw_tuple *tuple;

    if (!key || !value)
        return;

    tuple = calloc(1, sizeof(lxw_tuple));
    GOTO_LABEL_ON_MEM_ERROR(tuple, mem_error);

    tuple->key = lxw_strdup(key);
    GOTO_LABEL_ON_MEM_ERROR(tuple->key, mem_error);

    tuple->value = lxw_strdup(value);
    GOTO_LABEL_ON_MEM_ERROR(tuple->value, mem_error);

    STAILQ_INSERT_TAIL(self->overrides, tuple, list_pointers);

    return;

mem_error:
    if (tuple) {
        free(tuple->key);
        free(tuple->value);
        free(tuple);
    }
}

/*
 * Add the name of a worksheet to the ContentTypes overrides.
 */
void
lxw_ct_add_worksheet_name(lxw_content_types *self, const char *name)
{
    lxw_ct_add_override(self, name,
                        LXW_APP_DOCUMENT "spreadsheetml.worksheet+xml");
}

/*
 * Add the name of a chart to the ContentTypes overrides.
 */
void
lxw_ct_add_chart_name(lxw_content_types *self, const char *name)
{
    lxw_ct_add_override(self, name, LXW_APP_DOCUMENT "drawingml.chart+xml");
}

/*
 * Add the name of a drawing to the ContentTypes overrides.
 */
void
lxw_ct_add_drawing_name(lxw_content_types *self, const char *name)
{
    lxw_ct_add_override(self, name, LXW_APP_DOCUMENT "drawing+xml");
}

/*
 * Add the sharedStrings link to the ContentTypes overrides.
 */
void
lxw_ct_add_shared_strings(lxw_content_types *self)
{
    lxw_ct_add_override(self, "/xl/sharedStrings.xml",
                        LXW_APP_DOCUMENT "spreadsheetml.sharedStrings+xml");
}

/*
 * Add the calcChain link to the ContentTypes overrides.
 */
void
lxw_ct_add_calc_chain(lxw_content_types *self)
{
    lxw_ct_add_override(self, "/xl/calcChain.xml",
                        LXW_APP_DOCUMENT "spreadsheetml.calcChain+xml");
}

/*
 * Add the custom properties to the ContentTypes overrides.
 */
void
lxw_ct_add_custom_properties(lxw_content_types *self)
{
    lxw_ct_add_override(self, "/docProps/custom.xml",
                        LXW_APP_DOCUMENT "custom-properties+xml");
}
