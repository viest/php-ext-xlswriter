/*****************************************************************************
 * content_types - A library for creating Excel XLSX content_types files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/content_types.h"
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
 * Create a new content_types object.
 */
lxlsx_content_types *
lxlsx_content_types_new(void)
{
    lxlsx_content_types *content_types = calloc(1, sizeof(lxlsx_content_types));
    GOTO_LABEL_ON_MEM_ERROR(content_types, mem_error);

    content_types->default_types = calloc(1, sizeof(struct lxlsx_tuples));
    GOTO_LABEL_ON_MEM_ERROR(content_types->default_types, mem_error);
    STAILQ_INIT(content_types->default_types);

    content_types->overrides = calloc(1, sizeof(struct lxlsx_tuples));
    GOTO_LABEL_ON_MEM_ERROR(content_types->overrides, mem_error);
    STAILQ_INIT(content_types->overrides);

    lxlsx_ct_add_default(content_types, "rels",
                       LXLSX_APP_PACKAGE "relationships+xml");
    lxlsx_ct_add_default(content_types, "xml", "application/xml");

    lxlsx_ct_add_override(content_types, "/docProps/app.xml",
                        LXLSX_APP_DOCUMENT "extended-properties+xml");
    lxlsx_ct_add_override(content_types, "/docProps/core.xml",
                        LXLSX_APP_PACKAGE "core-properties+xml");
    lxlsx_ct_add_override(content_types, "/xl/styles.xml",
                        LXLSX_APP_DOCUMENT "spreadsheetml.styles+xml");
    lxlsx_ct_add_override(content_types, "/xl/theme/theme1.xml",
                        LXLSX_APP_DOCUMENT "theme+xml");

    return content_types;

mem_error:
    lxlsx_content_types_free(content_types);
    return NULL;
}

/*
 * Free a content_types object.
 */
void
lxlsx_content_types_free(lxlsx_content_types *content_types)
{
    lxlsx_tuple *default_type;
    lxlsx_tuple *override;

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
LXLSX_DEFINE_XML_DECLARATION(_content_types_xml_declaration, lxlsx_content_types)

/*
 * Write the <Types> element.
 */
STATIC void
_write_types(lxlsx_content_types *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", LXLSX_SCHEMA_CONTENT);

    lxlsx_xml_start_tag(self->file, "Types", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <Default> element.
 */
STATIC void
_write_default(lxlsx_content_types *self, const char *ext, const char *type)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("Extension", ext);
    LXLSX_PUSH_ATTRIBUTES_STR("ContentType", type);

    lxlsx_xml_empty_tag(self->file, "Default", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <Override> element.
 */
STATIC void
_write_override(lxlsx_content_types *self, const char *part_name,
                const char *type)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("PartName", part_name);
    LXLSX_PUSH_ATTRIBUTES_STR("ContentType", type);

    lxlsx_xml_empty_tag(self->file, "Override", &attributes);

    LXLSX_FREE_ATTRIBUTES();
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
_write_defaults(lxlsx_content_types *self)
{
    lxlsx_tuple *tuple;

    STAILQ_FOREACH(tuple, self->default_types, list_pointers) {
        _write_default(self, tuple->key, tuple->value);
    }
}

/*
 * Write out all of the <Override> types.
 */
STATIC void
_write_overrides(lxlsx_content_types *self)
{
    lxlsx_tuple *tuple;

    STAILQ_FOREACH(tuple, self->overrides, list_pointers) {
        _write_override(self, tuple->key, tuple->value);
    }
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_content_types_assemble_xml_file(lxlsx_content_types *self)
{
    /* Write the XML declaration. */
    _content_types_xml_declaration(self);

    _write_types(self);
    _write_defaults(self);
    _write_overrides(self);

    /* Close the content_types tag. */
    lxlsx_xml_end_tag(self->file, "Types");
}

static void
_ct_add(struct lxlsx_tuples *items, const char *key, const char *value)
{
    lxlsx_tuple *tuple;

    if (!items || !key || !value)
        return;

    tuple = calloc(1, sizeof(lxlsx_tuple));
    GOTO_LABEL_ON_MEM_ERROR(tuple, mem_error);

    tuple->key = lxlsx_strdup(key);
    GOTO_LABEL_ON_MEM_ERROR(tuple->key, mem_error);

    tuple->value = lxlsx_strdup(value);
    GOTO_LABEL_ON_MEM_ERROR(tuple->value, mem_error);

    STAILQ_INSERT_TAIL(items, tuple, list_pointers);

    return;

mem_error:
    if (tuple) {
        free(tuple->key);
        free(tuple->value);
        free(tuple);
    }
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
lxlsx_ct_add_default(lxlsx_content_types *self, const char *key,
                   const char *value)
{
    if (!self)
        return;
    _ct_add(self->default_types, key, value);
}

/*
 * Add elements to the ContentTypes overrides.
 */
void
lxlsx_ct_add_override(lxlsx_content_types *self, const char *key,
                    const char *value)
{
    if (!self)
        return;
    _ct_add(self->overrides, key, value);
}

/*
 * Add the name of a worksheet to the ContentTypes overrides.
 */
void
lxlsx_ct_add_worksheet_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name,
                        LXLSX_APP_DOCUMENT "spreadsheetml.worksheet+xml");
}

/*
 * Add the name of a chartsheet to the ContentTypes overrides.
 */
void
lxlsx_ct_add_chartsheet_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name,
                        LXLSX_APP_DOCUMENT "spreadsheetml.chartsheet+xml");
}

/*
 * Add the name of a chart to the ContentTypes overrides.
 */
void
lxlsx_ct_add_chart_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name, LXLSX_APP_DOCUMENT "drawingml.chart+xml");
}

/*
 * Add the name of a drawing to the ContentTypes overrides.
 */
void
lxlsx_ct_add_drawing_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name, LXLSX_APP_DOCUMENT "drawing+xml");
}

/*
 * Add the name of a table to the ContentTypes overrides.
 */
void
lxlsx_ct_add_table_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name,
                        LXLSX_APP_DOCUMENT "spreadsheetml.table+xml");
}

/*
 * Add the name of a VML drawing to the ContentTypes overrides.
 */
void
lxlsx_ct_add_vml_name(lxlsx_content_types *self)
{
    lxlsx_ct_add_default(self, "vml", LXLSX_APP_DOCUMENT "vmlDrawing");
}

/*
 * Add the name of a comment to the ContentTypes overrides.
 */
void
lxlsx_ct_add_comment_name(lxlsx_content_types *self, const char *name)
{
    lxlsx_ct_add_override(self, name,
                        LXLSX_APP_DOCUMENT "spreadsheetml.comments+xml");
}

/*
 * Add the sharedStrings link to the ContentTypes overrides.
 */
void
lxlsx_ct_add_shared_strings(lxlsx_content_types *self)
{
    lxlsx_ct_add_override(self, "/xl/sharedStrings.xml",
                        LXLSX_APP_DOCUMENT "spreadsheetml.sharedStrings+xml");
}

/*
 * Add the calcChain link to the ContentTypes overrides.
 */
void
lxlsx_ct_add_calc_chain(lxlsx_content_types *self)
{
    lxlsx_ct_add_override(self, "/xl/calcChain.xml",
                        LXLSX_APP_DOCUMENT "spreadsheetml.calcChain+xml");
}

/*
 * Add the custom properties to the ContentTypes overrides.
 */
void
lxlsx_ct_add_custom_properties(lxlsx_content_types *self)
{
    lxlsx_ct_add_override(self, "/docProps/custom.xml",
                        LXLSX_APP_DOCUMENT "custom-properties+xml");
}

/*
 * Add the metadata file to the ContentTypes overrides.
 */
void
lxlsx_ct_add_metadata(lxlsx_content_types *self)
{
    lxlsx_ct_add_override(self, "/xl/metadata.xml",
                        LXLSX_APP_DOCUMENT "spreadsheetml.sheetMetadata+xml");
}

/*
 * Add the richValue files to the ContentTypes overrides.
 */
void
lxlsx_ct_add_rich_value(lxlsx_content_types *self)
{
    lxlsx_ct_add_override(self, "/xl/richData/rdRichValueTypes.xml",
                        LXLSX_APP_MSEXCEL "rdrichvaluetypes+xml");

    lxlsx_ct_add_override(self, "/xl/richData/rdrichvalue.xml",
                        LXLSX_APP_MSEXCEL "rdrichvalue+xml");

    lxlsx_ct_add_override(self, "/xl/richData/rdrichvaluestructure.xml",
                        LXLSX_APP_MSEXCEL "rdrichvaluestructure+xml");

    lxlsx_ct_add_override(self, "/xl/richData/richValueRel.xml",
                        LXLSX_APP_MSEXCEL "richvaluerel+xml");

}
