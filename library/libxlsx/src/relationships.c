/*****************************************************************************
 * relationships - A library for creating Excel XLSX relationships files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include <string.h>
#include "libxlsx/xmlwriter.h"
#include "libxlsx/relationships.h"
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
 * Create a new relationships object.
 */
lxlsx_relationships *
lxlsx_relationships_new(void)
{
    lxlsx_relationships *rels = calloc(1, sizeof(lxlsx_relationships));
    GOTO_LABEL_ON_MEM_ERROR(rels, mem_error);

    rels->relationships = calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(rels->relationships, mem_error);
    STAILQ_INIT(rels->relationships);

    return rels;

mem_error:
    lxlsx_free_relationships(rels);
    return NULL;
}

/*
 * Free a relationships object.
 */
void
lxlsx_free_relationships(lxlsx_relationships *rels)
{
    lxlsx_rel_tuple *relationship;

    if (!rels)
        return;

    if (rels->relationships) {
        while (!STAILQ_EMPTY(rels->relationships)) {
            relationship = STAILQ_FIRST(rels->relationships);
            STAILQ_REMOVE_HEAD(rels->relationships, list_pointers);
            free(relationship->type);
            free(relationship->target);
            free(relationship->target_mode);
            free(relationship);
        }

        free(rels->relationships);
    }

    free(rels);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
LXLSX_DEFINE_XML_DECLARATION(_relationships_xml_declaration, lxlsx_relationships)

/*
 * Write the <Relationship> element.
 */
STATIC void
_write_relationship(lxlsx_relationships *self, const char *type,
                    const char *target, const char *target_mode)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH] = { 0 };

    self->rel_id++;
    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", self->rel_id);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("Id", r_id);
    LXLSX_PUSH_ATTRIBUTES_STR("Type", type);
    LXLSX_PUSH_ATTRIBUTES_STR("Target", target);

    if (target_mode)
        LXLSX_PUSH_ATTRIBUTES_STR("TargetMode", target_mode);

    lxlsx_xml_empty_tag(self->file, "Relationship", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <Relationships> element.
 */
STATIC void
_write_relationships(lxlsx_relationships *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_rel_tuple *rel;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", LXLSX_SCHEMA_PACKAGE);

    lxlsx_xml_start_tag(self->file, "Relationships", &attributes);

    STAILQ_FOREACH(rel, self->relationships, list_pointers) {
        _write_relationship(self, rel->type, rel->target, rel->target_mode);
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
lxlsx_relationships_assemble_xml_file(lxlsx_relationships *self)
{
    /* Write the XML declaration. */
    _relationships_xml_declaration(self);

    _write_relationships(self);

    /* Close the relationships tag. */
    lxlsx_xml_end_tag(self->file, "Relationships");
}

/*
 * Add a generic container relationship to XLSX .rels xml files.
 */
STATIC void
_add_relationship(lxlsx_relationships *self, const char *schema,
                  const char *type, const char *target,
                  const char *target_mode)
{
    lxlsx_rel_tuple *relationship;

    if (!schema || !type || !target)
        return;

    relationship = calloc(1, sizeof(lxlsx_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = calloc(1, LXLSX_MAX_ATTRIBUTE_LENGTH);
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    /* Add the schema to the relationship type. */
    lxlsx_snprintf(relationship->type, LXLSX_MAX_ATTRIBUTE_LENGTH, "%s%s",
                 schema, type);

    relationship->target = lxlsx_strdup(target);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    if (target_mode) {
        relationship->target_mode = lxlsx_strdup(target_mode);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);
    }

    STAILQ_INSERT_TAIL(self->relationships, relationship, list_pointers);

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add a document relationship to XLSX .rels xml files.
 */
void
lxlsx_add_document_relationship(lxlsx_relationships *self, const char *type,
                              const char *target)
{
    _add_relationship(self, LXLSX_SCHEMA_DOCUMENT, type, target, NULL);
}

/*
 * Add a package relationship to XLSX .rels xml files.
 */
void
lxlsx_add_package_relationship(lxlsx_relationships *self, const char *type,
                             const char *target)
{
    _add_relationship(self, LXLSX_SCHEMA_PACKAGE, type, target, NULL);
}

/*
 * Add a MS schema package relationship to XLSX .rels xml files.
 */
void
lxlsx_add_ms_package_relationship(lxlsx_relationships *self, const char *type,
                                const char *target)
{
    _add_relationship(self, LXLSX_SCHEMA_MS, type, target, NULL);
}

/*
 * Add a worksheet relationship to sheet .rels xml files.
 */
void
lxlsx_add_worksheet_relationship(lxlsx_relationships *self, const char *type,
                               const char *target, const char *target_mode)
{
    _add_relationship(self, LXLSX_SCHEMA_DOCUMENT, type, target, target_mode);
}

/*
 * Add a richValue relationship to sheet .rels xml files.
 */
void
lxlsx_add_rich_value_relationship(lxlsx_relationships *self)
{
    _add_relationship(self,
                      "http://schemas.microsoft.com/office/2022/10/relationships/",
                      "richValueRel", "richData/richValueRel.xml", NULL);
    _add_relationship(self,
                      "http://schemas.microsoft.com/office/2017/06/relationships/",
                      "rdRichValue", "richData/rdrichvalue.xml", NULL);
    _add_relationship(self,
                      "http://schemas.microsoft.com/office/2017/06/relationships/",
                      "rdRichValueStructure",
                      "richData/rdrichvaluestructure.xml", NULL);
    _add_relationship(self,
                      "http://schemas.microsoft.com/office/2017/06/relationships/",
                      "rdRichValueTypes", "richData/rdRichValueTypes.xml",
                      NULL);

}
