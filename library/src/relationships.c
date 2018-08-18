/*****************************************************************************
 * relationships - A library for creating Excel XLSX relationships files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <string.h>
#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/relationships.h"
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
 * Create a new relationships object.
 */
lxw_relationships *
lxw_relationships_new()
{
    lxw_relationships *rels = calloc(1, sizeof(lxw_relationships));
    GOTO_LABEL_ON_MEM_ERROR(rels, mem_error);

    rels->relationships = calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(rels->relationships, mem_error);
    STAILQ_INIT(rels->relationships);

    return rels;

mem_error:
    lxw_free_relationships(rels);
    return NULL;
}

/*
 * Free a relationships object.
 */
void
lxw_free_relationships(lxw_relationships *rels)
{
    lxw_rel_tuple *relationship;

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
STATIC void
_relationships_xml_declaration(lxw_relationships *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <Relationship> element.
 */
STATIC void
_write_relationship(lxw_relationships *self, const char *type,
                    const char *target, const char *target_mode)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH] = { 0 };

    self->rel_id++;
    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", self->rel_id);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("Id", r_id);
    LXW_PUSH_ATTRIBUTES_STR("Type", type);
    LXW_PUSH_ATTRIBUTES_STR("Target", target);

    if (target_mode)
        LXW_PUSH_ATTRIBUTES_STR("TargetMode", target_mode);

    lxw_xml_empty_tag(self->file, "Relationship", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Relationships> element.
 */
STATIC void
_write_relationships(lxw_relationships *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_rel_tuple *rel;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", LXW_SCHEMA_PACKAGE);

    lxw_xml_start_tag(self->file, "Relationships", &attributes);

    STAILQ_FOREACH(rel, self->relationships, list_pointers) {
        _write_relationship(self, rel->type, rel->target, rel->target_mode);
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
lxw_relationships_assemble_xml_file(lxw_relationships *self)
{
    /* Write the XML declaration. */
    _relationships_xml_declaration(self);

    _write_relationships(self);

    /* Close the relationships tag. */
    lxw_xml_end_tag(self->file, "Relationships");
}

/*
 * Add a generic container relationship to XLSX .rels xml files.
 */
STATIC void
_add_relationship(lxw_relationships *self, const char *schema,
                  const char *type, const char *target,
                  const char *target_mode)
{
    lxw_rel_tuple *relationship;

    if (!schema || !type || !target)
        return;

    relationship = calloc(1, sizeof(lxw_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = calloc(1, LXW_MAX_ATTRIBUTE_LENGTH);
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    /* Add the schema to the relationship type. */
    lxw_snprintf(relationship->type, LXW_MAX_ATTRIBUTE_LENGTH, "%s%s",
                 schema, type);

    relationship->target = lxw_strdup(target);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    if (target_mode) {
        relationship->target_mode = lxw_strdup(target_mode);
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
lxw_add_document_relationship(lxw_relationships *self, const char *type,
                              const char *target)
{
    _add_relationship(self, LXW_SCHEMA_DOCUMENT, type, target, NULL);
}

/*
 * Add a package relationship to XLSX .rels xml files.
 */
void
lxw_add_package_relationship(lxw_relationships *self, const char *type,
                             const char *target)
{
    _add_relationship(self, LXW_SCHEMA_PACKAGE, type, target, NULL);
}

/*
 * Add a MS schema package relationship to XLSX .rels xml files.
 */
void
lxw_add_ms_package_relationship(lxw_relationships *self, const char *type,
                                const char *target)
{
    _add_relationship(self, LXW_SCHEMA_MS, type, target, NULL);
}

/*
 * Add a worksheet relationship to sheet .rels xml files.
 */
void
lxw_add_worksheet_relationship(lxw_relationships *self, const char *type,
                               const char *target, const char *target_mode)
{
    _add_relationship(self, LXW_SCHEMA_DOCUMENT, type, target, target_mode);
}
