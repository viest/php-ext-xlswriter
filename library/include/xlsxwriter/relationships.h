/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * relationships - A libxlsxwriter library for creating Excel XLSX
 *                 relationships files.
 *
 */
#ifndef __LXW_RELATIONSHIPS_H__
#define __LXW_RELATIONSHIPS_H__

#include <stdint.h>

#include "common.h"

/* Define the queue.h STAILQ structs for the generic data structs. */
STAILQ_HEAD(lxw_rel_tuples, lxw_rel_tuple);

typedef struct lxw_rel_tuple {

    char *type;
    char *target;
    char *target_mode;

    STAILQ_ENTRY (lxw_rel_tuple) list_pointers;

} lxw_rel_tuple;

/*
 * Struct to represent a relationships.
 */
typedef struct lxw_relationships {

    FILE *file;

    uint32_t rel_id;
    struct lxw_rel_tuples *relationships;

} lxw_relationships;



/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_relationships *lxw_relationships_new();
void lxw_free_relationships(lxw_relationships *relationships);
void lxw_relationships_assemble_xml_file(lxw_relationships *self);

void lxw_add_document_relationship(lxw_relationships *self, const char *type,
                                   const char *target);
void lxw_add_package_relationship(lxw_relationships *self, const char *type,
                                  const char *target);
void lxw_add_ms_package_relationship(lxw_relationships *self,
                                     const char *type, const char *target);
void lxw_add_worksheet_relationship(lxw_relationships *self, const char *type,
                                    const char *target,
                                    const char *target_mode);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _relationships_xml_declaration(lxw_relationships *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_RELATIONSHIPS_H__ */
