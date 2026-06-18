/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * relationships - A libxlsxwriter library for creating Excel XLSX
 *                 relationships files.
 *
 */
#ifndef __LXLSX_RELATIONSHIPS_H__
#define __LXLSX_RELATIONSHIPS_H__

#include <stdint.h>

#include "common.h"

/* Define the queue.h STAILQ structs for the generic data structs. */
STAILQ_HEAD(lxlsx_rel_tuples, lxlsx_rel_tuple);

typedef struct lxlsx_rel_tuple {

    char *type;
    char *target;
    char *target_mode;

    STAILQ_ENTRY (lxlsx_rel_tuple) list_pointers;

} lxlsx_rel_tuple;

/*
 * Struct to represent a relationships.
 */
typedef struct lxlsx_relationships {

    FILE *file;

    uint32_t rel_id;
    struct lxlsx_rel_tuples *relationships;

} lxlsx_relationships;



/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_relationships *lxlsx_relationships_new(void);
void lxlsx_free_relationships(lxlsx_relationships *relationships);
void lxlsx_relationships_assemble_xml_file(lxlsx_relationships *self);

void lxlsx_add_document_relationship(lxlsx_relationships *self, const char *type,
                                   const char *target);
void lxlsx_add_package_relationship(lxlsx_relationships *self, const char *type,
                                  const char *target);
void lxlsx_add_ms_package_relationship(lxlsx_relationships *self,
                                     const char *type, const char *target);
void lxlsx_add_worksheet_relationship(lxlsx_relationships *self, const char *type,
                                    const char *target,
                                    const char *target_mode);
void lxlsx_add_rich_value_relationship(lxlsx_relationships *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _relationships_xml_declaration(lxlsx_relationships *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_RELATIONSHIPS_H__ */
