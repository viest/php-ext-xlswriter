/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * app - A libxlsxwriter library for creating Excel XLSX app files.
 *
 */
#ifndef __LXW_APP_H__
#define __LXW_APP_H__

#include <stdint.h>
#include <string.h>
#include "workbook.h"
#include "common.h"

/* Define the queue.h TAILQ structs for the App structs. */
STAILQ_HEAD(lxw_heading_pairs, lxw_heading_pair);
STAILQ_HEAD(lxw_part_names, lxw_part_name);

typedef struct lxw_heading_pair {

    char *key;
    char *value;

    STAILQ_ENTRY (lxw_heading_pair) list_pointers;

} lxw_heading_pair;

typedef struct lxw_part_name {

    char *name;

    STAILQ_ENTRY (lxw_part_name) list_pointers;

} lxw_part_name;

/* Struct to represent an App object. */
typedef struct lxw_app {

    FILE *file;

    struct lxw_heading_pairs *heading_pairs;
    struct lxw_part_names *part_names;
    lxw_doc_properties *properties;

    uint32_t num_heading_pairs;
    uint32_t num_part_names;

} lxw_app;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_app *lxw_app_new();
void lxw_app_free(lxw_app *app);
void lxw_app_assemble_xml_file(lxw_app *self);
void lxw_app_add_part_name(lxw_app *self, const char *name);
void lxw_app_add_heading_pair(lxw_app *self, const char *key,
                              const char *value);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _app_xml_declaration(lxw_app *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_APP_H__ */
