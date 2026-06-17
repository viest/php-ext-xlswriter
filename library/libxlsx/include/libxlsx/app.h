/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * app - A libxlsxwriter library for creating Excel XLSX app files.
 *
 */
#ifndef __LXLSX_APP_H__
#define __LXLSX_APP_H__

#include <stdint.h>
#include <string.h>
#include "workbook.h"
#include "common.h"

/* Define the queue.h TAILQ structs for the App structs. */
STAILQ_HEAD(lxlsx_heading_pairs, lxlsx_heading_pair);
STAILQ_HEAD(lxlsx_part_names, lxlsx_part_name);

typedef struct lxlsx_heading_pair {

    char *key;
    char *value;

    STAILQ_ENTRY (lxlsx_heading_pair) list_pointers;

} lxlsx_heading_pair;

typedef struct lxlsx_part_name {

    char *name;

    STAILQ_ENTRY (lxlsx_part_name) list_pointers;

} lxlsx_part_name;

/* Struct to represent an App object. */
typedef struct lxlsx_app {

    FILE *file;

    struct lxlsx_heading_pairs *heading_pairs;
    struct lxlsx_part_names *part_names;
    lxlsx_doc_properties *properties;

    uint32_t num_heading_pairs;
    uint32_t num_part_names;
    uint8_t doc_security;

} lxlsx_app;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_app *lxlsx_app_new(void);
void lxlsx_app_free(lxlsx_app *app);
void lxlsx_app_assemble_xml_file(lxlsx_app *self);
void lxlsx_app_add_part_name(lxlsx_app *self, const char *name);
void lxlsx_app_add_heading_pair(lxlsx_app *self, const char *key,
                              const char *value);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _app_xml_declaration(lxlsx_app *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_APP_H__ */
