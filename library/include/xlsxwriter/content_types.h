/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * content_types - A libxlsxwriter library for creating Excel XLSX
 *                 content_types files.
 *
 */
#ifndef __LXW_CONTENT_TYPES_H__
#define __LXW_CONTENT_TYPES_H__

#include <stdint.h>
#include <string.h>

#include "common.h"

#define LXW_APP_PACKAGE  "application/vnd.openxmlformats-package."
#define LXW_APP_DOCUMENT "application/vnd.openxmlformats-officedocument."

/*
 * Struct to represent a content_types.
 */
typedef struct lxw_content_types {

    FILE *file;

    struct lxw_tuples *default_types;
    struct lxw_tuples *overrides;

} lxw_content_types;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_content_types *lxw_content_types_new();
void lxw_content_types_free(lxw_content_types *content_types);
void lxw_content_types_assemble_xml_file(lxw_content_types *content_types);
void lxw_ct_add_default(lxw_content_types *content_types, const char *key,
                        const char *value);
void lxw_ct_add_override(lxw_content_types *content_types, const char *key,
                         const char *value);
void lxw_ct_add_worksheet_name(lxw_content_types *content_types,
                               const char *name);
void lxw_ct_add_chart_name(lxw_content_types *content_types,
                           const char *name);
void lxw_ct_add_drawing_name(lxw_content_types *content_types,
                             const char *name);
void lxw_ct_add_shared_strings(lxw_content_types *content_types);
void lxw_ct_add_calc_chain(lxw_content_types *content_types);
void lxw_ct_add_custom_properties(lxw_content_types *content_types);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _content_types_xml_declaration(lxw_content_types *self);
STATIC void _write_default(lxw_content_types *self, const char *ext,
                           const char *type);
STATIC void _write_override(lxw_content_types *self, const char *part_name,
                            const char *type);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_CONTENT_TYPES_H__ */
