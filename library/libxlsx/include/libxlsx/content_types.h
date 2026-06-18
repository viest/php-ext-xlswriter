/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * content_types - A libxlsxwriter library for creating Excel XLSX
 *                 content_types files.
 *
 */
#ifndef __LXLSX_CONTENT_TYPES_H__
#define __LXLSX_CONTENT_TYPES_H__

#include <stdint.h>
#include <string.h>

#include "common.h"

#define LXLSX_APP_PACKAGE  "application/vnd.openxmlformats-package."
#define LXLSX_APP_DOCUMENT "application/vnd.openxmlformats-officedocument."
#define LXLSX_APP_MSEXCEL  "application/vnd.ms-excel."

/*
 * Struct to represent a content_types.
 */
typedef struct lxlsx_content_types {

    FILE *file;

    struct lxlsx_tuples *default_types;
    struct lxlsx_tuples *overrides;

} lxlsx_content_types;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_content_types *lxlsx_content_types_new(void);
void lxlsx_content_types_free(lxlsx_content_types *content_types);
void lxlsx_content_types_assemble_xml_file(lxlsx_content_types *content_types);
void lxlsx_ct_add_default(lxlsx_content_types *content_types, const char *key,
                        const char *value);
void lxlsx_ct_add_override(lxlsx_content_types *content_types, const char *key,
                         const char *value);
void lxlsx_ct_add_worksheet_name(lxlsx_content_types *content_types,
                               const char *name);
void lxlsx_ct_add_chartsheet_name(lxlsx_content_types *content_types,
                                const char *name);
void lxlsx_ct_add_chart_name(lxlsx_content_types *content_types,
                           const char *name);
void lxlsx_ct_add_drawing_name(lxlsx_content_types *content_types,
                             const char *name);
void lxlsx_ct_add_table_name(lxlsx_content_types *content_types,
                           const char *name);
void lxlsx_ct_add_comment_name(lxlsx_content_types *content_types,
                             const char *name);
void lxlsx_ct_add_vml_name(lxlsx_content_types *content_types);

void lxlsx_ct_add_shared_strings(lxlsx_content_types *content_types);
void lxlsx_ct_add_calc_chain(lxlsx_content_types *content_types);
void lxlsx_ct_add_custom_properties(lxlsx_content_types *content_types);
void lxlsx_ct_add_metadata(lxlsx_content_types *content_types);
void lxlsx_ct_add_rich_value(lxlsx_content_types *content_types);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _content_types_xml_declaration(lxlsx_content_types *self);
STATIC void _write_default(lxlsx_content_types *self, const char *ext,
                           const char *type);
STATIC void _write_override(lxlsx_content_types *self, const char *part_name,
                            const char *type);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_CONTENT_TYPES_H__ */
