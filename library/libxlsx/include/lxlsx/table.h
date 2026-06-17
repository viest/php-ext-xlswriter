/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * table - A libxlsxwriter library for creating Excel XLSX table files.
 *
 */
#ifndef __LXW_TABLE_H__
#define __LXW_TABLE_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a table object.
 */
typedef struct lxw_table {

    FILE *file;

    struct lxw_table_obj *table_obj;

} lxw_table;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_table *lxw_table_new(void);
void lxw_table_free(lxw_table *table);
void lxw_table_assemble_xml_file(lxw_table *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _table_xml_declaration(lxw_table *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_TABLE_H__ */
