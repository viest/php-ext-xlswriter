/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * table - A libxlsxwriter library for creating Excel XLSX table files.
 *
 */
#ifndef __LXLSX_TABLE_H__
#define __LXLSX_TABLE_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a table object.
 */
typedef struct lxlsx_table {

    FILE *file;

    struct lxlsx_table_obj *lxlsx_table_obj;

} lxlsx_table;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_table *lxlsx_table_new(void);
void lxlsx_table_free(lxlsx_table *table);
void lxlsx_table_assemble_xml_file(lxlsx_table *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _table_xml_declaration(lxlsx_table *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_TABLE_H__ */
