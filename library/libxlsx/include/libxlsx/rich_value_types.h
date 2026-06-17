/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * lxlsx_rich_value_types - A libxlsxwriter library for creating Excel XLSX lxlsx_rich_value_types files.
 *
 */
#ifndef __LXLSX_RICH_VALUE_TYPES_H__
#define __LXLSX_RICH_VALUE_TYPES_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a lxlsx_rich_value_types object.
 */
typedef struct lxlsx_rich_value_types {

    FILE *file;

} lxlsx_rich_value_types;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_rich_value_types *lxlsx_rich_value_types_new(void);
void lxlsx_rich_value_types_free(lxlsx_rich_value_types *lxlsx_rich_value_types);
void lxlsx_rich_value_types_assemble_xml_file(lxlsx_rich_value_types *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_types_xml_declaration(lxlsx_rich_value_types *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_RICH_VALUE_TYPES_H__ */
