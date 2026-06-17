/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * rich_value_types - A libxlsxwriter library for creating Excel XLSX rich_value_types files.
 *
 */
#ifndef __LXW_RICH_VALUE_TYPES_H__
#define __LXW_RICH_VALUE_TYPES_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a rich_value_types object.
 */
typedef struct lxw_rich_value_types {

    FILE *file;

} lxw_rich_value_types;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_rich_value_types *lxw_rich_value_types_new(void);
void lxw_rich_value_types_free(lxw_rich_value_types *rich_value_types);
void lxw_rich_value_types_assemble_xml_file(lxw_rich_value_types *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_types_xml_declaration(lxw_rich_value_types *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_RICH_VALUE_TYPES_H__ */
