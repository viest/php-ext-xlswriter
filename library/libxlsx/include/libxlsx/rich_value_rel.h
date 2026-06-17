/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * lxlsx_rich_value_rel - A libxlsxwriter library for creating Excel XLSX lxlsx_rich_value_rel files.
 *
 */
#ifndef __LXLSX_RICH_VALUE_REL_H__
#define __LXLSX_RICH_VALUE_REL_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a lxlsx_rich_value_rel object.
 */
typedef struct lxlsx_rich_value_rel {

    FILE *file;
    uint32_t num_embedded_images;

} lxlsx_rich_value_rel;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_rich_value_rel *lxlsx_rich_value_rel_new(void);
void lxlsx_rich_value_rel_free(lxlsx_rich_value_rel *lxlsx_rich_value_rel);
void lxlsx_rich_value_rel_assemble_xml_file(lxlsx_rich_value_rel *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_rel_xml_declaration(lxlsx_rich_value_rel *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_RICH_VALUE_REL_H__ */
