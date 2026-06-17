/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * rich_value_rel - A libxlsxwriter library for creating Excel XLSX rich_value_rel files.
 *
 */
#ifndef __LXW_RICH_VALUE_REL_H__
#define __LXW_RICH_VALUE_REL_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a rich_value_rel object.
 */
typedef struct lxw_rich_value_rel {

    FILE *file;
    uint32_t num_embedded_images;

} lxw_rich_value_rel;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_rich_value_rel *lxw_rich_value_rel_new(void);
void lxw_rich_value_rel_free(lxw_rich_value_rel *rich_value_rel);
void lxw_rich_value_rel_assemble_xml_file(lxw_rich_value_rel *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_rel_xml_declaration(lxw_rich_value_rel *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_RICH_VALUE_REL_H__ */
