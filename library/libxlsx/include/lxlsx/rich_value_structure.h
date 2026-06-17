/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * rich_value_structure - A libxlsxwriter library for creating Excel XLSX rich_value_structure files.
 *
 */
#ifndef __LXW_RICH_VALUE_STRUCTURE_H__
#define __LXW_RICH_VALUE_STRUCTURE_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a rich_value_structure object.
 */
typedef struct lxw_rich_value_structure {

    FILE *file;
    uint8_t has_embedded_image_descriptions;

} lxw_rich_value_structure;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_rich_value_structure *lxw_rich_value_structure_new(void);
void lxw_rich_value_structure_free(lxw_rich_value_structure
                                   *rich_value_structure);
void lxw_rich_value_structure_assemble_xml_file(lxw_rich_value_structure
                                                *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_structure_xml_declaration(lxw_rich_value_structure
                                                  *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_RICH_VALUE_STRUCTURE_H__ */
