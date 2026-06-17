/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * lxlsx_rich_value_structure - A libxlsxwriter library for creating Excel XLSX lxlsx_rich_value_structure files.
 *
 */
#ifndef __LXLSX_RICH_VALUE_STRUCTURE_H__
#define __LXLSX_RICH_VALUE_STRUCTURE_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a lxlsx_rich_value_structure object.
 */
typedef struct lxlsx_rich_value_structure {

    FILE *file;
    uint8_t has_embedded_image_descriptions;

} lxlsx_rich_value_structure;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_rich_value_structure *lxlsx_rich_value_structure_new(void);
void lxlsx_rich_value_structure_free(lxlsx_rich_value_structure
                                   *lxlsx_rich_value_structure);
void lxlsx_rich_value_structure_assemble_xml_file(lxlsx_rich_value_structure
                                                *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_structure_xml_declaration(lxlsx_rich_value_structure
                                                  *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_RICH_VALUE_STRUCTURE_H__ */
