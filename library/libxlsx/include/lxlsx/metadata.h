/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * metadata - A libxlsxwriter library for creating Excel XLSX metadata files.
 *
 */
#ifndef __LXLSX_METADATA_H__
#define __LXLSX_METADATA_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a metadata object.
 */
typedef struct lxlsx_metadata {

    FILE *file;
    uint8_t has_dynamic_functions;
    uint8_t has_embedded_images;
    uint32_t num_embedded_images;

} lxlsx_metadata;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_metadata *lxlsx_metadata_new(void);
void lxlsx_metadata_free(lxlsx_metadata *metadata);
void lxlsx_metadata_assemble_xml_file(lxlsx_metadata *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _metadata_xml_declaration(lxlsx_metadata *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_METADATA_H__ */
