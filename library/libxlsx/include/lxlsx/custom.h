/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * custom - A libxlsxwriter library for creating Excel custom property files.
 *
 */
#ifndef __LXLSX_CUSTOM_H__
#define __LXLSX_CUSTOM_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a custom property file object.
 */
typedef struct lxlsx_custom {

    FILE *file;

    struct lxlsx_custom_properties *lxlsx_custom_properties;
    uint32_t pid;

} lxlsx_custom;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_custom *lxlsx_custom_new(void);
void lxlsx_custom_free(lxlsx_custom *custom);
void lxlsx_custom_assemble_xml_file(lxlsx_custom *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _custom_xml_declaration(lxlsx_custom *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_CUSTOM_H__ */
