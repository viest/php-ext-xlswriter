/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * custom - A libxlsxwriter library for creating Excel custom property files.
 *
 */
#ifndef __LXW_CUSTOM_H__
#define __LXW_CUSTOM_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a custom property file object.
 */
typedef struct lxw_custom {

    FILE *file;

    struct lxw_custom_properties *custom_properties;
    uint32_t pid;

} lxw_custom;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_custom *lxw_custom_new();
void lxw_custom_free(lxw_custom *custom);
void lxw_custom_assemble_xml_file(lxw_custom *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _custom_xml_declaration(lxw_custom *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_CUSTOM_H__ */
