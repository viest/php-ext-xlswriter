/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * core - A libxlsxwriter library for creating Excel XLSX core files.
 *
 */
#ifndef __LXW_CORE_H__
#define __LXW_CORE_H__

#include <stdint.h>

#include "workbook.h"
#include "common.h"

/*
 * Struct to represent a core.
 */
typedef struct lxw_core {

    FILE *file;
    lxw_doc_properties *properties;

} lxw_core;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_core *lxw_core_new();
void lxw_core_free(lxw_core *core);
void lxw_core_assemble_xml_file(lxw_core *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _core_xml_declaration(lxw_core *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_CORE_H__ */
