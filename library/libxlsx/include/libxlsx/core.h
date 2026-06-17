/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * core - A libxlsxwriter library for creating Excel XLSX core files.
 *
 */
#ifndef __LXLSX_CORE_H__
#define __LXLSX_CORE_H__

#include <stdint.h>

#include "workbook.h"
#include "common.h"

/*
 * Struct to represent a core.
 */
typedef struct lxlsx_core {

    FILE *file;
    lxlsx_doc_properties *properties;

} lxlsx_core;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_core *lxlsx_core_new(void);
void lxlsx_core_free(lxlsx_core *core);
void lxlsx_core_assemble_xml_file(lxlsx_core *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _core_xml_declaration(lxlsx_core *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_CORE_H__ */
