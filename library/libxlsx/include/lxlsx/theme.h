/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * theme - A libxlsxwriter library for creating Excel XLSX theme files.
 *
 */
#ifndef __LXLSX_THEME_H__
#define __LXLSX_THEME_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a theme.
 */
typedef struct lxlsx_theme {

    FILE *file;
} lxlsx_theme;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_theme *lxlsx_theme_new(void);
void lxlsx_theme_free(lxlsx_theme *theme);
void lxlsx_theme_xml_declaration(lxlsx_theme *self);
void lxlsx_theme_assemble_xml_file(lxlsx_theme *self);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_THEME_H__ */
