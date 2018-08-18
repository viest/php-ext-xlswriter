/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * theme - A libxlsxwriter library for creating Excel XLSX theme files.
 *
 */
#ifndef __LXW_THEME_H__
#define __LXW_THEME_H__

#include <stdint.h>

#include "common.h"

/*
 * Struct to represent a theme.
 */
typedef struct lxw_theme {

    FILE *file;
} lxw_theme;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_theme *lxw_theme_new();
void lxw_theme_free(lxw_theme *theme);
void lxw_theme_xml_declaration(lxw_theme *self);
void lxw_theme_assemble_xml_file(lxw_theme *self);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_THEME_H__ */
