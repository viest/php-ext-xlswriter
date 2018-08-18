/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * styles - A libxlsxwriter library for creating Excel XLSX styles files.
 *
 */
#ifndef __LXW_STYLES_H__
#define __LXW_STYLES_H__

#include <stdint.h>

#include "format.h"

/*
 * Struct to represent a styles.
 */
typedef struct lxw_styles {

    FILE *file;
    uint32_t font_count;
    uint32_t xf_count;
    uint32_t dxf_count;
    uint32_t num_format_count;
    uint32_t border_count;
    uint32_t fill_count;
    struct lxw_formats *xf_formats;
    struct lxw_formats *dxf_formats;

} lxw_styles;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_styles *lxw_styles_new();
void lxw_styles_free(lxw_styles *styles);
void lxw_styles_assemble_xml_file(lxw_styles *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _styles_xml_declaration(lxw_styles *self);
STATIC void _write_style_sheet(lxw_styles *self);
STATIC void _write_font_size(lxw_styles *self, double font_size);
STATIC void _write_font_color_theme(lxw_styles *self, uint8_t theme);
STATIC void _write_font_name(lxw_styles *self, const char *font_name);
STATIC void _write_font_family(lxw_styles *self, uint8_t font_family);
STATIC void _write_font_scheme(lxw_styles *self, const char *font_scheme);
STATIC void _write_font(lxw_styles *self, lxw_format *format);
STATIC void _write_fonts(lxw_styles *self);
STATIC void _write_default_fill(lxw_styles *self, const char *pattern);
STATIC void _write_fills(lxw_styles *self);
STATIC void _write_border(lxw_styles *self, lxw_format *format);
STATIC void _write_borders(lxw_styles *self);
STATIC void _write_style_xf(lxw_styles *self);
STATIC void _write_cell_style_xfs(lxw_styles *self);
STATIC void _write_xf(lxw_styles *self, lxw_format *format);
STATIC void _write_cell_xfs(lxw_styles *self);
STATIC void _write_cell_style(lxw_styles *self);
STATIC void _write_cell_styles(lxw_styles *self);
STATIC void _write_dxfs(lxw_styles *self);
STATIC void _write_table_styles(lxw_styles *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_STYLES_H__ */
