/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * styles - A libxlsxwriter library for creating Excel XLSX styles files.
 *
 */
#ifndef __LXLSX_STYLES_H__
#define __LXLSX_STYLES_H__

#include <stdint.h>
#include <ctype.h>

#include "format.h"

/*
 * Struct to represent a styles.
 */
typedef struct lxlsx_styles {

    FILE *file;
    uint32_t font_count;
    uint32_t xf_count;
    uint32_t dxf_count;
    uint32_t num_format_count;
    uint32_t border_count;
    uint32_t fill_count;
    struct lxlsx_formats *xf_formats;
    struct lxlsx_formats *dxf_formats;
    uint8_t has_hyperlink;
    uint16_t hyperlink_font_id;
    uint8_t has_comments;

} lxlsx_styles;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_styles *lxlsx_styles_new(void);
void lxlsx_styles_free(lxlsx_styles *styles);
void lxlsx_styles_assemble_xml_file(lxlsx_styles *self);
void lxlsx_styles_write_string_fragment(lxlsx_styles *self, const char *string);
void lxlsx_styles_write_rich_font(lxlsx_styles *styles, lxlsx_format *format);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _styles_xml_declaration(lxlsx_styles *self);
STATIC void _write_style_sheet(lxlsx_styles *self);
STATIC void _write_font_size(lxlsx_styles *self, double font_size);
STATIC void _write_font_color_theme(lxlsx_styles *self, uint8_t theme);
STATIC void _write_font_name(lxlsx_styles *self, const char *font_name,
                             uint8_t is_rich_string);
STATIC void _write_font_family(lxlsx_styles *self, uint8_t font_family);
STATIC void _write_font_scheme(lxlsx_styles *self, const char *font_scheme);
STATIC void _write_font(lxlsx_styles *self, lxlsx_format *format, uint8_t is_dxf,
                        uint8_t is_rich_string);
STATIC void _write_fonts(lxlsx_styles *self);
STATIC void _write_default_fill(lxlsx_styles *self, const char *pattern);
STATIC void _write_fills(lxlsx_styles *self);
STATIC void _write_border(lxlsx_styles *self, lxlsx_format *format,
                          uint8_t is_dxf);
STATIC void _write_borders(lxlsx_styles *self);
STATIC void _write_style_xf(lxlsx_styles *self, uint8_t has_hyperlink,
                            uint16_t font_id);
STATIC void _write_cell_style_xfs(lxlsx_styles *self);
STATIC void _write_xf(lxlsx_styles *self, lxlsx_format *format);
STATIC void _write_cell_xfs(lxlsx_styles *self);
STATIC void _write_cell_style(lxlsx_styles *self, char *name, uint8_t xf_id,
                              uint8_t builtin_id);
STATIC void _write_cell_styles(lxlsx_styles *self);
STATIC void _write_dxfs(lxlsx_styles *self);
STATIC void _write_table_styles(lxlsx_styles *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_STYLES_H__ */
