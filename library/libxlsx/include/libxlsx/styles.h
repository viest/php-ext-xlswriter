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


/* XLSX styles read API. */

#ifndef LXLSX_READER_STYLES_H
#define LXLSX_READER_STYLES_H

#include "libxlsx/common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_styles lxlsx_reader_styles;

typedef enum {
    LXLSX_READER_FMT_CATEGORY_GENERAL = 0,
    LXLSX_READER_FMT_CATEGORY_NUMBER,
    LXLSX_READER_FMT_CATEGORY_PERCENT,
    LXLSX_READER_FMT_CATEGORY_DATE,
    LXLSX_READER_FMT_CATEGORY_TIME,
    LXLSX_READER_FMT_CATEGORY_DATETIME,
    LXLSX_READER_FMT_CATEGORY_CURRENCY,
    LXLSX_READER_FMT_CATEGORY_TEXT,
    LXLSX_READER_FMT_CATEGORY_CUSTOM
} lxlsx_reader_fmt_category;

/* 8-byte AARRGGBB hex color, plus terminator. Empty string = unset/auto. */
#define LXLSX_READER_COLOR_LEN 9

typedef enum {
    LXLSX_READER_UNDERLINE_NONE   = 0,
    LXLSX_READER_UNDERLINE_SINGLE = 1,
    LXLSX_READER_UNDERLINE_DOUBLE = 2,
    LXLSX_READER_UNDERLINE_SINGLE_ACCOUNTING = 3,
    LXLSX_READER_UNDERLINE_DOUBLE_ACCOUNTING = 4
} lxlsx_reader_underline;

typedef struct {
    const char   *name;                 /* font family name (NULL if unset) */
    double        size;                 /* point size; 0 if unset           */
    char          color[LXLSX_READER_COLOR_LEN]; /* AARRGGBB hex; ""=unset/theme     */
    int           bold;                 /* 0/1                              */
    int           italic;
    int           strike;
    lxlsx_reader_underline underline;
} lxlsx_reader_font;

typedef struct {
    const char *pattern_type;           /* "solid"/"none"/"darkGrid"/...    */
    char        fg_color[LXLSX_READER_COLOR_LEN];
    char        bg_color[LXLSX_READER_COLOR_LEN];
} lxlsx_reader_fill;

typedef struct {
    const char *style;                  /* "thin"/"medium"/"none"/...       */
    char        color[LXLSX_READER_COLOR_LEN];
} lxlsx_reader_border_side;

typedef struct {
    lxlsx_reader_border_side left;
    lxlsx_reader_border_side right;
    lxlsx_reader_border_side top;
    lxlsx_reader_border_side bottom;
} lxlsx_reader_border;

typedef struct {
    const char *horizontal;             /* "general"/"left"/"center"/...    */
    const char *vertical;               /* "top"/"center"/"bottom"/...      */
    int         wrap_text;              /* 0/1 */
    int         indent;
    int         rotation;
} lxlsx_reader_alignment;

typedef struct {
    uint16_t         num_fmt_id;
    lxlsx_reader_fmt_category category;
    const char      *format_string;

    uint32_t         font_id;
    uint32_t         fill_id;
    uint32_t         border_id;

    int              has_alignment;
    lxlsx_reader_alignment    alignment;
    int              locked;            /* protection: locked (default 1)  */
    int              hidden;            /* protection: hidden (default 0)  */
} lxlsx_reader_xf;

typedef struct {
    int                 has_font;
    lxlsx_reader_font   font;
    int                 has_fill;
    lxlsx_reader_fill   fill;
    int                 has_border;
    lxlsx_reader_border border;
} lxlsx_reader_dxf;

const lxlsx_reader_xf     *lxlsx_reader_styles_get_xf    (const lxlsx_reader_styles *st, uint32_t style_id);
const lxlsx_reader_font   *lxlsx_reader_styles_get_font  (const lxlsx_reader_styles *st, uint32_t font_id);
const lxlsx_reader_fill   *lxlsx_reader_styles_get_fill  (const lxlsx_reader_styles *st, uint32_t fill_id);
const lxlsx_reader_border *lxlsx_reader_styles_get_border(const lxlsx_reader_styles *st, uint32_t border_id);
size_t                     lxlsx_reader_styles_dxf_count (const lxlsx_reader_styles *st);
const lxlsx_reader_dxf    *lxlsx_reader_styles_get_dxf   (const lxlsx_reader_styles *st, uint32_t dxf_id);

#ifdef __cplusplus
}
#endif

#endif
