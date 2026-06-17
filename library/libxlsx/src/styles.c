/*****************************************************************************
 * styles - A library for creating Excel XLSX styles files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/styles.h"
#include "libxlsx/utility.h"

/*
 * Forward declarations.
 */
STATIC void _write_font(lxlsx_styles *self, lxlsx_format *format, uint8_t is_dxf,
                        uint8_t is_rich_string);

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new styles object.
 */
lxlsx_styles *
lxlsx_styles_new(void)
{
    lxlsx_styles *styles = calloc(1, sizeof(lxlsx_styles));
    GOTO_LABEL_ON_MEM_ERROR(styles, mem_error);

    styles->xf_formats = calloc(1, sizeof(struct lxlsx_formats));
    GOTO_LABEL_ON_MEM_ERROR(styles->xf_formats, mem_error);
    STAILQ_INIT(styles->xf_formats);

    styles->dxf_formats = calloc(1, sizeof(struct lxlsx_formats));
    GOTO_LABEL_ON_MEM_ERROR(styles->dxf_formats, mem_error);
    STAILQ_INIT(styles->dxf_formats);

    return styles;

mem_error:
    lxlsx_styles_free(styles);
    return NULL;
}

/*
 * Free a styles object.
 */
void
lxlsx_styles_free(lxlsx_styles *styles)
{
    lxlsx_format *format;

    if (!styles)
        return;

    /* Free the xf formats in the styles. */
    if (styles->xf_formats) {
        while (!STAILQ_EMPTY(styles->xf_formats)) {
            format = STAILQ_FIRST(styles->xf_formats);
            STAILQ_REMOVE_HEAD(styles->xf_formats, list_pointers);
            free(format);
        }
        free(styles->xf_formats);
    }

    /* Free the dxf formats in the styles. */
    if (styles->dxf_formats) {
        while (!STAILQ_EMPTY(styles->dxf_formats)) {
            format = STAILQ_FIRST(styles->dxf_formats);
            STAILQ_REMOVE_HEAD(styles->dxf_formats, list_pointers);
            free(format);
        }
        free(styles->dxf_formats);
    }

    free(styles);
}

/*
 * Write the <t> element for rich strings.
 */
void
lxlsx_styles_write_string_fragment(lxlsx_styles *self, const char *string)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Add attribute to preserve leading or trailing whitespace. */
    if (isspace((unsigned char) string[0])
        || isspace((unsigned char) string[strlen(string) - 1]))
        LXLSX_PUSH_ATTRIBUTES_STR("xml:space", "preserve");

    lxlsx_xml_data_element(self->file, "t", string, &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

void
lxlsx_styles_write_rich_font(lxlsx_styles *self, lxlsx_format *format)
{

    _write_font(self, format, LXLSX_FALSE, LXLSX_TRUE);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
STATIC void
_styles_xml_declaration(lxlsx_styles *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <styleSheet> element.
 */
STATIC void
_write_style_sheet(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns",
                            "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    lxlsx_xml_start_tag(self->file, "styleSheet", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <numFmt> element.
 */
STATIC void
_write_num_fmt(lxlsx_styles *self, uint16_t num_fmt_id, char *lxlsx_format_code)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char *lxlsx_format_codes[] = {
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "($#,##0_);($#,##0)",
        "($#,##0_);[Red]($#,##0)",
        "($#,##0.00_);($#,##0.00)",
        "($#,##0.00_);[Red]($#,##0.00)",
        "0%",
        "0.00%",
        "0.00E+00",
        "# ?/?",
        "# ?" "?/?" "?",        /* Split string to avoid unintentional trigraph. */
        "m/d/yy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "h:mm AM/PM",
        "h:mm:ss AM/PM",
        "h:mm",
        "h:mm:ss",
        "m/d/yy h:mm",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "General",
        "(#,##0_);(#,##0)",
        "(#,##0_);[Red](#,##0)",
        "(#,##0.00_);(#,##0.00)",
        "(#,##0.00_);[Red](#,##0.00)",
        "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
        "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "mm:ss",
        "[h]:mm:ss",
        "mm:ss.0",
        "##0.0E+0",
        "@"
    };

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("numFmtId", num_fmt_id);

    if (num_fmt_id < 50)
        LXLSX_PUSH_ATTRIBUTES_STR("formatCode", lxlsx_format_codes[num_fmt_id]);
    else if (num_fmt_id < 164)
        LXLSX_PUSH_ATTRIBUTES_STR("formatCode", "General");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("formatCode", lxlsx_format_code);

    lxlsx_xml_empty_tag(self->file, "numFmt", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <numFmts> element.
 */
STATIC void
_write_num_fmts(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;
    uint16_t last_format_index = 0;

    if (!self->num_format_count)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->num_format_count);

    lxlsx_xml_start_tag(self->file, "numFmts", &attributes);

    /* Write the numFmts elements. */
    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {

        /* Ignore built-in number formats, i.e., < 164. */
        if (format->num_format_index < 164)
            continue;

        /* Ignore duplicates which have an already used index. */
        if (format->num_format_index <= last_format_index)
            continue;

        _write_num_fmt(self, format->num_format_index, format->num_format);

        last_format_index = format->num_format_index;
    }

    lxlsx_xml_end_tag(self->file, "numFmts");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sz> element.
 */
STATIC void
_write_font_size(lxlsx_styles *self, double font_size)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_DBL("val", font_size);

    lxlsx_xml_empty_tag(self->file, "sz", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <color> element for themes.
 */
STATIC void
_write_font_color_theme(lxlsx_styles *self, uint8_t theme)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("theme", theme);

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <color> element for RGB colors.
 */
STATIC void
_write_font_color_rgb(lxlsx_styles *self, int32_t rgb)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb_str[LXLSX_ATTR_32];

    lxlsx_snprintf(rgb_str, LXLSX_ATTR_32, "FF%06X", rgb & LXLSX_COLOR_MASK);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <color> element for indexed colors.
 */
STATIC void
_write_font_color_indexed(lxlsx_styles *self, uint8_t index)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("indexed", index);

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <name> element.
 */
STATIC void
_write_font_name(lxlsx_styles *self, const char *font_name,
                 uint8_t is_rich_string)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (*font_name)
        LXLSX_PUSH_ATTRIBUTES_STR("val", font_name);
    else
        LXLSX_PUSH_ATTRIBUTES_STR("val", LXLSX_DEFAULT_FONT_NAME);

    if (is_rich_string)
        lxlsx_xml_empty_tag(self->file, "rFont", &attributes);
    else
        lxlsx_xml_empty_tag(self->file, "name", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <family> element.
 */
STATIC void
_write_font_family(lxlsx_styles *self, uint8_t font_family)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("val", font_family);

    lxlsx_xml_empty_tag(self->file, "family", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <charset> element.
 */
STATIC void
_write_font_charset(lxlsx_styles *self, uint8_t font_charset)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("val", font_charset);

    lxlsx_xml_empty_tag(self->file, "charset", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <scheme> element.
 */
STATIC void
_write_font_scheme(lxlsx_styles *self, const char *font_scheme)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (*font_scheme)
        LXLSX_PUSH_ATTRIBUTES_STR("val", font_scheme);
    else
        LXLSX_PUSH_ATTRIBUTES_STR("val", "minor");

    lxlsx_xml_empty_tag(self->file, "scheme", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the underline font element.
 */
STATIC void
_write_font_underline(lxlsx_styles *self, uint8_t underline)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Handle the underline variants. */
    if (underline == LXLSX_UNDERLINE_DOUBLE)
        LXLSX_PUSH_ATTRIBUTES_STR("val", "double");
    else if (underline == LXLSX_UNDERLINE_SINGLE_ACCOUNTING)
        LXLSX_PUSH_ATTRIBUTES_STR("val", "singleAccounting");
    else if (underline == LXLSX_UNDERLINE_DOUBLE_ACCOUNTING)
        LXLSX_PUSH_ATTRIBUTES_STR("val", "doubleAccounting");
    /* Default to single underline. */

    lxlsx_xml_empty_tag(self->file, "u", &attributes);

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the font <condense> element.
 */
STATIC void
_write_font_condense(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("val", "0");

    lxlsx_xml_empty_tag(self->file, "condense", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the font <extend> element.
 */
STATIC void
_write_font_extend(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("val", "0");

    lxlsx_xml_empty_tag(self->file, "extend", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <vertAlign> font sub-element.
 */
STATIC void
_write_font_vert_align(lxlsx_styles *self, const char *align)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("val", align);

    lxlsx_xml_empty_tag(self->file, "vertAlign", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <font> element.
 */
STATIC void
_write_font(lxlsx_styles *self, lxlsx_format *format, uint8_t is_dxf,
            uint8_t is_rich_string)
{
    if (is_rich_string)
        lxlsx_xml_start_tag(self->file, "rPr", NULL);
    else
        lxlsx_xml_start_tag(self->file, "font", NULL);

    if (format->font_condense)
        _write_font_condense(self);

    if (format->font_extend)
        _write_font_extend(self);

    if (format->bold)
        lxlsx_xml_empty_tag(self->file, "b", NULL);

    if (format->italic)
        lxlsx_xml_empty_tag(self->file, "i", NULL);

    if (format->font_strikeout)
        lxlsx_xml_empty_tag(self->file, "strike", NULL);

    if (format->font_outline)
        lxlsx_xml_empty_tag(self->file, "outline", NULL);

    if (format->font_shadow)
        lxlsx_xml_empty_tag(self->file, "shadow", NULL);

    if (format->underline)
        _write_font_underline(self, format->underline);

    if (format->font_script == LXLSX_FONT_SUPERSCRIPT)
        _write_font_vert_align(self, "superscript");

    if (format->font_script == LXLSX_FONT_SUBSCRIPT)
        _write_font_vert_align(self, "subscript");

    if (!is_dxf && format->font_size > 0.0)
        _write_font_size(self, format->font_size);

    if (format->theme)
        _write_font_color_theme(self, format->theme);
    else if (format->color_indexed)
        _write_font_color_indexed(self, format->color_indexed);
    else if (format->font_color != LXLSX_COLOR_UNSET)
        _write_font_color_rgb(self, format->font_color);
    else if (!is_dxf)
        _write_font_color_theme(self, LXLSX_DEFAULT_FONT_THEME);

    if (!is_dxf) {
        _write_font_name(self, format->font_name, is_rich_string);
        _write_font_family(self, format->font_family);

        if (format->font_charset)
            _write_font_charset(self, format->font_charset);

        /* Only write the scheme element for the default font type if it
         * isn't a hyperlink. */
        if ((!*format->font_name
             || strcmp(LXLSX_DEFAULT_FONT_NAME, format->font_name) == 0)
            && !format->hyperlink) {
            _write_font_scheme(self, format->font_scheme);
        }
    }

    if (format->hyperlink) {
        self->has_hyperlink = LXLSX_TRUE;

        if (self->hyperlink_font_id == 0)
            self->hyperlink_font_id = format->font_index;
    }

    if (is_rich_string)
        lxlsx_xml_end_tag(self->file, "rPr");
    else
        lxlsx_xml_end_tag(self->file, "font");
}

/*
 * Write the <font> element for comments.
 */
STATIC void
_write_comment_font(lxlsx_styles *self)
{
    lxlsx_xml_start_tag(self->file, "font", NULL);

    _write_font_size(self, 8);
    _write_font_color_indexed(self, 81);
    _write_font_name(self, "Tahoma", LXLSX_FALSE);
    _write_font_family(self, 2);

    lxlsx_xml_end_tag(self->file, "font");
}

/*
 * Write the <fonts> element.
 */
STATIC void
_write_fonts(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;
    uint32_t count;

    LXLSX_INIT_ATTRIBUTES();

    count = self->font_count;
    if (self->has_comments)
        count++;

    LXLSX_PUSH_ATTRIBUTES_INT("count", count);

    lxlsx_xml_start_tag(self->file, "fonts", &attributes);

    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
        if (format->has_font)
            _write_font(self, format, LXLSX_FALSE, LXLSX_FALSE);
    }

    if (self->has_comments)
        _write_comment_font(self);

    lxlsx_xml_end_tag(self->file, "fonts");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the default <fill> element.
 */
STATIC void
_write_default_fill(lxlsx_styles *self, const char *pattern)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("patternType", pattern);

    lxlsx_xml_start_tag(self->file, "fill", NULL);
    lxlsx_xml_empty_tag(self->file, "patternFill", &attributes);
    lxlsx_xml_end_tag(self->file, "fill");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <fgColor> element.
 */
STATIC void
_write_fg_color(lxlsx_styles *self, lxlsx_color_t color)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb_str[LXLSX_ATTR_32];

    LXLSX_INIT_ATTRIBUTES();

    lxlsx_snprintf(rgb_str, LXLSX_ATTR_32, "FF%06X", color & LXLSX_COLOR_MASK);
    LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

    lxlsx_xml_empty_tag(self->file, "fgColor", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <bgColor> element.
 */
STATIC void
_write_bg_color(lxlsx_styles *self, lxlsx_color_t color, uint8_t pattern)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb_str[LXLSX_ATTR_32];

    LXLSX_INIT_ATTRIBUTES();

    if (color == LXLSX_COLOR_UNSET) {
        if (pattern <= LXLSX_PATTERN_SOLID) {
            LXLSX_PUSH_ATTRIBUTES_STR("indexed", "64");
            lxlsx_xml_empty_tag(self->file, "bgColor", &attributes);
        }
    }
    else {
        lxlsx_snprintf(rgb_str, LXLSX_ATTR_32, "FF%06X", color & LXLSX_COLOR_MASK);
        LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb_str);
        lxlsx_xml_empty_tag(self->file, "bgColor", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <fill> element.
 */
STATIC void
_write_fill(lxlsx_styles *self, lxlsx_format *format, uint8_t is_dxf)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    uint8_t pattern = format->pattern;
    lxlsx_color_t bg_color = format->bg_color;
    lxlsx_color_t fg_color = format->fg_color;

    char *patterns[] = {
        "none",
        "solid",
        "mediumGray",
        "darkGray",
        "lightGray",
        "darkHorizontal",
        "darkVertical",
        "darkDown",
        "darkUp",
        "darkGrid",
        "darkTrellis",
        "lightHorizontal",
        "lightVertical",
        "lightDown",
        "lightUp",
        "lightGrid",
        "lightTrellis",
        "gray125",
        "gray0625",
    };

    if (is_dxf) {
        bg_color = format->dxf_bg_color;
        fg_color = format->dxf_fg_color;
    }

    LXLSX_INIT_ATTRIBUTES();

    /* Special handling for pattern only case. */
    if (!bg_color && !fg_color && pattern) {
        _write_default_fill(self, patterns[pattern]);
        LXLSX_FREE_ATTRIBUTES();
        return;
    }

    lxlsx_xml_start_tag(self->file, "fill", NULL);

    /* None/Solid patterns are handled differently for dxf formats. */
    if (pattern && !(is_dxf && pattern <= LXLSX_PATTERN_SOLID))
        LXLSX_PUSH_ATTRIBUTES_STR("patternType", patterns[pattern]);

    lxlsx_xml_start_tag(self->file, "patternFill", &attributes);

    if (fg_color != LXLSX_COLOR_UNSET)
        _write_fg_color(self, fg_color);

    _write_bg_color(self, bg_color, pattern);

    lxlsx_xml_end_tag(self->file, "patternFill");
    lxlsx_xml_end_tag(self->file, "fill");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <fills> element.
 */
STATIC void
_write_fills(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->fill_count);

    lxlsx_xml_start_tag(self->file, "fills", &attributes);

    /* Write the default fills. */
    _write_default_fill(self, "none");
    _write_default_fill(self, "gray125");

    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
        if (format->has_fill)
            _write_fill(self, format, LXLSX_FALSE);
    }

    lxlsx_xml_end_tag(self->file, "fills");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the border <color> element.
 */
STATIC void
_write_border_color(lxlsx_styles *self, lxlsx_color_t color)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb_str[LXLSX_ATTR_32];

    LXLSX_INIT_ATTRIBUTES();

    if (color != LXLSX_COLOR_UNSET) {
        lxlsx_snprintf(rgb_str, LXLSX_ATTR_32, "FF%06X", color & LXLSX_COLOR_MASK);
        LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb_str);
    }
    else {
        LXLSX_PUSH_ATTRIBUTES_STR("auto", "1");
    }

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <border> sub elements such as <right>, <top>, etc.
 */
STATIC void
_write_sub_border(lxlsx_styles *self, const char *type, uint8_t style,
                  lxlsx_color_t color)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    char *border_styles[] = {
        "none",
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    };

    if (!style) {
        lxlsx_xml_empty_tag(self->file, type, NULL);
        return;
    }

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("style", border_styles[style]);

    lxlsx_xml_start_tag(self->file, type, &attributes);

    _write_border_color(self, color);

    lxlsx_xml_end_tag(self->file, type);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <border> element.
 */
STATIC void
_write_border(lxlsx_styles *self, lxlsx_format *format, uint8_t is_dxf)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Add attributes for diagonal borders. */
    if (format->diag_type == LXLSX_DIAGONAL_BORDER_UP) {
        LXLSX_PUSH_ATTRIBUTES_STR("diagonalUp", "1");
    }
    else if (format->diag_type == LXLSX_DIAGONAL_BORDER_DOWN) {
        LXLSX_PUSH_ATTRIBUTES_STR("diagonalDown", "1");
    }
    else if (format->diag_type == LXLSX_DIAGONAL_BORDER_UP_DOWN) {
        LXLSX_PUSH_ATTRIBUTES_STR("diagonalUp", "1");
        LXLSX_PUSH_ATTRIBUTES_STR("diagonalDown", "1");
    }

    /* Ensure that a default diag border is set if the diag type is set. */
    if (format->diag_type && !format->diag_border) {
        format->diag_border = LXLSX_BORDER_THIN;
    }

    /* Write the start border tag. */
    lxlsx_xml_start_tag(self->file, "border", &attributes);

    /* Write the <border> sub elements. */
    _write_sub_border(self, "left", format->left, format->left_color);
    _write_sub_border(self, "right", format->right, format->right_color);
    _write_sub_border(self, "top", format->top, format->top_color);
    _write_sub_border(self, "bottom", format->bottom, format->bottom_color);

    if (is_dxf) {
        _write_sub_border(self, "vertical", 0, LXLSX_COLOR_UNSET);
        _write_sub_border(self, "horizontal", 0, LXLSX_COLOR_UNSET);
    }

    /* Conditional DXF formats don't allow diagonal borders. */
    if (!is_dxf)
        _write_sub_border(self, "diagonal",
                          format->diag_border, format->diag_color);

    lxlsx_xml_end_tag(self->file, "border");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <borders> element.
 */
STATIC void
_write_borders(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->border_count);

    lxlsx_xml_start_tag(self->file, "borders", &attributes);

    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
        if (format->has_border)
            _write_border(self, format, LXLSX_FALSE);
    }

    lxlsx_xml_end_tag(self->file, "borders");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <alignment> element for hyperlinks.
 */
STATIC void
_write_hyperlink_alignment(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("vertical", "top");

    lxlsx_xml_empty_tag(self->file, "alignment", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <protection> element for hyperlinks.
 */
STATIC void
_write_hyperlink_protection(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("locked", "0");

    lxlsx_xml_empty_tag(self->file, "protection", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xf> element for styles.
 */
STATIC void
_write_style_xf(lxlsx_styles *self, uint8_t has_hyperlink, uint16_t font_id)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("numFmtId", "0");
    LXLSX_PUSH_ATTRIBUTES_INT("fontId", font_id);
    LXLSX_PUSH_ATTRIBUTES_STR("fillId", "0");
    LXLSX_PUSH_ATTRIBUTES_STR("borderId", "0");

    if (has_hyperlink) {
        LXLSX_PUSH_ATTRIBUTES_STR("applyNumberFormat", "0");
        LXLSX_PUSH_ATTRIBUTES_STR("applyFill", "0");
        LXLSX_PUSH_ATTRIBUTES_STR("applyBorder", "0");
        LXLSX_PUSH_ATTRIBUTES_STR("applyAlignment", "0");
        LXLSX_PUSH_ATTRIBUTES_STR("applyProtection", "0");

        lxlsx_xml_start_tag(self->file, "xf", &attributes);
        _write_hyperlink_alignment(self);
        _write_hyperlink_protection(self);
        lxlsx_xml_end_tag(self->file, "xf");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "xf", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cellStyleXfs> element.
 */
STATIC void
_write_cell_style_xfs(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (self->has_hyperlink)
        LXLSX_PUSH_ATTRIBUTES_STR("count", "2");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("count", "1");

    lxlsx_xml_start_tag(self->file, "cellStyleXfs", &attributes);
    _write_style_xf(self, LXLSX_FALSE, 0);

    if (self->has_hyperlink)
        _write_style_xf(self, self->has_hyperlink, self->hyperlink_font_id);

    lxlsx_xml_end_tag(self->file, "cellStyleXfs");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Check if a format struct has alignment properties set and the
 * "applyAlignment" attribute should be set.
 */
STATIC uint8_t
_apply_alignment(lxlsx_format *format)
{
    return format->text_h_align != LXLSX_ALIGN_NONE
        || format->text_v_align != LXLSX_ALIGN_NONE
        || format->indent != 0
        || format->rotation != 0
        || format->text_wrap != 0
        || format->shrink != 0 || format->reading_order != 0;
}

/*
 * Check if a format struct has alignment properties set apart from the
 * LXLSX_ALIGN_VERTICAL_BOTTOM which Excel treats as a default.
 */
STATIC uint8_t
_has_alignment(lxlsx_format *format)
{
    return format->text_h_align != LXLSX_ALIGN_NONE
        || !(format->text_v_align == LXLSX_ALIGN_NONE ||
             format->text_v_align == LXLSX_ALIGN_VERTICAL_BOTTOM)
        || format->indent != 0
        || format->rotation != 0
        || format->text_wrap != 0
        || format->shrink != 0 || format->reading_order != 0;
}

/*
 * Write the <alignment> element.
 */
STATIC void
_write_alignment(lxlsx_styles *self, lxlsx_format *format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    int16_t rotation = format->rotation;

    LXLSX_INIT_ATTRIBUTES();

    /* Indent is only allowed for some alignment properties. */
    /* If it is defined for any other alignment or no alignment has been  */
    /* set then default to left alignment. */
    if (format->indent
        && format->text_h_align != LXLSX_ALIGN_LEFT
        && format->text_h_align != LXLSX_ALIGN_RIGHT
        && format->text_h_align != LXLSX_ALIGN_DISTRIBUTED
        && format->text_v_align != LXLSX_ALIGN_VERTICAL_TOP
        && format->text_v_align != LXLSX_ALIGN_VERTICAL_BOTTOM
        && format->text_v_align != LXLSX_ALIGN_VERTICAL_DISTRIBUTED) {
        format->text_h_align = LXLSX_ALIGN_LEFT;
    }

    /* Check for properties that are mutually exclusive. */
    if (format->text_wrap)
        format->shrink = 0;

    if (format->text_h_align == LXLSX_ALIGN_FILL)
        format->shrink = 0;

    if (format->text_h_align == LXLSX_ALIGN_JUSTIFY)
        format->shrink = 0;

    if (format->text_h_align == LXLSX_ALIGN_DISTRIBUTED)
        format->shrink = 0;

    if (format->text_h_align != LXLSX_ALIGN_DISTRIBUTED)
        format->just_distrib = 0;

    if (format->indent)
        format->just_distrib = 0;

    if (format->text_h_align == LXLSX_ALIGN_LEFT)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "left");

    if (format->text_h_align == LXLSX_ALIGN_CENTER)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "center");

    if (format->text_h_align == LXLSX_ALIGN_RIGHT)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "right");

    if (format->text_h_align == LXLSX_ALIGN_FILL)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "fill");

    if (format->text_h_align == LXLSX_ALIGN_JUSTIFY)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "justify");

    if (format->text_h_align == LXLSX_ALIGN_CENTER_ACROSS)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "centerContinuous");

    if (format->text_h_align == LXLSX_ALIGN_DISTRIBUTED)
        LXLSX_PUSH_ATTRIBUTES_STR("horizontal", "distributed");

    if (format->just_distrib)
        LXLSX_PUSH_ATTRIBUTES_STR("justifyLastLine", "1");

    if (format->text_v_align == LXLSX_ALIGN_VERTICAL_TOP)
        LXLSX_PUSH_ATTRIBUTES_STR("vertical", "top");

    if (format->text_v_align == LXLSX_ALIGN_VERTICAL_CENTER)
        LXLSX_PUSH_ATTRIBUTES_STR("vertical", "center");

    if (format->text_v_align == LXLSX_ALIGN_VERTICAL_JUSTIFY)
        LXLSX_PUSH_ATTRIBUTES_STR("vertical", "justify");

    if (format->text_v_align == LXLSX_ALIGN_VERTICAL_DISTRIBUTED)
        LXLSX_PUSH_ATTRIBUTES_STR("vertical", "distributed");

    /* Map rotation to Excel values. */
    if (rotation) {
        if (rotation == 270)
            rotation = 255;
        else if (rotation < 0)
            rotation = -rotation + 90;

        LXLSX_PUSH_ATTRIBUTES_INT("textRotation", rotation);
    }

    if (format->indent)
        LXLSX_PUSH_ATTRIBUTES_INT("indent", format->indent);

    if (format->text_wrap)
        LXLSX_PUSH_ATTRIBUTES_STR("wrapText", "1");

    if (format->shrink)
        LXLSX_PUSH_ATTRIBUTES_STR("shrinkToFit", "1");

    if (format->reading_order == 1)
        LXLSX_PUSH_ATTRIBUTES_STR("readingOrder", "1");

    if (format->reading_order == 2)
        LXLSX_PUSH_ATTRIBUTES_STR("readingOrder", "2");

    if (!STAILQ_EMPTY(&attributes))
        lxlsx_xml_empty_tag(self->file, "alignment", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <protection> element.
 */
STATIC void
_write_protection(lxlsx_styles *self, lxlsx_format *format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (!format->locked)
        LXLSX_PUSH_ATTRIBUTES_STR("locked", "0");

    if (format->hidden)
        LXLSX_PUSH_ATTRIBUTES_STR("hidden", "1");

    lxlsx_xml_empty_tag(self->file, "protection", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xf> element.
 */
STATIC void
_write_xf(lxlsx_styles *self, lxlsx_format *format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint8_t has_protection = (!format->locked) | format->hidden;
    uint8_t has_alignment = _has_alignment(format);
    uint8_t apply_alignment = _apply_alignment(format);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("numFmtId", format->num_format_index);
    LXLSX_PUSH_ATTRIBUTES_INT("fontId", format->font_index);
    LXLSX_PUSH_ATTRIBUTES_INT("fillId", format->fill_index);
    LXLSX_PUSH_ATTRIBUTES_INT("borderId", format->border_index);
    LXLSX_PUSH_ATTRIBUTES_INT("xfId", format->xf_id);

    if (format->quote_prefix)
        LXLSX_PUSH_ATTRIBUTES_STR("quotePrefix", "1");

    if (format->num_format_index > 0)
        LXLSX_PUSH_ATTRIBUTES_STR("applyNumberFormat", "1");

    /* Add applyFont attribute if XF format uses a font element. */
    if (format->font_index > 0 && !format->hyperlink)
        LXLSX_PUSH_ATTRIBUTES_STR("applyFont", "1");

    /* Add applyFill attribute if XF format uses a fill element. */
    if (format->fill_index > 0)
        LXLSX_PUSH_ATTRIBUTES_STR("applyFill", "1");

    /* Add applyBorder attribute if XF format uses a border element. */
    if (format->border_index > 0)
        LXLSX_PUSH_ATTRIBUTES_STR("applyBorder", "1");

    /* We can also have applyAlignment without a sub-element. */
    if (apply_alignment || format->hyperlink)
        LXLSX_PUSH_ATTRIBUTES_STR("applyAlignment", "1");

    if (has_protection || format->hyperlink)
        LXLSX_PUSH_ATTRIBUTES_STR("applyProtection", "1");

    /* Write XF with sub-elements if required. */
    if (has_alignment || has_protection) {
        lxlsx_xml_start_tag(self->file, "xf", &attributes);

        if (has_alignment)
            _write_alignment(self, format);

        if (has_protection)
            _write_protection(self, format);

        lxlsx_xml_end_tag(self->file, "xf");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "xf", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cellXfs> element.
 */
STATIC void
_write_cell_xfs(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;
    uint32_t count = self->xf_count;
    uint32_t i = 0;

    /* If the last format is "font_only" it is for the comment font and
     * shouldn't be counted. This is a workaround to get the last object
     * in the list since STAILQ_LAST() requires __containerof and isn't
     * ANSI compatible. */
    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
        i++;
        if (i == self->xf_count && format->font_only)
            count--;
    }

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", count);

    lxlsx_xml_start_tag(self->file, "cellXfs", &attributes);

    STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
        if (!format->font_only)
            _write_xf(self, format);
    }

    lxlsx_xml_end_tag(self->file, "cellXfs");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cellStyle> element.
 */
STATIC void
_write_cell_style(lxlsx_styles *self, char *name, uint8_t xf_id,
                  uint8_t builtin_id)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("name", name);
    LXLSX_PUSH_ATTRIBUTES_INT("xfId", xf_id);
    LXLSX_PUSH_ATTRIBUTES_INT("builtinId", builtin_id);

    lxlsx_xml_empty_tag(self->file, "cellStyle", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cellStyles> element.
 */
STATIC void
_write_cell_styles(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    LXLSX_INIT_ATTRIBUTES();

    if (self->has_hyperlink)
        LXLSX_PUSH_ATTRIBUTES_STR("count", "2");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("count", "1");

    lxlsx_xml_start_tag(self->file, "cellStyles", &attributes);

    if (self->has_hyperlink)
        _write_cell_style(self, "Hyperlink", 1, 8);

    _write_cell_style(self, "Normal", 0, 0);

    lxlsx_xml_end_tag(self->file, "cellStyles");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <dxfs> element.
 *
 */
STATIC void
_write_dxfs(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_format *format;
    uint32_t count = self->dxf_count;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", count);

    if (count) {
        lxlsx_xml_start_tag(self->file, "dxfs", &attributes);

        STAILQ_FOREACH(format, self->dxf_formats, list_pointers) {
            lxlsx_xml_start_tag(self->file, "dxf", NULL);

            if (format->has_dxf_font)
                _write_font(self, format, LXLSX_TRUE, LXLSX_FALSE);

            if (format->num_format_index)
                _write_num_fmt(self, format->num_format_index,
                               format->num_format);

            if (format->has_dxf_fill)
                _write_fill(self, format, LXLSX_TRUE);

            if (format->has_dxf_border)
                _write_border(self, format, LXLSX_TRUE);

            lxlsx_xml_end_tag(self->file, "dxf");
        }

        lxlsx_xml_end_tag(self->file, "dxfs");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "dxfs", &attributes);
    }
    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <tableStyles> element.
 */
STATIC void
_write_table_styles(lxlsx_styles *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("count", "0");
    LXLSX_PUSH_ATTRIBUTES_STR("defaultTableStyle", "TableStyleMedium9");
    LXLSX_PUSH_ATTRIBUTES_STR("defaultPivotStyle", "PivotStyleLight16");

    lxlsx_xml_empty_tag(self->file, "tableStyles", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void
lxlsx_styles_assemble_xml_file(lxlsx_styles *self)
{
    /* Write the XML declaration. */
    _styles_xml_declaration(self);

    /* Add the style sheet. */
    _write_style_sheet(self);

    /* Write the number formats. */
    _write_num_fmts(self);

    /* Write the fonts. */
    _write_fonts(self);

    /* Write the fills. */
    _write_fills(self);

    /* Write the borders element. */
    _write_borders(self);

    /* Write the cellStyleXfs element. */
    _write_cell_style_xfs(self);

    /* Write the cellXfs element. */
    _write_cell_xfs(self);

    /* Write the cellStyles element. */
    _write_cell_styles(self);

    /* Write the dxfs element. */
    _write_dxfs(self);

    /* Write the tableStyles element. */
    _write_table_styles(self);

    /* Close the style sheet tag. */
    lxlsx_xml_end_tag(self->file, "styleSheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/


/****************************************************************************
 *
 * XLSX styles read support.
 *
 ****************************************************************************/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "styles_private.h"
#include "numfmt.h"
#include "xlsx_util.h"
#include "xml_pump.h"

typedef struct {
    uint16_t  id;
    char     *fmt;
} lxlsx_reader_user_numfmt;

struct lxlsx_reader_styles {
    lxlsx_reader_user_numfmt *user_fmts;
    size_t           user_fmts_count;
    size_t           user_fmts_cap;

    lxlsx_reader_xf          *xfs;
    size_t           xfs_count;
    size_t           xfs_cap;

    lxlsx_reader_font        *fonts;
    size_t           fonts_count;
    size_t           fonts_cap;

    lxlsx_reader_fill        *fills;
    size_t           fills_count;
    size_t           fills_cap;

    lxlsx_reader_border      *borders;
    size_t           borders_count;
    size_t           borders_cap;

    lxlsx_reader_dxf         *dxfs;
    size_t           dxfs_count;
    size_t           dxfs_cap;

    char             theme_colors[12][LXLSX_READER_COLOR_LEN];
};

/* ========================================================================= */
/* Public lookup                                                             */
/* ========================================================================= */

const lxlsx_reader_xf *lxlsx_reader_styles_get_xf(const lxlsx_reader_styles *st, uint32_t style_id)
{
    if (!st) return NULL;
    if ((size_t)style_id >= st->xfs_count) return NULL;
    return &st->xfs[style_id];
}

const lxlsx_reader_font *lxlsx_reader_styles_get_font(const lxlsx_reader_styles *st, uint32_t font_id)
{
    if (!st) return NULL;
    if ((size_t)font_id >= st->fonts_count) return NULL;
    return &st->fonts[font_id];
}

const lxlsx_reader_fill *lxlsx_reader_styles_get_fill(const lxlsx_reader_styles *st, uint32_t fill_id)
{
    if (!st) return NULL;
    if ((size_t)fill_id >= st->fills_count) return NULL;
    return &st->fills[fill_id];
}

const lxlsx_reader_border *lxlsx_reader_styles_get_border(const lxlsx_reader_styles *st, uint32_t border_id)
{
    if (!st) return NULL;
    if ((size_t)border_id >= st->borders_count) return NULL;
    return &st->borders[border_id];
}

size_t lxlsx_reader_styles_dxf_count(const lxlsx_reader_styles *st)
{
    return st ? st->dxfs_count : 0;
}

const lxlsx_reader_dxf *lxlsx_reader_styles_get_dxf(const lxlsx_reader_styles *st, uint32_t dxf_id)
{
    if (!st) return NULL;
    if ((size_t)dxf_id >= st->dxfs_count) return NULL;
    return &st->dxfs[dxf_id];
}

size_t lxlsx_reader_styles_count(const lxlsx_reader_styles *st)
{
    return st ? st->xfs_count : 0;
}

static void free_font_fields(lxlsx_reader_font *font)
{
    if (!font) return;
    free((void *)font->name);
    font->name = NULL;
}

static void free_fill_fields(lxlsx_reader_fill *fill)
{
    if (!fill) return;
    free((void *)fill->pattern_type);
    fill->pattern_type = NULL;
}

static void free_border_fields(lxlsx_reader_border *border)
{
    if (!border) return;
    free((void *)border->left.style);
    free((void *)border->right.style);
    free((void *)border->top.style);
    free((void *)border->bottom.style);
    border->left.style = NULL;
    border->right.style = NULL;
    border->top.style = NULL;
    border->bottom.style = NULL;
}

static void free_dxf_fields(lxlsx_reader_dxf *dxf)
{
    if (!dxf) return;
    free_font_fields(&dxf->font);
    free_fill_fields(&dxf->fill);
    free_border_fields(&dxf->border);
}

void lxlsx_reader_styles_free(lxlsx_reader_styles *st)
{
    size_t i;
    if (!st) return;

    for (i = 0; i < st->user_fmts_count; i++) free(st->user_fmts[i].fmt);
    free(st->user_fmts);

    for (i = 0; i < st->fonts_count; i++) free_font_fields(&st->fonts[i]);
    free(st->fonts);

    for (i = 0; i < st->fills_count; i++) free_fill_fields(&st->fills[i]);
    free(st->fills);

    for (i = 0; i < st->borders_count; i++) free_border_fields(&st->borders[i]);
    free(st->borders);

    for (i = 0; i < st->dxfs_count; i++) free_dxf_fields(&st->dxfs[i]);
    free(st->dxfs);

    for (i = 0; i < st->xfs_count; i++) {
        free((void *)st->xfs[i].alignment.horizontal);
        free((void *)st->xfs[i].alignment.vertical);
    }
    free(st->xfs);

    free(st);
}

/* ========================================================================= */
/* numFmt                                                                    */
/* ========================================================================= */

static const char *resolve_format(lxlsx_reader_styles *st, uint16_t id, lxlsx_reader_fmt_category *cat_out)
{
    size_t      i;
    const char *builtin = lxlsx_reader_numfmt_builtin_format(id);
    if (builtin) {
        if (cat_out) *cat_out = lxlsx_reader_numfmt_builtin_category(id);
        return builtin;
    }
    for (i = 0; i < st->user_fmts_count; i++) {
        if (st->user_fmts[i].id == id) {
            const char *f = st->user_fmts[i].fmt;
            if (cat_out) *cat_out = lxlsx_reader_numfmt_classify(f);
            return f;
        }
    }
    if (cat_out) *cat_out = LXLSX_READER_FMT_CATEGORY_GENERAL;
    return NULL;
}

static int append_user_fmt(lxlsx_reader_styles *st, uint16_t id, const char *fmt)
{
    if (st->user_fmts_count >= st->user_fmts_cap) {
        size_t cap = st->user_fmts_cap ? st->user_fmts_cap * 2 : 8;
        lxlsx_reader_user_numfmt *nb = (lxlsx_reader_user_numfmt *)realloc(
            st->user_fmts, cap * sizeof(*nb));
        if (!nb) return -1;
        st->user_fmts     = nb;
        st->user_fmts_cap = cap;
    }
    st->user_fmts[st->user_fmts_count].id  = id;
    st->user_fmts[st->user_fmts_count].fmt = strdup(fmt ? fmt : "");
    if (!st->user_fmts[st->user_fmts_count].fmt) return -1;
    st->user_fmts_count++;
    return 0;
}

/* ========================================================================= */
/* Array growers                                                             */
/* ========================================================================= */

#define GROW_ARRAY(arr, count, cap, type, init_cap)                          \
    do {                                                                      \
        if ((count) >= (cap)) {                                               \
            size_t _nc = (cap) ? (cap) * 2 : (init_cap);                      \
            type *_nb  = (type *)realloc((arr), _nc * sizeof(*(arr)));        \
            if (!_nb) return -1;                                              \
            (arr) = _nb;                                                      \
            (cap) = _nc;                                                      \
        }                                                                     \
    } while (0)

static void init_font(lxlsx_reader_font *font)
{
    memset(font, 0, sizeof(*font));
    font->underline = LXLSX_READER_UNDERLINE_NONE;
}

static void init_fill(lxlsx_reader_fill *fill)
{
    memset(fill, 0, sizeof(*fill));
}

static void init_border(lxlsx_reader_border *border)
{
    memset(border, 0, sizeof(*border));
}

static int begin_font(lxlsx_reader_styles *st)
{
    GROW_ARRAY(st->fonts, st->fonts_count, st->fonts_cap, lxlsx_reader_font, 8);
    init_font(&st->fonts[st->fonts_count]);
    return 0;
}

static int begin_fill(lxlsx_reader_styles *st)
{
    GROW_ARRAY(st->fills, st->fills_count, st->fills_cap, lxlsx_reader_fill, 8);
    init_fill(&st->fills[st->fills_count]);
    return 0;
}

static int begin_border(lxlsx_reader_styles *st)
{
    GROW_ARRAY(st->borders, st->borders_count, st->borders_cap, lxlsx_reader_border, 8);
    init_border(&st->borders[st->borders_count]);
    return 0;
}

static int begin_dxf(lxlsx_reader_styles *st)
{
    GROW_ARRAY(st->dxfs, st->dxfs_count, st->dxfs_cap, lxlsx_reader_dxf, 4);
    memset(&st->dxfs[st->dxfs_count], 0, sizeof(lxlsx_reader_dxf));
    init_font(&st->dxfs[st->dxfs_count].font);
    init_fill(&st->dxfs[st->dxfs_count].fill);
    init_border(&st->dxfs[st->dxfs_count].border);
    return 0;
}

static int begin_xf(lxlsx_reader_styles *st, const char **attrs)
{
    GROW_ARRAY(st->xfs, st->xfs_count, st->xfs_cap, lxlsx_reader_xf, 32);
    {
        lxlsx_reader_xf *x = &st->xfs[st->xfs_count];
        memset(x, 0, sizeof(*x));
        x->locked = 1;  /* Excel default — protection.locked is on unless set off */

        const char *id_s     = lxlsx_reader_xml_attr(attrs, "numFmtId");
        const char *font_s   = lxlsx_reader_xml_attr(attrs, "fontId");
        const char *fill_s   = lxlsx_reader_xml_attr(attrs, "fillId");
        const char *border_s = lxlsx_reader_xml_attr(attrs, "borderId");

        x->num_fmt_id = id_s     ? (uint16_t)strtoul(id_s,     NULL, 10) : 0;
        x->font_id    = font_s   ? (uint32_t)strtoul(font_s,   NULL, 10) : 0;
        x->fill_id    = fill_s   ? (uint32_t)strtoul(fill_s,   NULL, 10) : 0;
        x->border_id  = border_s ? (uint32_t)strtoul(border_s, NULL, 10) : 0;
        x->format_string = resolve_format(st, x->num_fmt_id, &x->category);
    }
    return 0;
}

/* ========================================================================= */
/* color extraction                                                          */
/* ========================================================================= */

static const char *INDEXED_COLORS[] = {
    "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00",
    "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
    "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00",
    "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
    "FF800000", "FF008000", "FF000080", "FF808000",
    "FF800080", "FF008080", "FFC0C0C0", "FF808080",
    "FF9999FF", "FF993366", "FFFFFFCC", "FFCCFFFF",
    "FF660066", "FFFF8080", "FF0066CC", "FFCCCCFF",
    "FF000080", "FFFF00FF", "FFFFFF00", "FF00FFFF",
    "FF800080", "FF800000", "FF008080", "FF0000FF",
    "FF00CCFF", "FFCCFFFF", "FFCCFFCC", "FFFFFF99",
    "FF99CCFF", "FFFF99CC", "FFCC99FF", "FFFFCC99",
    "FF3366FF", "FF33CCCC", "FF99CC00", "FFFFCC00",
    "FFFF9900", "FFFF6600", "FF666699", "FF969696",
    "FF003366", "FF339966", "FF003300", "FF333300",
    "FF993300", "FF993366", "FF333399", "FF333333"
};

static void copy_argb_color(char out[LXLSX_READER_COLOR_LEN], const char *value)
{
    size_t len;
    if (!out) return;
    out[0] = 0;
    if (!value || !*value) return;

    len = strlen(value);
    if (len == 6) {
        out[0] = 'F';
        out[1] = 'F';
        memcpy(out + 2, value, 6);
        out[8] = 0;
        return;
    }

    if (len >= 8) {
        memcpy(out, value, 8);
        out[8] = 0;
    }
}

static int hex_byte(const char *p)
{
    char buf[3];
    if (!p) return 0;
    buf[0] = p[0];
    buf[1] = p[1];
    buf[2] = 0;
    return (int)strtol(buf, NULL, 16);
}

static int clamp_byte(double value)
{
    if (value < 0.0) return 0;
    if (value > 255.0) return 255;
    return (int)(value + 0.5);
}

static int tint_channel(int channel, double tint)
{
    if (tint < 0.0)
        return clamp_byte((double)channel * (1.0 + tint));
    return clamp_byte((double)channel * (1.0 - tint) + 255.0 * tint);
}

static void apply_tint(char color[LXLSX_READER_COLOR_LEN], double tint)
{
    int r, g, b;
    if (!color || strlen(color) < 8 || tint == 0.0) return;

    r = hex_byte(color + 2);
    g = hex_byte(color + 4);
    b = hex_byte(color + 6);

    snprintf(color + 2, 7, "%02X%02X%02X",
             tint_channel(r, tint),
             tint_channel(g, tint),
             tint_channel(b, tint));
}

static void init_default_theme(lxlsx_reader_styles *st)
{
    static const char *DEFAULT_THEME[12] = {
        "FFFFFFFF", "FF000000", "FFEEECE1", "FF1F497D",
        "FF4F81BD", "FFC0504D", "FF9BBB59", "FF8064A2",
        "FF4BACC6", "FFF79646", "FF0000FF", "FF800080"
    };
    size_t i;
    if (!st) return;
    for (i = 0; i < 12; i++)
        copy_argb_color(st->theme_colors[i], DEFAULT_THEME[i]);
}

static int theme_slot_for_name(const char *name)
{
    if (!name) return -1;
    if (lxlsx_reader_xml_name_eq(name, "lt1")) return 0;
    if (lxlsx_reader_xml_name_eq(name, "dk1")) return 1;
    if (lxlsx_reader_xml_name_eq(name, "lt2")) return 2;
    if (lxlsx_reader_xml_name_eq(name, "dk2")) return 3;
    if (lxlsx_reader_xml_name_eq(name, "accent1")) return 4;
    if (lxlsx_reader_xml_name_eq(name, "accent2")) return 5;
    if (lxlsx_reader_xml_name_eq(name, "accent3")) return 6;
    if (lxlsx_reader_xml_name_eq(name, "accent4")) return 7;
    if (lxlsx_reader_xml_name_eq(name, "accent5")) return 8;
    if (lxlsx_reader_xml_name_eq(name, "accent6")) return 9;
    if (lxlsx_reader_xml_name_eq(name, "hlink")) return 10;
    if (lxlsx_reader_xml_name_eq(name, "folHlink")) return 11;
    return -1;
}

typedef struct {
    lxlsx_reader_styles *st;
    int in_clr_scheme;
    int current_slot;
} theme_ctx;

static void theme_on_start(void *ud, const char *name, const char **attrs)
{
    theme_ctx *c = (theme_ctx *)ud;
    int slot;
    const char *v;

    if (lxlsx_reader_xml_name_eq(name, "clrScheme")) {
        c->in_clr_scheme = 1;
        return;
    }

    if (!c->in_clr_scheme)
        return;

    slot = theme_slot_for_name(name);
    if (slot >= 0) {
        c->current_slot = slot;
        return;
    }

    if (c->current_slot < 0)
        return;

    if (lxlsx_reader_xml_name_eq(name, "srgbClr")) {
        v = lxlsx_reader_xml_attr(attrs, "val");
        if (v)
            copy_argb_color(c->st->theme_colors[c->current_slot], v);
    } else if (lxlsx_reader_xml_name_eq(name, "sysClr")) {
        v = lxlsx_reader_xml_attr(attrs, "lastClr");
        if (v)
            copy_argb_color(c->st->theme_colors[c->current_slot], v);
    }
}

static void theme_on_end(void *ud, const char *name)
{
    theme_ctx *c = (theme_ctx *)ud;
    if (lxlsx_reader_xml_name_eq(name, "clrScheme")) {
        c->in_clr_scheme = 0;
        c->current_slot = -1;
        return;
    }
    if (c->current_slot >= 0 && theme_slot_for_name(name) == c->current_slot)
        c->current_slot = -1;
}

typedef struct {
    char *path;
} theme_find_ctx;

static int find_theme_entry_cb(const char *name, void *userdata)
{
    theme_find_ctx *ctx = (theme_find_ctx *)userdata;
    size_t len;

    if (ctx->path || !name)
        return 0;

    len = strlen(name);
    if (strncmp(name, "xl/theme/", 9) == 0 && len > 4 &&
        lxlsx_reader_ascii_case_eq(name + len - 4, ".xml")) {
        ctx->path = strdup(name);
        return ctx->path ? 1 : 0;
    }
    return 0;
}

static void load_theme_colors(lxlsx_reader_zip *zip, lxlsx_reader_styles *st)
{
    theme_find_ctx find_ctx;
    char *theme_path;
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    theme_ctx ctx;

    init_default_theme(st);
    if (!zip || !st)
        return;

    memset(&find_ctx, 0, sizeof(find_ctx));
    (void)lxlsx_reader_zip_iterate_entries(zip, find_theme_entry_cb, &find_ctx);
    theme_path = find_ctx.path ? find_ctx.path : strdup("xl/theme/theme1.xml");
    if (!theme_path)
        return;

    zf = lxlsx_reader_zip_open_entry(zip, theme_path);
    free(theme_path);
    if (!zf)
        return;

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxlsx_reader_zip_close_entry(zf);
        return;
    }

    memset(&ctx, 0, sizeof(ctx));
    ctx.st = st;
    ctx.current_slot = -1;
    lxlsx_reader_xml_pump_set_handlers(pump, theme_on_start, theme_on_end, NULL, &ctx);
    (void)lxlsx_reader_xml_pump_run(pump);

    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
}

static void extract_color(const lxlsx_reader_styles *st,
                          const char **attrs,
                          char out[LXLSX_READER_COLOR_LEN])
{
    const char *rgb;
    const char *indexed;
    const char *theme;
    const char *tint;
    int have_color = 0;

    out[0] = 0;
    if (!attrs) return;

    rgb = lxlsx_reader_xml_attr(attrs, "rgb");
    if (rgb) {
        copy_argb_color(out, rgb);
        have_color = out[0] != 0;
    }

    indexed = lxlsx_reader_xml_attr(attrs, "indexed");
    if (!have_color && indexed) {
        unsigned long idx = strtoul(indexed, NULL, 10);
        if (idx < sizeof(INDEXED_COLORS) / sizeof(INDEXED_COLORS[0])) {
            copy_argb_color(out, INDEXED_COLORS[idx]);
            have_color = out[0] != 0;
        }
    }

    theme = lxlsx_reader_xml_attr(attrs, "theme");
    if (!have_color && theme && st) {
        unsigned long idx = strtoul(theme, NULL, 10);
        if (idx < 12 && st->theme_colors[idx][0]) {
            copy_argb_color(out, st->theme_colors[idx]);
            have_color = out[0] != 0;
        }
    }

    tint = lxlsx_reader_xml_attr(attrs, "tint");
    if (have_color && tint) {
        apply_tint(out, strtod(tint, NULL));
    }
}

static int parse_bool_attr(const char **attrs)
{
    const char *v = lxlsx_reader_xml_attr(attrs, "val");
    if (!v) return 1;          /* presence alone == true */
    if (*v == '0' || *v == 'f' || *v == 'F') return 0;
    return 1;
}

/* ========================================================================= */
/* SAX state machine                                                         */
/* ========================================================================= */

typedef enum {
    SC_NONE,
    SC_NUMFMTS,
    SC_FONTS,
    SC_FONT,
    SC_FILLS,
    SC_FILL,
    SC_PATTERN_FILL,
    SC_BORDERS,
    SC_BORDER,
    SC_BORDER_LEFT,
    SC_BORDER_RIGHT,
    SC_BORDER_TOP,
    SC_BORDER_BOTTOM,
    SC_CELL_XFS,
    SC_CELL_STYLE_XFS,
    SC_XF,
    SC_DXFS,
    SC_DXF,
    SC_DXF_FONT,
    SC_DXF_FILL,
    SC_DXF_PATTERN_FILL,
    SC_DXF_BORDER,
    SC_DXF_BORDER_LEFT,
    SC_DXF_BORDER_RIGHT,
    SC_DXF_BORDER_TOP,
    SC_DXF_BORDER_BOTTOM
} scope_t;

typedef struct {
    lxlsx_reader_styles *st;
    scope_t     stack[16];
    int         depth;
    int         in_alignment;        /* parsed inside SC_XF */
    int         in_protection;       /* parsed inside SC_XF */
} parse_ctx;

static scope_t top(parse_ctx *c) { return c->depth > 0 ? c->stack[c->depth - 1] : SC_NONE; }
static void push(parse_ctx *c, scope_t s) {
    if (c->depth < (int)(sizeof(c->stack)/sizeof(c->stack[0]))) c->stack[c->depth++] = s;
}
static void pop(parse_ctx *c) { if (c->depth > 0) c->depth--; }

static lxlsx_reader_dxf *current_dxf(parse_ctx *c)
{
    if (!c || c->st->dxfs_count >= c->st->dxfs_cap) return NULL;
    return &c->st->dxfs[c->st->dxfs_count];
}

static lxlsx_reader_border_side *border_side_for_scope(lxlsx_reader_border *b, scope_t scope)
{
    if (!b) return NULL;
    switch (scope) {
    case SC_BORDER_LEFT:
    case SC_DXF_BORDER_LEFT:
        return &b->left;
    case SC_BORDER_RIGHT:
    case SC_DXF_BORDER_RIGHT:
        return &b->right;
    case SC_BORDER_TOP:
    case SC_DXF_BORDER_TOP:
        return &b->top;
    case SC_BORDER_BOTTOM:
    case SC_DXF_BORDER_BOTTOM:
        return &b->bottom;
    default:
        return NULL;
    }
}

static lxlsx_reader_border_side *current_side(parse_ctx *c)
{
    scope_t scope = top(c);
    if (scope == SC_BORDER_LEFT || scope == SC_BORDER_RIGHT ||
        scope == SC_BORDER_TOP || scope == SC_BORDER_BOTTOM) {
        return border_side_for_scope(&c->st->borders[c->st->borders_count], scope);
    }
    if (scope == SC_DXF_BORDER_LEFT || scope == SC_DXF_BORDER_RIGHT ||
        scope == SC_DXF_BORDER_TOP || scope == SC_DXF_BORDER_BOTTOM) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        return dxf ? border_side_for_scope(&dxf->border, scope) : NULL;
    }
    return NULL;
}

static int border_side_scope(const char *name, int is_dxf, scope_t *out)
{
    if (!out) return 0;
    if (lxlsx_reader_xml_name_eq(name, "left")) {
        *out = is_dxf ? SC_DXF_BORDER_LEFT : SC_BORDER_LEFT;
        return 1;
    }
    if (lxlsx_reader_xml_name_eq(name, "right")) {
        *out = is_dxf ? SC_DXF_BORDER_RIGHT : SC_BORDER_RIGHT;
        return 1;
    }
    if (lxlsx_reader_xml_name_eq(name, "top")) {
        *out = is_dxf ? SC_DXF_BORDER_TOP : SC_BORDER_TOP;
        return 1;
    }
    if (lxlsx_reader_xml_name_eq(name, "bottom")) {
        *out = is_dxf ? SC_DXF_BORDER_BOTTOM : SC_BORDER_BOTTOM;
        return 1;
    }
    return 0;
}

static void parse_font_child(const lxlsx_reader_styles *st,
                             lxlsx_reader_font *f,
                             const char *name,
                             const char **attrs)
{
    if (!f) return;
    if (lxlsx_reader_xml_name_eq(name, "sz")) {
        const char *v = lxlsx_reader_xml_attr(attrs, "val");
        if (v) f->size = strtod(v, NULL);
    } else if (lxlsx_reader_xml_name_eq(name, "name") || lxlsx_reader_xml_name_eq(name, "rFont")) {
        const char *v = lxlsx_reader_xml_attr(attrs, "val");
        if (v) { free((void *)f->name); f->name = strdup(v); }
    } else if (lxlsx_reader_xml_name_eq(name, "color")) {
        extract_color(st, attrs, f->color);
    } else if (lxlsx_reader_xml_name_eq(name, "b")) {
        f->bold = parse_bool_attr(attrs);
    } else if (lxlsx_reader_xml_name_eq(name, "i")) {
        f->italic = parse_bool_attr(attrs);
    } else if (lxlsx_reader_xml_name_eq(name, "strike")) {
        f->strike = parse_bool_attr(attrs);
    } else if (lxlsx_reader_xml_name_eq(name, "u")) {
        const char *v = lxlsx_reader_xml_attr(attrs, "val");
        if (!v || strcmp(v, "single") == 0)               f->underline = LXLSX_READER_UNDERLINE_SINGLE;
        else if (strcmp(v, "double") == 0)                f->underline = LXLSX_READER_UNDERLINE_DOUBLE;
        else if (strcmp(v, "singleAccounting") == 0)      f->underline = LXLSX_READER_UNDERLINE_SINGLE_ACCOUNTING;
        else if (strcmp(v, "doubleAccounting") == 0)      f->underline = LXLSX_READER_UNDERLINE_DOUBLE_ACCOUNTING;
        else                                              f->underline = LXLSX_READER_UNDERLINE_SINGLE;
    }
}

static void parse_pattern_fill_start(lxlsx_reader_fill *fill, const char **attrs)
{
    const char *pt;
    if (!fill) return;
    pt = lxlsx_reader_xml_attr(attrs, "patternType");
    if (pt) { free((void *)fill->pattern_type); fill->pattern_type = strdup(pt); }
}

static void parse_fill_color_child(const lxlsx_reader_styles *st,
                                   lxlsx_reader_fill *fill,
                                   const char *name,
                                   const char **attrs)
{
    if (!fill) return;
    if (lxlsx_reader_xml_name_eq(name, "fgColor"))
        extract_color(st, attrs, fill->fg_color);
    else if (lxlsx_reader_xml_name_eq(name, "bgColor"))
        extract_color(st, attrs, fill->bg_color);
}

static int parse_border_side_start(lxlsx_reader_border *border,
                                   const char *name,
                                   const char **attrs,
                                   int is_dxf,
                                   scope_t *out_scope)
{
    lxlsx_reader_border_side *side;
    if (!border_side_scope(name, is_dxf, out_scope)) return 0;
    side = border_side_for_scope(border, *out_scope);
    if (side) {
        const char *st_attr = lxlsx_reader_xml_attr(attrs, "style");
        if (st_attr) { free((void *)side->style); side->style = strdup(st_attr); }
    }
    return 1;
}

static void on_start(void *ud, const char *name, const char **attrs)
{
    parse_ctx *c  = (parse_ctx *)ud;
    scope_t    sc = top(c);

    /* --- numFmts ------------------------------------------------------ */
    if (lxlsx_reader_xml_name_eq(name, "numFmts"))    { push(c, SC_NUMFMTS);   return; }

    if (sc == SC_NUMFMTS && lxlsx_reader_xml_name_eq(name, "numFmt")) {
        const char *id_s  = lxlsx_reader_xml_attr(attrs, "numFmtId");
        const char *fmt_s = lxlsx_reader_xml_attr(attrs, "formatCode");
        if (id_s && fmt_s) append_user_fmt(c->st, (uint16_t)strtoul(id_s, NULL, 10), fmt_s);
        return;
    }

    /* --- dxfs --------------------------------------------------------- */
    if (lxlsx_reader_xml_name_eq(name, "dxfs")) { push(c, SC_DXFS); return; }
    if (sc == SC_DXFS && lxlsx_reader_xml_name_eq(name, "dxf")) {
        begin_dxf(c->st);
        push(c, SC_DXF);
        return;
    }
    if (sc == SC_DXF) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        if (dxf && lxlsx_reader_xml_name_eq(name, "font")) {
            dxf->has_font = 1;
            push(c, SC_DXF_FONT);
        } else if (dxf && lxlsx_reader_xml_name_eq(name, "fill")) {
            dxf->has_fill = 1;
            push(c, SC_DXF_FILL);
        } else if (dxf && lxlsx_reader_xml_name_eq(name, "border")) {
            dxf->has_border = 1;
            push(c, SC_DXF_BORDER);
        }
        return;
    }
    if (sc == SC_DXF_FONT) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        if (dxf) parse_font_child(c->st, &dxf->font, name, attrs);
        return;
    }
    if (sc == SC_DXF_FILL && lxlsx_reader_xml_name_eq(name, "patternFill")) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        if (dxf) parse_pattern_fill_start(&dxf->fill, attrs);
        push(c, SC_DXF_PATTERN_FILL);
        return;
    }
    if (sc == SC_DXF_PATTERN_FILL) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        if (dxf) parse_fill_color_child(c->st, &dxf->fill, name, attrs);
        return;
    }
    if (sc == SC_DXF_BORDER) {
        lxlsx_reader_dxf *dxf = current_dxf(c);
        scope_t ns;
        if (dxf && parse_border_side_start(&dxf->border, name, attrs, 1, &ns))
            push(c, ns);
        return;
    }
    if (sc == SC_DXF_BORDER_LEFT || sc == SC_DXF_BORDER_RIGHT ||
        sc == SC_DXF_BORDER_TOP || sc == SC_DXF_BORDER_BOTTOM) {
        if (lxlsx_reader_xml_name_eq(name, "color")) {
            lxlsx_reader_border_side *side = current_side(c);
            if (side) extract_color(c->st, attrs, side->color);
        }
        return;
    }

    /* --- fonts -------------------------------------------------------- */
    if (lxlsx_reader_xml_name_eq(name, "fonts"))      { push(c, SC_FONTS);     return; }
    if (sc == SC_FONTS && lxlsx_reader_xml_name_eq(name, "font")) {
        begin_font(c->st);
        push(c, SC_FONT);
        return;
    }
    if (sc == SC_FONT) {
        lxlsx_reader_font *f = &c->st->fonts[c->st->fonts_count];
        parse_font_child(c->st, f, name, attrs);
        return;
    }

    /* --- fills -------------------------------------------------------- */
    if (lxlsx_reader_xml_name_eq(name, "fills"))      { push(c, SC_FILLS);     return; }
    if (sc == SC_FILLS && lxlsx_reader_xml_name_eq(name, "fill")) {
        begin_fill(c->st);
        push(c, SC_FILL);
        return;
    }
    if (sc == SC_FILL && lxlsx_reader_xml_name_eq(name, "patternFill")) {
        lxlsx_reader_fill   *fl  = &c->st->fills[c->st->fills_count];
        parse_pattern_fill_start(fl, attrs);
        push(c, SC_PATTERN_FILL);
        return;
    }
    if (sc == SC_PATTERN_FILL) {
        lxlsx_reader_fill *fl = &c->st->fills[c->st->fills_count];
        parse_fill_color_child(c->st, fl, name, attrs);
        return;
    }

    /* --- borders ------------------------------------------------------ */
    if (lxlsx_reader_xml_name_eq(name, "borders"))    { push(c, SC_BORDERS);   return; }
    if (sc == SC_BORDERS && lxlsx_reader_xml_name_eq(name, "border")) {
        begin_border(c->st);
        push(c, SC_BORDER);
        return;
    }
    if (sc == SC_BORDER) {
        lxlsx_reader_border *b   = &c->st->borders[c->st->borders_count];
        scope_t          ns;

        if (parse_border_side_start(b, name, attrs, 0, &ns)) {
            push(c, ns);
        }
        return;
    }
    if (sc == SC_BORDER_LEFT || sc == SC_BORDER_RIGHT
        || sc == SC_BORDER_TOP || sc == SC_BORDER_BOTTOM) {
        if (lxlsx_reader_xml_name_eq(name, "color")) {
            lxlsx_reader_border_side *side = current_side(c);
            if (side) extract_color(c->st, attrs, side->color);
        }
        return;
    }

    /* --- cellXfs ------------------------------------------------------ */
    if (lxlsx_reader_xml_name_eq(name, "cellXfs"))      { push(c, SC_CELL_XFS);       return; }
    if (lxlsx_reader_xml_name_eq(name, "cellStyleXfs")) { push(c, SC_CELL_STYLE_XFS); return; }

    if (sc == SC_CELL_XFS && lxlsx_reader_xml_name_eq(name, "xf")) {
        begin_xf(c->st, attrs);
        push(c, SC_XF);
        return;
    }

    if (sc == SC_XF) {
        lxlsx_reader_xf *x = &c->st->xfs[c->st->xfs_count];
        if (lxlsx_reader_xml_name_eq(name, "alignment")) {
            const char *h = lxlsx_reader_xml_attr(attrs, "horizontal");
            const char *v = lxlsx_reader_xml_attr(attrs, "vertical");
            const char *w = lxlsx_reader_xml_attr(attrs, "wrapText");
            const char *i = lxlsx_reader_xml_attr(attrs, "indent");
            const char *r = lxlsx_reader_xml_attr(attrs, "textRotation");
            x->has_alignment = 1;
            if (h) { free((void *)x->alignment.horizontal); x->alignment.horizontal = strdup(h); }
            if (v) { free((void *)x->alignment.vertical);   x->alignment.vertical   = strdup(v); }
            if (w) x->alignment.wrap_text = (strcmp(w, "1") == 0 || strcmp(w, "true") == 0);
            if (i) x->alignment.indent    = (int)strtol(i, NULL, 10);
            if (r) x->alignment.rotation  = (int)strtol(r, NULL, 10);
            return;
        }
        if (lxlsx_reader_xml_name_eq(name, "protection")) {
            const char *l = lxlsx_reader_xml_attr(attrs, "locked");
            const char *h = lxlsx_reader_xml_attr(attrs, "hidden");
            if (l) x->locked = (strcmp(l, "1") == 0 || strcmp(l, "true") == 0);
            if (h) x->hidden = (strcmp(h, "1") == 0 || strcmp(h, "true") == 0);
            return;
        }
    }
}

static void on_end(void *ud, const char *name)
{
    parse_ctx *c = (parse_ctx *)ud;
    scope_t    sc = top(c);

    if (lxlsx_reader_xml_name_eq(name, "numFmts")    && sc == SC_NUMFMTS)    pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "font")   && sc == SC_DXF_FONT)  pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "patternFill") && sc == SC_DXF_PATTERN_FILL) pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "fill")   && sc == SC_DXF_FILL)  pop(c);
    else if ((lxlsx_reader_xml_name_eq(name, "left")   && sc == SC_DXF_BORDER_LEFT)
          || (lxlsx_reader_xml_name_eq(name, "right")  && sc == SC_DXF_BORDER_RIGHT)
          || (lxlsx_reader_xml_name_eq(name, "top")    && sc == SC_DXF_BORDER_TOP)
          || (lxlsx_reader_xml_name_eq(name, "bottom") && sc == SC_DXF_BORDER_BOTTOM)) {
        pop(c);
    }
    else if (lxlsx_reader_xml_name_eq(name, "border") && sc == SC_DXF_BORDER) pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "dxf")    && sc == SC_DXF)       { c->st->dxfs_count++;    pop(c); }
    else if (lxlsx_reader_xml_name_eq(name, "dxfs")   && sc == SC_DXFS)      pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "font")  && sc == SC_FONT)       { c->st->fonts_count++;   pop(c); }
    else if (lxlsx_reader_xml_name_eq(name, "fonts") && sc == SC_FONTS)      pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "patternFill") && sc == SC_PATTERN_FILL) pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "fill")  && sc == SC_FILL)       { c->st->fills_count++;   pop(c); }
    else if (lxlsx_reader_xml_name_eq(name, "fills") && sc == SC_FILLS)      pop(c);
    else if ((lxlsx_reader_xml_name_eq(name, "left")   && sc == SC_BORDER_LEFT)
          || (lxlsx_reader_xml_name_eq(name, "right")  && sc == SC_BORDER_RIGHT)
          || (lxlsx_reader_xml_name_eq(name, "top")    && sc == SC_BORDER_TOP)
          || (lxlsx_reader_xml_name_eq(name, "bottom") && sc == SC_BORDER_BOTTOM)) {
        pop(c);
    }
    else if (lxlsx_reader_xml_name_eq(name, "border")  && sc == SC_BORDER)   { c->st->borders_count++; pop(c); }
    else if (lxlsx_reader_xml_name_eq(name, "borders") && sc == SC_BORDERS)  pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "xf") && sc == SC_XF)            { c->st->xfs_count++;     pop(c); }
    else if (lxlsx_reader_xml_name_eq(name, "cellXfs") && sc == SC_CELL_XFS) pop(c);
    else if (lxlsx_reader_xml_name_eq(name, "cellStyleXfs") && sc == SC_CELL_STYLE_XFS) pop(c);
}

/* ========================================================================= */
/* entry point                                                               */
/* ========================================================================= */

lxlsx_reader_error lxlsx_reader_styles_load(lxlsx_reader_zip *zip, const char *path, lxlsx_reader_styles **out)
{
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    lxlsx_reader_styles   *st;
    parse_ctx     ctx;
    lxlsx_reader_error     rc;

    if (!zip || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;

    st = (lxlsx_reader_styles *)calloc(1, sizeof(*st));
    if (!st) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    load_theme_colors(zip, st);

    if (!path) {
        *out = st;
        return LXLSX_READER_NO_ERROR;
    }

    zf = lxlsx_reader_zip_open_entry(zip, path);
    if (!zf) {
        *out = st;
        return LXLSX_READER_NO_ERROR;
    }

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxlsx_reader_zip_close_entry(zf);
        free(st);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }

    memset(&ctx, 0, sizeof(ctx));
    ctx.st = st;

    lxlsx_reader_xml_pump_set_handlers(pump, on_start, on_end, NULL, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);

    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);

    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_styles_free(st);
        return rc;
    }
    *out = st;
    return LXLSX_READER_NO_ERROR;
}


/****************************************************************************
 *
 * XLSX number format classification.
 *
 ****************************************************************************/

#include <ctype.h>
#include <string.h>

#include "numfmt.h"

/* OOXML 18.8.30 built-in numFmtIds. Locale-dependent slots
 * (5..8, 23..36, 41..44) are intentionally NULL — they fall through to the
 * GENERAL category and are treated as numbers. */
static const struct {
    uint16_t        id;
    const char      *fmt;
    lxlsx_reader_fmt_category cat;
} BUILTIN_NUMFMTS[] = {
    {  0, "General",           LXLSX_READER_FMT_CATEGORY_GENERAL  },
    {  1, "0",                 LXLSX_READER_FMT_CATEGORY_NUMBER   },
    {  2, "0.00",              LXLSX_READER_FMT_CATEGORY_NUMBER   },
    {  3, "#,##0",             LXLSX_READER_FMT_CATEGORY_NUMBER   },
    {  4, "#,##0.00",          LXLSX_READER_FMT_CATEGORY_NUMBER   },
    {  9, "0%",                LXLSX_READER_FMT_CATEGORY_PERCENT  },
    { 10, "0.00%",             LXLSX_READER_FMT_CATEGORY_PERCENT  },
    { 11, "0.00E+00",          LXLSX_READER_FMT_CATEGORY_NUMBER   },
    { 12, "# ?/?",             LXLSX_READER_FMT_CATEGORY_NUMBER   },
    { 13, "# ?\?/?\?",         LXLSX_READER_FMT_CATEGORY_NUMBER   },
    { 14, "m/d/yyyy",          LXLSX_READER_FMT_CATEGORY_DATE     },
    { 15, "d-mmm-yy",          LXLSX_READER_FMT_CATEGORY_DATE     },
    { 16, "d-mmm",             LXLSX_READER_FMT_CATEGORY_DATE     },
    { 17, "mmm-yy",            LXLSX_READER_FMT_CATEGORY_DATE     },
    { 18, "h:mm AM/PM",        LXLSX_READER_FMT_CATEGORY_TIME     },
    { 19, "h:mm:ss AM/PM",     LXLSX_READER_FMT_CATEGORY_TIME     },
    { 20, "h:mm",              LXLSX_READER_FMT_CATEGORY_TIME     },
    { 21, "h:mm:ss",           LXLSX_READER_FMT_CATEGORY_TIME     },
    { 22, "m/d/yyyy h:mm",     LXLSX_READER_FMT_CATEGORY_DATETIME },
    { 37, "#,##0 ;(#,##0)",                LXLSX_READER_FMT_CATEGORY_NUMBER },
    { 38, "#,##0 ;[Red](#,##0)",           LXLSX_READER_FMT_CATEGORY_NUMBER },
    { 39, "#,##0.00;(#,##0.00)",           LXLSX_READER_FMT_CATEGORY_NUMBER },
    { 40, "#,##0.00;[Red](#,##0.00)",      LXLSX_READER_FMT_CATEGORY_NUMBER },
    { 45, "mm:ss",             LXLSX_READER_FMT_CATEGORY_TIME     },
    { 46, "[h]:mm:ss",         LXLSX_READER_FMT_CATEGORY_TIME     },
    { 47, "mmss.0",            LXLSX_READER_FMT_CATEGORY_TIME     },
    { 48, "##0.0E+0",          LXLSX_READER_FMT_CATEGORY_NUMBER   },
    { 49, "@",                 LXLSX_READER_FMT_CATEGORY_TEXT     },
};

const char *lxlsx_reader_numfmt_builtin_format(uint16_t id)
{
    size_t i;
    for (i = 0; i < sizeof(BUILTIN_NUMFMTS) / sizeof(BUILTIN_NUMFMTS[0]); i++) {
        if (BUILTIN_NUMFMTS[i].id == id) return BUILTIN_NUMFMTS[i].fmt;
    }
    return NULL;
}

lxlsx_reader_fmt_category lxlsx_reader_numfmt_builtin_category(uint16_t id)
{
    size_t i;
    for (i = 0; i < sizeof(BUILTIN_NUMFMTS) / sizeof(BUILTIN_NUMFMTS[0]); i++) {
        if (BUILTIN_NUMFMTS[i].id == id) return BUILTIN_NUMFMTS[i].cat;
    }
    return LXLSX_READER_FMT_CATEGORY_GENERAL;
}

/* Classify a custom format string into a category. Heuristic — strict from
 * date/time markers down to general. Skips characters inside quoted literals
 * ("..."), bracketed colour/condition tags ([Red], [>=0]), and escaped
 * characters (\x). */
lxlsx_reader_fmt_category lxlsx_reader_numfmt_classify(const char *fmt)
{
    int has_y = 0, has_m_letter = 0, has_d = 0;
    int has_h = 0, has_s = 0;
    int has_percent = 0, has_currency = 0, has_at = 0;
    int has_digit = 0;
    const char *p;

    if (!fmt) return LXLSX_READER_FMT_CATEGORY_GENERAL;

    for (p = fmt; *p; p++) {
        char c = *p;
        if (c == '"') {
            p++;
            while (*p && *p != '"') p++;
            if (!*p) break;
            continue;
        }
        if (c == '[') {
            while (*p && *p != ']') p++;
            if (!*p) break;
            continue;
        }
        if (c == '\\' && *(p + 1)) {
            p++;
            continue;
        }
        switch (c) {
        case 'y': case 'Y': has_y = 1; break;
        case 'd': case 'D': has_d = 1; break;
        case 'm': case 'M': has_m_letter = 1; break;
        case 'h': case 'H': has_h = 1; break;
        case 's': case 'S': has_s = 1; break;
        case '%':           has_percent  = 1; break;
        case '@':           has_at       = 1; break;
        case '$': case '\xa3': /* £ */
                            has_currency = 1; break;
        case '0': case '#': case '?': has_digit = 1; break;
        default: break;
        }
    }

    /* Multi-byte currency symbols (UTF-8 encoded ¥, €) — coarse detection:
     * any 0xC2..0xF4 byte not matched above is treated as a hint of currency. */
    for (p = fmt; *p; p++) {
        unsigned char b = (unsigned char)*p;
        if (b >= 0xC2 && b <= 0xF4) { has_currency = 1; break; }
    }

    if (has_y || has_d) {
        if (has_h || has_s) return LXLSX_READER_FMT_CATEGORY_DATETIME;
        return LXLSX_READER_FMT_CATEGORY_DATE;
    }
    if (has_h || has_s) return LXLSX_READER_FMT_CATEGORY_TIME;
    if (has_m_letter && (has_h || has_s)) return LXLSX_READER_FMT_CATEGORY_TIME;

    if (has_percent)  return LXLSX_READER_FMT_CATEGORY_PERCENT;
    if (has_currency) return LXLSX_READER_FMT_CATEGORY_CURRENCY;
    if (has_at && !has_digit) return LXLSX_READER_FMT_CATEGORY_TEXT;
    if (has_digit)    return LXLSX_READER_FMT_CATEGORY_NUMBER;

    return LXLSX_READER_FMT_CATEGORY_GENERAL;
}
