/*****************************************************************************
 * styles - A library for creating Excel XLSX styles files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/styles.h"
#include "lxlsx/utility.h"

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
