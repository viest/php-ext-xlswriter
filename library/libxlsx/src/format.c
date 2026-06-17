/*****************************************************************************
 * format - A library for creating Excel XLSX format files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/format.h"
#include "lxlsx/utility.h"

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new format object.
 */
lxlsx_format *
lxlsx_format_new(void)
{
    lxlsx_format *format = calloc(1, sizeof(lxlsx_format));
    GOTO_LABEL_ON_MEM_ERROR(format, mem_error);

    format->xf_format_indices = NULL;
    format->dxf_format_indices = NULL;

    format->xf_index = LXLSX_PROPERTY_UNSET;
    format->dxf_index = LXLSX_PROPERTY_UNSET;
    format->xf_id = 0;

    format->font_name[0] = '\0';
    format->font_scheme[0] = '\0';
    format->num_format[0] = '\0';
    format->num_format_index = 0;
    format->font_index = 0;
    format->has_font = LXLSX_FALSE;
    format->has_dxf_font = LXLSX_FALSE;
    format->font_size = 11.0;
    format->bold = LXLSX_FALSE;
    format->italic = LXLSX_FALSE;
    format->font_color = LXLSX_COLOR_UNSET;
    format->underline = LXLSX_UNDERLINE_NONE;
    format->font_strikeout = LXLSX_FALSE;
    format->font_outline = LXLSX_FALSE;
    format->font_shadow = LXLSX_FALSE;
    format->font_script = LXLSX_FALSE;
    format->font_family = LXLSX_DEFAULT_FONT_FAMILY;
    format->font_charset = LXLSX_FALSE;
    format->font_condense = LXLSX_FALSE;
    format->font_extend = LXLSX_FALSE;
    format->theme = 0;
    format->hyperlink = LXLSX_FALSE;

    format->hidden = LXLSX_FALSE;
    format->locked = LXLSX_TRUE;

    format->text_h_align = LXLSX_ALIGN_NONE;
    format->text_wrap = LXLSX_FALSE;
    format->text_v_align = LXLSX_ALIGN_NONE;
    format->text_justlast = LXLSX_FALSE;
    format->rotation = 0;

    format->fg_color = LXLSX_COLOR_UNSET;
    format->bg_color = LXLSX_COLOR_UNSET;
    format->pattern = LXLSX_PATTERN_NONE;
    format->has_fill = LXLSX_FALSE;
    format->has_dxf_fill = LXLSX_FALSE;
    format->fill_index = 0;
    format->fill_count = 0;

    format->border_index = 0;
    format->has_border = LXLSX_FALSE;
    format->has_dxf_border = LXLSX_FALSE;
    format->border_count = 0;

    format->bottom = LXLSX_BORDER_NONE;
    format->left = LXLSX_BORDER_NONE;
    format->right = LXLSX_BORDER_NONE;
    format->top = LXLSX_BORDER_NONE;
    format->diag_border = LXLSX_BORDER_NONE;
    format->diag_type = LXLSX_BORDER_NONE;
    format->bottom_color = LXLSX_COLOR_UNSET;
    format->left_color = LXLSX_COLOR_UNSET;
    format->right_color = LXLSX_COLOR_UNSET;
    format->top_color = LXLSX_COLOR_UNSET;
    format->diag_color = LXLSX_COLOR_UNSET;

    format->indent = 0;
    format->shrink = LXLSX_FALSE;
    format->merge_range = LXLSX_FALSE;
    format->reading_order = 0;
    format->just_distrib = LXLSX_FALSE;
    format->color_indexed = LXLSX_FALSE;
    format->font_only = LXLSX_FALSE;

    format->quote_prefix = LXLSX_FALSE;

    return format;

mem_error:
    lxlsx_format_free(format);
    return NULL;
}

/*
 * Free a format object.
 */
void
lxlsx_format_free(lxlsx_format *format)
{
    if (!format)
        return;

    free(format);
    format = NULL;
}

/*
 * Check a user input border.
 */
STATIC uint8_t
_check_border(uint8_t border)
{
    if (border >= LXLSX_BORDER_THIN && border <= LXLSX_BORDER_SLANT_DASH_DOT)
        return border;
    else
        return LXLSX_BORDER_NONE;
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Returns a format struct suitable for hashing as a lookup key. This is
 * mainly a memcpy with any pointer members set to NULL.
 */
STATIC lxlsx_format *
_get_format_key(lxlsx_format *self)
{
    lxlsx_format *key = calloc(1, sizeof(lxlsx_format));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    memcpy(key, self, sizeof(lxlsx_format));

    /* Set pointer members to NULL since they aren't part of the comparison. */
    key->xf_format_indices = NULL;
    key->dxf_format_indices = NULL;
    key->num_xf_formats = NULL;
    key->num_dxf_formats = NULL;
    key->list_pointers.stqe_next = NULL;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a font struct suitable for hashing as a lookup key.
 */
lxlsx_font *
lxlsx_format_get_font_key(lxlsx_format *self)
{
    lxlsx_font *key = calloc(1, sizeof(lxlsx_font));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    LXLSX_FORMAT_FIELD_COPY(key->font_name, self->font_name);
    key->font_size = self->font_size;
    key->bold = self->bold;
    key->italic = self->italic;
    key->underline = self->underline;
    key->theme = self->theme;
    key->font_color = self->font_color;
    key->font_strikeout = self->font_strikeout;
    key->font_outline = self->font_outline;
    key->font_shadow = self->font_shadow;
    key->font_script = self->font_script;
    key->font_family = self->font_family;
    key->font_charset = self->font_charset;
    key->font_condense = self->font_condense;
    key->font_extend = self->font_extend;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a border struct suitable for hashing as a lookup key.
 */
lxlsx_border *
lxlsx_format_get_border_key(lxlsx_format *self)
{
    lxlsx_border *key = calloc(1, sizeof(lxlsx_border));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    key->bottom = self->bottom;
    key->left = self->left;
    key->right = self->right;
    key->top = self->top;
    key->diag_border = self->diag_border;
    key->diag_type = self->diag_type;
    key->bottom_color = self->bottom_color;
    key->left_color = self->left_color;
    key->right_color = self->right_color;
    key->top_color = self->top_color;
    key->diag_color = self->diag_color;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a pattern fill struct suitable for hashing as a lookup key.
 */
lxlsx_fill *
lxlsx_format_get_fill_key(lxlsx_format *self)
{
    lxlsx_fill *key = calloc(1, sizeof(lxlsx_fill));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    key->fg_color = self->fg_color;
    key->bg_color = self->bg_color;
    key->pattern = self->pattern;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns the XF index number used by Excel to identify a format.
 */
int32_t
lxlsx_format_get_xf_index(lxlsx_format *self)
{
    lxlsx_format *lxlsx_format_key;
    lxlsx_format *existing_format;
    lxlsx_hash_element *hash_element;
    lxlsx_hash_table *formats_hash_table = self->xf_format_indices;
    int32_t index;

    /* Note: The formats_hash_table/xf_format_indices contains the unique and
     * more importantly the *used* formats in the workbook.
     */

    /* Format already has an index number so return it. */
    if (self->xf_index != LXLSX_PROPERTY_UNSET) {
        return self->xf_index;
    }

    /* Otherwise, the format doesn't have an index number so we assign one.
     * First generate a unique key to identify the format in the hash table.
     */
    lxlsx_format_key = _get_format_key(self);

    /* Return the default format index if the key generation failed. */
    if (!lxlsx_format_key)
        return 0;

    /* Look up the format in the hash table. */
    hash_element =
        lxlsx_hash_key_exists(formats_hash_table, lxlsx_format_key,
                            sizeof(lxlsx_format));

    if (hash_element) {
        /* Format matches existing format with an index. */
        free(lxlsx_format_key);
        existing_format = hash_element->value;
        return existing_format->xf_index;
    }
    else {
        /* New format requiring an index. */
        index = formats_hash_table->unique_count;
        self->xf_index = index;
        lxlsx_insert_hash_element(formats_hash_table, lxlsx_format_key, self,
                                sizeof(lxlsx_format));
        return index;
    }
}

/*
 * Returns the DXF index number used by Excel to identify a format.
 */
int32_t
lxlsx_format_get_dxf_index(lxlsx_format *self)
{
    lxlsx_format *lxlsx_format_key;
    lxlsx_format *existing_format;
    lxlsx_hash_element *hash_element;
    lxlsx_hash_table *formats_hash_table = self->dxf_format_indices;
    int32_t index;

    /* Note: The formats_hash_table/dxf_format_indices contains the unique and
     * more importantly the *used* formats in the workbook.
     */

    /* Format already has an index number so return it. */
    if (self->dxf_index != LXLSX_PROPERTY_UNSET) {
        return self->dxf_index;
    }

    /* Otherwise, the format doesn't have an index number so we assign one.
     * First generate a unique key to identify the format in the hash table.
     */
    lxlsx_format_key = _get_format_key(self);

    /* Return the default format index if the key generation failed. */
    if (!lxlsx_format_key)
        return 0;

    /* Look up the format in the hash table. */
    hash_element =
        lxlsx_hash_key_exists(formats_hash_table, lxlsx_format_key,
                            sizeof(lxlsx_format));

    if (hash_element) {
        /* Format matches existing format with an index. */
        free(lxlsx_format_key);
        existing_format = hash_element->value;
        return existing_format->dxf_index;
    }
    else {
        /* New format requiring an index. */
        index = formats_hash_table->unique_count;
        self->dxf_index = index;
        lxlsx_insert_hash_element(formats_hash_table, lxlsx_format_key, self,
                                sizeof(lxlsx_format));
        return index;
    }
}

/*
 * Set the font_name property.
 */
void
lxlsx_format_set_font_name(lxlsx_format *self, const char *font_name)
{
    LXLSX_FORMAT_FIELD_COPY(self->font_name, font_name);
}

/*
 * Set the font_size property.
 */
void
lxlsx_format_set_font_size(lxlsx_format *self, double size)
{

    if (size >= LXLSX_MIN_FONT_SIZE && size <= LXLSX_MAX_FONT_SIZE)
        self->font_size = size;
}

/*
 * Set the font_color property.
 */
void
lxlsx_format_set_font_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->font_color = color;
}

/*
 * Set the bold property.
 */
void
lxlsx_format_set_bold(lxlsx_format *self)
{
    self->bold = LXLSX_TRUE;
}

/*
 * Set the italic property.
 */

void
lxlsx_format_set_italic(lxlsx_format *self)
{
    self->italic = LXLSX_TRUE;
}

/*
 * Set the underline property.
 */
void
lxlsx_format_set_underline(lxlsx_format *self, uint8_t style)
{
    if (style >= LXLSX_UNDERLINE_SINGLE
        && style <= LXLSX_UNDERLINE_DOUBLE_ACCOUNTING)
        self->underline = style;
}

/*
 * Set the font_strikeout property.
 */
void
lxlsx_format_set_font_strikeout(lxlsx_format *self)
{
    self->font_strikeout = LXLSX_TRUE;
}

/*
 * Set the font_script property.
 */
void
lxlsx_format_set_font_script(lxlsx_format *self, uint8_t style)
{
    if (style >= LXLSX_FONT_SUPERSCRIPT && style <= LXLSX_FONT_SUBSCRIPT)
        self->font_script = style;
}

/*
 * Set the font_outline property.
 */
void
lxlsx_format_set_font_outline(lxlsx_format *self)
{
    self->font_outline = LXLSX_TRUE;
}

/*
 * Set the font_shadow property.
 */
void
lxlsx_format_set_font_shadow(lxlsx_format *self)
{
    self->font_shadow = LXLSX_TRUE;
}

/*
 * Set the num_format property.
 */
void
lxlsx_format_set_num_format(lxlsx_format *self, const char *num_format)
{
    LXLSX_FORMAT_FIELD_COPY(self->num_format, num_format);
}

/*
 * Set the unlocked property.
 */
void
lxlsx_format_set_unlocked(lxlsx_format *self)
{
    self->locked = LXLSX_FALSE;
}

/*
 * Set the hidden property.
 */
void
lxlsx_format_set_hidden(lxlsx_format *self)
{
    self->hidden = LXLSX_TRUE;
}

/*
 * Set the align property.
 */
void
lxlsx_format_set_align(lxlsx_format *self, uint8_t value)
{
    if (value >= LXLSX_ALIGN_LEFT && value <= LXLSX_ALIGN_DISTRIBUTED) {
        self->text_h_align = value;
    }

    if (value >= LXLSX_ALIGN_VERTICAL_TOP
        && value <= LXLSX_ALIGN_VERTICAL_DISTRIBUTED) {
        self->text_v_align = value;
    }
}

/*
 * Set the text_wrap property.
 */
void
lxlsx_format_set_text_wrap(lxlsx_format *self)
{
    self->text_wrap = LXLSX_TRUE;
}

/*
 * Set the rotation property.
 */
void
lxlsx_format_set_rotation(lxlsx_format *self, int16_t angle)
{
    /* Convert user angle to Excel angle. */
    if (angle == 270) {
        self->rotation = 255;
    }
    else if (angle >= -90 && angle <= 90) {
        if (angle < 0)
            angle = -angle + 90;

        self->rotation = angle;
    }
    else {
        LXLSX_WARN("Rotation rotation outside range: -90 <= angle <= 90.");
        self->rotation = 0;
    }
}

/*
 * Set the indent property.
 */
void
lxlsx_format_set_indent(lxlsx_format *self, uint8_t value)
{
    self->indent = value;
}

/*
 * Set the shrink property.
 */
void
lxlsx_format_set_shrink(lxlsx_format *self)
{
    self->shrink = LXLSX_TRUE;
}

/*
 * Set the text_justlast property.
 */
void
lxlsx_format_set_text_justlast(lxlsx_format *self)
{
    self->text_justlast = LXLSX_TRUE;
}

/*
 * Set the pattern property.
 */
void
lxlsx_format_set_pattern(lxlsx_format *self, uint8_t value)
{
    if (value > LXLSX_PATTERN_GRAY_0625) {
        LXLSX_WARN_FORMAT1("lxlsx_format_set_pattern(): invalid pattern value: %d",
                         value);
        return;
    }

    self->pattern = value;
}

/*
 * Set the bg_color property.
 */
void
lxlsx_format_set_bg_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->bg_color = color;
}

/*
 * Set the fg_color property.
 */
void
lxlsx_format_set_fg_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->fg_color = color;
}

/*
 * Set the border property.
 */
void
lxlsx_format_set_border(lxlsx_format *self, uint8_t style)
{
    style = _check_border(style);
    self->bottom = style;
    self->top = style;
    self->left = style;
    self->right = style;
}

/*
 * Set the border_color property.
 */
void
lxlsx_format_set_border_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->bottom_color = color;
    self->top_color = color;
    self->left_color = color;
    self->right_color = color;
}

/*
 * Set the bottom property.
 */
void
lxlsx_format_set_bottom(lxlsx_format *self, uint8_t style)
{
    self->bottom = _check_border(style);
}

/*
 * Set the bottom_color property.
 */
void
lxlsx_format_set_bottom_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->bottom_color = color;
}

/*
 * Set the left property.
 */
void
lxlsx_format_set_left(lxlsx_format *self, uint8_t style)
{
    self->left = _check_border(style);
}

/*
 * Set the left_color property.
 */
void
lxlsx_format_set_left_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->left_color = color;
}

/*
 * Set the right property.
 */
void
lxlsx_format_set_right(lxlsx_format *self, uint8_t style)
{
    self->right = _check_border(style);
}

/*
 * Set the right_color property.
 */
void
lxlsx_format_set_right_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->right_color = color;
}

/*
 * Set the top property.
 */
void
lxlsx_format_set_top(lxlsx_format *self, uint8_t style)
{
    self->top = _check_border(style);
}

/*
 * Set the top_color property.
 */
void
lxlsx_format_set_top_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->top_color = color;
}

/*
 * Set the diag_type property.
 */
void
lxlsx_format_set_diag_type(lxlsx_format *self, uint8_t type)
{
    if (type >= LXLSX_DIAGONAL_BORDER_UP && type <= LXLSX_DIAGONAL_BORDER_UP_DOWN)
        self->diag_type = type;
}

/*
 * Set the diag_color property.
 */
void
lxlsx_format_set_diag_color(lxlsx_format *self, lxlsx_color_t color)
{
    self->diag_color = color;
}

/*
 * Set the diag_border property.
 */
void
lxlsx_format_set_diag_border(lxlsx_format *self, uint8_t style)
{
    if (style > LXLSX_BORDER_SLANT_DASH_DOT) {
        LXLSX_WARN_FORMAT1("lxlsx_format_set_diag_border(): invalid border style: %d",
                         style);
        return;
    }

    self->diag_border = style;
}

/*
 * Set the num_format_index property.
 */
void
lxlsx_format_set_num_format_index(lxlsx_format *self, uint8_t value)
{
    self->num_format_index = value;
}

/*
 * Set the valign property.
 */
void
lxlsx_format_set_valign(lxlsx_format *self, uint8_t value)
{
    if (value > LXLSX_ALIGN_VERTICAL_DISTRIBUTED) {
        LXLSX_WARN_FORMAT1
            ("lxlsx_format_set_valign(): invalid vertical alignment value: %d",
             value);
        return;
    }

    self->text_v_align = value;
}

/*
 * Set the reading_order property.
 */
void
lxlsx_format_set_reading_order(lxlsx_format *self, uint8_t value)
{
    self->reading_order = value;
}

/*
 * Set the font_family property.
 */
void
lxlsx_format_set_font_family(lxlsx_format *self, uint8_t value)
{
    self->font_family = value;
}

/*
 * Set the font_charset property.
 */
void
lxlsx_format_set_font_charset(lxlsx_format *self, uint8_t value)
{
    self->font_charset = value;
}

/*
 * Set the font_scheme property.
 */
void
lxlsx_format_set_font_scheme(lxlsx_format *self, const char *font_scheme)
{
    LXLSX_FORMAT_FIELD_COPY(self->font_scheme, font_scheme);
}

/*
 * Set the font_condense property.
 */
void
lxlsx_format_set_font_condense(lxlsx_format *self)
{
    self->font_condense = LXLSX_TRUE;
}

/*
 * Set the font_extend property.
 */
void
lxlsx_format_set_font_extend(lxlsx_format *self)
{
    self->font_extend = LXLSX_TRUE;
}

/*
 * Set the theme property.
 */
void
lxlsx_format_set_theme(lxlsx_format *self, uint8_t value)
{
    self->theme = value;
}

/*
 * Set the color_indexed property.
 */
void
lxlsx_format_set_color_indexed(lxlsx_format *self, uint8_t value)
{
    self->color_indexed = value;
}

/*
 * Set the font_only property.
 */
void
lxlsx_format_set_font_only(lxlsx_format *self)
{
    self->font_only = LXLSX_TRUE;
}

/*
 * Set the theme property.
 */
void
lxlsx_format_set_hyperlink(lxlsx_format *self)
{
    self->hyperlink = LXLSX_TRUE;
    self->xf_id = 1;
    self->underline = LXLSX_UNDERLINE_SINGLE;
    self->theme = 10;
}

/*
 * Set the quote_prefix property.
 */
void
lxlsx_format_set_quote_prefix(lxlsx_format *self)
{
    self->quote_prefix = LXLSX_TRUE;
}
