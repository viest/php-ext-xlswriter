/*****************************************************************************
 * format - A library for creating Excel XLSX format files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/format.h"
#include "xlsxwriter/utility.h"

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new format object.
 */
lxw_format *
lxw_format_new()
{
    lxw_format *format = calloc(1, sizeof(lxw_format));
    GOTO_LABEL_ON_MEM_ERROR(format, mem_error);

    format->xf_format_indices = NULL;

    format->xf_index = LXW_PROPERTY_UNSET;
    format->dxf_index = LXW_PROPERTY_UNSET;

    format->font_name[0] = '\0';
    format->font_scheme[0] = '\0';
    format->num_format[0] = '\0';
    format->num_format_index = 0;
    format->font_index = 0;
    format->has_font = LXW_FALSE;
    format->has_dxf_font = LXW_FALSE;
    format->font_size = 11.0;
    format->bold = LXW_FALSE;
    format->italic = LXW_FALSE;
    format->font_color = LXW_COLOR_UNSET;
    format->underline = LXW_FALSE;
    format->font_strikeout = LXW_FALSE;
    format->font_outline = LXW_FALSE;
    format->font_shadow = LXW_FALSE;
    format->font_script = LXW_FALSE;
    format->font_family = LXW_DEFAULT_FONT_FAMILY;
    format->font_charset = LXW_FALSE;
    format->font_condense = LXW_FALSE;
    format->font_extend = LXW_FALSE;
    format->theme = LXW_FALSE;
    format->hyperlink = LXW_FALSE;

    format->hidden = LXW_FALSE;
    format->locked = LXW_TRUE;

    format->text_h_align = LXW_ALIGN_NONE;
    format->text_wrap = LXW_FALSE;
    format->text_v_align = LXW_ALIGN_NONE;
    format->text_justlast = LXW_FALSE;
    format->rotation = 0;

    format->fg_color = LXW_COLOR_UNSET;
    format->bg_color = LXW_COLOR_UNSET;
    format->pattern = LXW_PATTERN_NONE;
    format->has_fill = LXW_FALSE;
    format->has_dxf_fill = LXW_FALSE;
    format->fill_index = 0;
    format->fill_count = 0;

    format->border_index = 0;
    format->has_border = LXW_FALSE;
    format->has_dxf_border = LXW_FALSE;
    format->border_count = 0;

    format->bottom = LXW_BORDER_NONE;
    format->left = LXW_BORDER_NONE;
    format->right = LXW_BORDER_NONE;
    format->top = LXW_BORDER_NONE;
    format->diag_border = LXW_BORDER_NONE;
    format->diag_type = LXW_BORDER_NONE;
    format->bottom_color = LXW_COLOR_UNSET;
    format->left_color = LXW_COLOR_UNSET;
    format->right_color = LXW_COLOR_UNSET;
    format->top_color = LXW_COLOR_UNSET;
    format->diag_color = LXW_COLOR_UNSET;

    format->indent = 0;
    format->shrink = LXW_FALSE;
    format->merge_range = LXW_FALSE;
    format->reading_order = 0;
    format->just_distrib = LXW_FALSE;
    format->color_indexed = LXW_FALSE;
    format->font_only = LXW_FALSE;

    return format;

mem_error:
    lxw_format_free(format);
    return NULL;
}

/*
 * Free a format object.
 */
void
lxw_format_free(lxw_format *format)
{
    if (!format)
        return;

    free(format);
    format = NULL;
}

/*
 * Check a user input color.
 */
lxw_color_t
lxw_format_check_color(lxw_color_t color)
{
    if (color == LXW_COLOR_UNSET)
        return color;
    else
        return color & LXW_COLOR_MASK;
}

/*
 * Check a user input border.
 */
STATIC uint8_t
_check_border(uint8_t border)
{
    if (border >= LXW_BORDER_THIN && border <= LXW_BORDER_SLANT_DASH_DOT)
        return border;
    else
        return LXW_BORDER_NONE;
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
STATIC lxw_format *
_get_format_key(lxw_format *self)
{
    lxw_format *key = calloc(1, sizeof(lxw_format));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    memcpy(key, self, sizeof(lxw_format));

    /* Set pointer members to NULL since they aren't part of the comparison. */
    key->xf_format_indices = NULL;
    key->num_xf_formats = NULL;
    key->list_pointers.stqe_next = NULL;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a font struct suitable for hashing as a lookup key.
 */
lxw_font *
lxw_format_get_font_key(lxw_format *self)
{
    lxw_font *key = calloc(1, sizeof(lxw_font));
    GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

    LXW_FORMAT_FIELD_COPY(key->font_name, self->font_name);
    key->font_size = self->font_size;
    key->bold = self->bold;
    key->italic = self->italic;
    key->font_color = self->font_color;
    key->underline = self->underline;
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
lxw_border *
lxw_format_get_border_key(lxw_format *self)
{
    lxw_border *key = calloc(1, sizeof(lxw_border));
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
lxw_fill *
lxw_format_get_fill_key(lxw_format *self)
{
    lxw_fill *key = calloc(1, sizeof(lxw_fill));
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
lxw_format_get_xf_index(lxw_format *self)
{
    lxw_format *format_key;
    lxw_format *existing_format;
    lxw_hash_element *hash_element;
    lxw_hash_table *formats_hash_table = self->xf_format_indices;
    int32_t index;

    /* Note: The formats_hash_table/xf_format_indices contains the unique and
     * more importantly the *used* formats in the workbook.
     */

    /* Format already has an index number so return it. */
    if (self->xf_index != LXW_PROPERTY_UNSET) {
        return self->xf_index;
    }

    /* Otherwise, the format doesn't have an index number so we assign one.
     * First generate a unique key to identify the format in the hash table.
     */
    format_key = _get_format_key(self);

    /* Return the default format index if the key generation failed. */
    if (!format_key)
        return 0;

    /* Look up the format in the hash table. */
    hash_element =
        lxw_hash_key_exists(formats_hash_table, format_key,
                            sizeof(lxw_format));

    if (hash_element) {
        /* Format matches existing format with an index. */
        free(format_key);
        existing_format = hash_element->value;
        return existing_format->xf_index;
    }
    else {
        /* New format requiring an index. */
        index = formats_hash_table->unique_count;
        self->xf_index = index;
        lxw_insert_hash_element(formats_hash_table, format_key, self,
                                sizeof(lxw_format));
        return index;
    }
}

/*
 * Set the font_name property.
 */
void
format_set_font_name(lxw_format *self, const char *font_name)
{
    LXW_FORMAT_FIELD_COPY(self->font_name, font_name);
}

/*
 * Set the font_size property.
 */
void
format_set_font_size(lxw_format *self, double size)
{

    if (size >= LXW_MIN_FONT_SIZE && size <= LXW_MAX_FONT_SIZE)
        self->font_size = size;
}

/*
 * Set the font_color property.
 */
void
format_set_font_color(lxw_format *self, lxw_color_t color)
{
    self->font_color = lxw_format_check_color(color);
}

/*
 * Set the bold property.
 */
void
format_set_bold(lxw_format *self)
{
    self->bold = LXW_TRUE;
}

/*
 * Set the italic property.
 */

void
format_set_italic(lxw_format *self)
{
    self->italic = LXW_TRUE;
}

/*
 * Set the underline property.
 */
void
format_set_underline(lxw_format *self, uint8_t style)
{
    if (style >= LXW_UNDERLINE_SINGLE
        && style <= LXW_UNDERLINE_DOUBLE_ACCOUNTING)
        self->underline = style;
}

/*
 * Set the font_strikeout property.
 */
void
format_set_font_strikeout(lxw_format *self)
{
    self->font_strikeout = LXW_TRUE;
}

/*
 * Set the font_script property.
 */
void
format_set_font_script(lxw_format *self, uint8_t style)
{
    if (style >= LXW_FONT_SUPERSCRIPT && style <= LXW_FONT_SUBSCRIPT)
        self->font_script = style;
}

/*
 * Set the font_outline property.
 */
void
format_set_font_outline(lxw_format *self)
{
    self->font_outline = LXW_TRUE;
}

/*
 * Set the font_shadow property.
 */
void
format_set_font_shadow(lxw_format *self)
{
    self->font_shadow = LXW_TRUE;
}

/*
 * Set the num_format property.
 */
void
format_set_num_format(lxw_format *self, const char *num_format)
{
    LXW_FORMAT_FIELD_COPY(self->num_format, num_format);
}

/*
 * Set the unlocked property.
 */
void
format_set_unlocked(lxw_format *self)
{
    self->locked = LXW_FALSE;
}

/*
 * Set the hidden property.
 */
void
format_set_hidden(lxw_format *self)
{
    self->hidden = LXW_TRUE;
}

/*
 * Set the align property.
 */
void
format_set_align(lxw_format *self, uint8_t value)
{
    if (value >= LXW_ALIGN_LEFT && value <= LXW_ALIGN_DISTRIBUTED) {
        self->text_h_align = value;
    }

    if (value >= LXW_ALIGN_VERTICAL_TOP
        && value <= LXW_ALIGN_VERTICAL_DISTRIBUTED) {
        self->text_v_align = value;
    }
}

/*
 * Set the text_wrap property.
 */
void
format_set_text_wrap(lxw_format *self)
{
    self->text_wrap = LXW_TRUE;
}

/*
 * Set the rotation property.
 */
void
format_set_rotation(lxw_format *self, int16_t angle)
{
    /* Convert user angle to Excel angle. */
    if (angle == 270) {
        self->rotation = 255;
    }
    else if (angle >= -90 || angle <= 90) {
        if (angle < 0)
            angle = -angle + 90;

        self->rotation = angle;
    }
    else {
        LXW_WARN("Rotation rotation outside range: -90 <= angle <= 90.");
        self->rotation = 0;
    }
}

/*
 * Set the indent property.
 */
void
format_set_indent(lxw_format *self, uint8_t value)
{
    self->indent = value;
}

/*
 * Set the shrink property.
 */
void
format_set_shrink(lxw_format *self)
{
    self->shrink = LXW_TRUE;
}

/*
 * Set the text_justlast property.
 */
void
format_set_text_justlast(lxw_format *self)
{
    self->text_justlast = LXW_TRUE;
}

/*
 * Set the pattern property.
 */
void
format_set_pattern(lxw_format *self, uint8_t value)
{
    self->pattern = value;
}

/*
 * Set the bg_color property.
 */
void
format_set_bg_color(lxw_format *self, lxw_color_t color)
{
    self->bg_color = lxw_format_check_color(color);
}

/*
 * Set the fg_color property.
 */
void
format_set_fg_color(lxw_format *self, lxw_color_t color)
{
    self->fg_color = lxw_format_check_color(color);
}

/*
 * Set the border property.
 */
void
format_set_border(lxw_format *self, uint8_t style)
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
format_set_border_color(lxw_format *self, lxw_color_t color)
{
    color = lxw_format_check_color(color);
    self->bottom_color = color;
    self->top_color = color;
    self->left_color = color;
    self->right_color = color;
}

/*
 * Set the bottom property.
 */
void
format_set_bottom(lxw_format *self, uint8_t style)
{
    self->bottom = _check_border(style);
}

/*
 * Set the bottom_color property.
 */
void
format_set_bottom_color(lxw_format *self, lxw_color_t color)
{
    self->bottom_color = lxw_format_check_color(color);
}

/*
 * Set the left property.
 */
void
format_set_left(lxw_format *self, uint8_t style)
{
    self->left = _check_border(style);
}

/*
 * Set the left_color property.
 */
void
format_set_left_color(lxw_format *self, lxw_color_t color)
{
    self->left_color = lxw_format_check_color(color);
}

/*
 * Set the right property.
 */
void
format_set_right(lxw_format *self, uint8_t style)
{
    self->right = _check_border(style);
}

/*
 * Set the right_color property.
 */
void
format_set_right_color(lxw_format *self, lxw_color_t color)
{
    self->right_color = lxw_format_check_color(color);
}

/*
 * Set the top property.
 */
void
format_set_top(lxw_format *self, uint8_t style)
{
    self->top = _check_border(style);
}

/*
 * Set the top_color property.
 */
void
format_set_top_color(lxw_format *self, lxw_color_t color)
{
    self->top_color = lxw_format_check_color(color);
}

/*
 * Set the diag_type property.
 */
void
format_set_diag_type(lxw_format *self, uint8_t type)
{
    if (type >= LXW_DIAGONAL_BORDER_UP && type <= LXW_DIAGONAL_BORDER_UP_DOWN)
        self->diag_type = type;
}

/*
 * Set the diag_color property.
 */
void
format_set_diag_color(lxw_format *self, lxw_color_t color)
{
    self->diag_color = lxw_format_check_color(color);
}

/*
 * Set the diag_border property.
 */
void
format_set_diag_border(lxw_format *self, uint8_t style)
{
    self->diag_border = style;
}

/*
 * Set the num_format_index property.
 */
void
format_set_num_format_index(lxw_format *self, uint8_t value)
{
    self->num_format_index = value;
}

/*
 * Set the valign property.
 */
void
format_set_valign(lxw_format *self, uint8_t value)
{
    self->text_v_align = value;
}

/*
 * Set the reading_order property.
 */
void
format_set_reading_order(lxw_format *self, uint8_t value)
{
    self->reading_order = value;
}

/*
 * Set the font_family property.
 */
void
format_set_font_family(lxw_format *self, uint8_t value)
{
    self->font_family = value;
}

/*
 * Set the font_charset property.
 */
void
format_set_font_charset(lxw_format *self, uint8_t value)
{
    self->font_charset = value;
}

/*
 * Set the font_scheme property.
 */
void
format_set_font_scheme(lxw_format *self, const char *font_scheme)
{
    LXW_FORMAT_FIELD_COPY(self->font_scheme, font_scheme);
}

/*
 * Set the font_condense property.
 */
void
format_set_font_condense(lxw_format *self)
{
    self->font_condense = LXW_TRUE;
}

/*
 * Set the font_extend property.
 */
void
format_set_font_extend(lxw_format *self)
{
    self->font_extend = LXW_TRUE;
}

/*
 * Set the theme property.
 */
void
format_set_theme(lxw_format *self, uint8_t value)
{
    self->theme = value;
}
