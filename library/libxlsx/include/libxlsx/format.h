/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 */

/**
 * @page lxlsx_format_page The Format object
 *
 * The Format object represents an the formatting properties that can be
 * applied to a cell including: fonts, colors, patterns,
 * borders, alignment and number formatting.
 *
 * See @ref format.h for full details of the functionality.
 *
 * @file format.h
 *
 * @brief Functions and properties for adding formatting to cells in Excel.
 *
 * This section describes the functions and properties that are available for
 * formatting cells in Excel.
 *
 * The properties of a cell that can be formatted include: fonts, colors,
 * patterns, borders, alignment and number formatting.
 *
 * @image html formats_intro.png
 *
 * Formats in `libxlsxwriter` are accessed via the lxlsx_format
 * struct. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the lxlsx_workbook_add_format() method as
 * follows:
 *
 * @code
 *     lxlsx_format *format = lxlsx_workbook_add_format(workbook);
 * @endcode
 *
 * The members of the lxlsx_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in this section.
 * For example:
 *
 * @code
 *    // Create the Format.
 *    lxlsx_format *format = lxlsx_workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    lxlsx_format_set_bold(format);
 *    lxlsx_format_set_font_color(format, LXLSX_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 * The full range of formatting options that can be applied using
 * `libxlsxwriter` are shown below.
 *
 */
#ifndef __LXLSX_FORMAT_H__
#define __LXLSX_FORMAT_H__

#include <stdint.h>
#include <string.h>
#include "hash_table.h"

#include "common.h"

/**
 * @brief The type for RGB colors in libxlsxwriter.
 *
 * The type for RGB colors in libxlsxwriter. The valid range is `0x000000`
 * (black) to `0xFFFFFF` (white). See @ref working_with_colors.
 */
typedef uint32_t lxlsx_color_t;

#define LXLSX_FORMAT_FIELD_LEN            128
#define LXLSX_DEFAULT_FONT_NAME           "Calibri"
#define LXLSX_DEFAULT_FONT_FAMILY         2
#define LXLSX_DEFAULT_FONT_THEME          1
#define LXLSX_PROPERTY_UNSET              -1
#define LXLSX_COLOR_UNSET                 0x000000
#define LXLSX_COLOR_MASK                  0xFFFFFF
#define LXLSX_MIN_FONT_SIZE               1.0
#define LXLSX_MAX_FONT_SIZE               409.0

#define LXLSX_FORMAT_FIELD_COPY(dst, src)             \
    do{                                             \
        strncpy(dst, src, LXLSX_FORMAT_FIELD_LEN -1); \
        dst[LXLSX_FORMAT_FIELD_LEN - 1] = '\0';       \
    } while (0)

/** Format underline values for lxlsx_format_set_underline(). */
enum lxlsx_format_underlines {
    LXLSX_UNDERLINE_NONE = 0,

    /** Single underline */
    LXLSX_UNDERLINE_SINGLE,

    /** Double underline */
    LXLSX_UNDERLINE_DOUBLE,

    /** Single accounting underline */
    LXLSX_UNDERLINE_SINGLE_ACCOUNTING,

    /** Double accounting underline */
    LXLSX_UNDERLINE_DOUBLE_ACCOUNTING
};

/** Superscript and subscript values for lxlsx_format_set_font_script(). */
enum lxlsx_format_scripts {

    /** Superscript font */
    LXLSX_FONT_SUPERSCRIPT = 1,

    /** Subscript font */
    LXLSX_FONT_SUBSCRIPT
};

/** Alignment values for lxlsx_format_set_align(). */
enum lxlsx_format_alignments {
    /** No alignment. Cell will use Excel's default for the data type */
    LXLSX_ALIGN_NONE = 0,

    /** Left horizontal alignment */
    LXLSX_ALIGN_LEFT,

    /** Center horizontal alignment */
    LXLSX_ALIGN_CENTER,

    /** Right horizontal alignment */
    LXLSX_ALIGN_RIGHT,

    /** Cell fill horizontal alignment */
    LXLSX_ALIGN_FILL,

    /** Justify horizontal alignment */
    LXLSX_ALIGN_JUSTIFY,

    /** Center Across horizontal alignment */
    LXLSX_ALIGN_CENTER_ACROSS,

    /** Left horizontal alignment */
    LXLSX_ALIGN_DISTRIBUTED,

    /** Top vertical alignment */
    LXLSX_ALIGN_VERTICAL_TOP,

    /** Bottom vertical alignment */
    LXLSX_ALIGN_VERTICAL_BOTTOM,

    /** Center vertical alignment */
    LXLSX_ALIGN_VERTICAL_CENTER,

    /** Justify vertical alignment */
    LXLSX_ALIGN_VERTICAL_JUSTIFY,

    /** Distributed vertical alignment */
    LXLSX_ALIGN_VERTICAL_DISTRIBUTED
};

/**
 * Diagonal border types.
 *
 */
enum lxlsx_format_diagonal_types {

    /** Cell diagonal border from bottom left to top right. */
    LXLSX_DIAGONAL_BORDER_UP = 1,

    /** Cell diagonal border from top left to bottom right. */
    LXLSX_DIAGONAL_BORDER_DOWN,

    /** Cell diagonal border in both directions. */
    LXLSX_DIAGONAL_BORDER_UP_DOWN
};

/** Predefined values for common colors. */
enum lxlsx_defined_colors {
    /** Black */
    LXLSX_COLOR_BLACK = 0x1000000,

    /** Blue */
    LXLSX_COLOR_BLUE = 0x0000FF,

    /** Brown */
    LXLSX_COLOR_BROWN = 0x800000,

    /** Cyan */
    LXLSX_COLOR_CYAN = 0x00FFFF,

    /** Gray */
    LXLSX_COLOR_GRAY = 0x808080,

    /** Green */
    LXLSX_COLOR_GREEN = 0x008000,

    /** Lime */
    LXLSX_COLOR_LIME = 0x00FF00,

    /** Magenta */
    LXLSX_COLOR_MAGENTA = 0xFF00FF,

    /** Navy */
    LXLSX_COLOR_NAVY = 0x000080,

    /** Orange */
    LXLSX_COLOR_ORANGE = 0xFF6600,

    /** Pink */
    LXLSX_COLOR_PINK = 0xFF00FF,

    /** Purple */
    LXLSX_COLOR_PURPLE = 0x800080,

    /** Red */
    LXLSX_COLOR_RED = 0xFF0000,

    /** Silver */
    LXLSX_COLOR_SILVER = 0xC0C0C0,

    /** White */
    LXLSX_COLOR_WHITE = 0xFFFFFF,

    /** Yellow */
    LXLSX_COLOR_YELLOW = 0xFFFF00
};

/** Pattern value for use with lxlsx_format_set_pattern(). */
enum lxlsx_format_patterns {
    /** Empty pattern */
    LXLSX_PATTERN_NONE = 0,

    /** Solid pattern */
    LXLSX_PATTERN_SOLID,

    /** Medium gray pattern */
    LXLSX_PATTERN_MEDIUM_GRAY,

    /** Dark gray pattern */
    LXLSX_PATTERN_DARK_GRAY,

    /** Light gray pattern */
    LXLSX_PATTERN_LIGHT_GRAY,

    /** Dark horizontal line pattern */
    LXLSX_PATTERN_DARK_HORIZONTAL,

    /** Dark vertical line pattern */
    LXLSX_PATTERN_DARK_VERTICAL,

    /** Dark diagonal stripe pattern */
    LXLSX_PATTERN_DARK_DOWN,

    /** Reverse dark diagonal stripe pattern */
    LXLSX_PATTERN_DARK_UP,

    /** Dark grid pattern */
    LXLSX_PATTERN_DARK_GRID,

    /** Dark trellis pattern */
    LXLSX_PATTERN_DARK_TRELLIS,

    /** Light horizontal Line pattern */
    LXLSX_PATTERN_LIGHT_HORIZONTAL,

    /** Light vertical line pattern */
    LXLSX_PATTERN_LIGHT_VERTICAL,

    /** Light diagonal stripe pattern */
    LXLSX_PATTERN_LIGHT_DOWN,

    /** Reverse light diagonal stripe pattern */
    LXLSX_PATTERN_LIGHT_UP,

    /** Light grid pattern */
    LXLSX_PATTERN_LIGHT_GRID,

    /** Light trellis pattern */
    LXLSX_PATTERN_LIGHT_TRELLIS,

    /** 12.5% gray pattern */
    LXLSX_PATTERN_GRAY_125,

    /** 6.25% gray pattern */
    LXLSX_PATTERN_GRAY_0625
};

/** Cell border styles for use with lxlsx_format_set_border(). */
enum lxlsx_format_borders {
    /** No border */
    LXLSX_BORDER_NONE,

    /** Thin border style */
    LXLSX_BORDER_THIN,

    /** Medium border style */
    LXLSX_BORDER_MEDIUM,

    /** Dashed border style */
    LXLSX_BORDER_DASHED,

    /** Dotted border style */
    LXLSX_BORDER_DOTTED,

    /** Thick border style */
    LXLSX_BORDER_THICK,

    /** Double border style */
    LXLSX_BORDER_DOUBLE,

    /** Hair border style */
    LXLSX_BORDER_HAIR,

    /** Medium dashed border style */
    LXLSX_BORDER_MEDIUM_DASHED,

    /** Dash-dot border style */
    LXLSX_BORDER_DASH_DOT,

    /** Medium dash-dot border style */
    LXLSX_BORDER_MEDIUM_DASH_DOT,

    /** Dash-dot-dot border style */
    LXLSX_BORDER_DASH_DOT_DOT,

    /** Medium dash-dot-dot border style */
    LXLSX_BORDER_MEDIUM_DASH_DOT_DOT,

    /** Slant dash-dot border style */
    LXLSX_BORDER_SLANT_DASH_DOT
};

/**
 * @brief Struct to represent the formatting properties of an Excel format.
 *
 * Formats in `libxlsxwriter` are accessed via this struct.
 *
 * The members of the lxlsx_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in format.h.
 *
 * For example:
 *
 * @code
 *    // Create the Format.
 *    lxlsx_format *format = lxlsx_workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    lxlsx_format_set_bold(format);
 *    lxlsx_format_set_font_color(format, LXLSX_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 */
typedef struct lxlsx_format {

    FILE *file;

    lxlsx_hash_table *xf_format_indices;
    lxlsx_hash_table *dxf_format_indices;
    uint16_t *num_xf_formats;
    uint16_t *num_dxf_formats;

    int32_t xf_index;
    int32_t dxf_index;
    int32_t xf_id;

    char num_format[LXLSX_FORMAT_FIELD_LEN];
    char font_name[LXLSX_FORMAT_FIELD_LEN];
    char font_scheme[LXLSX_FORMAT_FIELD_LEN];
    uint16_t num_format_index;
    uint16_t font_index;
    uint8_t has_font;
    uint8_t has_dxf_font;
    double font_size;
    uint8_t bold;
    uint8_t italic;
    lxlsx_color_t font_color;
    uint8_t underline;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    uint8_t theme;
    uint8_t hyperlink;

    uint8_t hidden;
    uint8_t locked;

    uint8_t text_h_align;
    uint8_t text_wrap;
    uint8_t text_v_align;
    uint8_t text_justlast;
    int16_t rotation;

    lxlsx_color_t fg_color;
    lxlsx_color_t bg_color;
    lxlsx_color_t dxf_fg_color;
    lxlsx_color_t dxf_bg_color;
    uint8_t pattern;
    uint8_t has_fill;
    uint8_t has_dxf_fill;
    int32_t fill_index;
    int32_t fill_count;

    int32_t border_index;
    uint8_t has_border;
    uint8_t has_dxf_border;
    int32_t border_count;

    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;
    lxlsx_color_t bottom_color;
    lxlsx_color_t diag_color;
    lxlsx_color_t left_color;
    lxlsx_color_t right_color;
    lxlsx_color_t top_color;

    uint8_t indent;
    uint8_t shrink;
    uint8_t merge_range;
    uint8_t reading_order;
    uint8_t just_distrib;
    uint8_t color_indexed;
    uint8_t font_only;

    uint8_t quote_prefix;

    STAILQ_ENTRY (lxlsx_format) list_pointers;
} lxlsx_format;

/*
 * Struct to represent the font component of a format.
 */
typedef struct lxlsx_font {

    char font_name[LXLSX_FORMAT_FIELD_LEN];
    double font_size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    uint8_t theme;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    lxlsx_color_t font_color;
} lxlsx_font;

/*
 * Struct to represent the border component of a format.
 */
typedef struct lxlsx_border {

    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;

    lxlsx_color_t bottom_color;
    lxlsx_color_t diag_color;
    lxlsx_color_t left_color;
    lxlsx_color_t right_color;
    lxlsx_color_t top_color;

} lxlsx_border;

/*
 * Struct to represent the fill component of a format.
 */
typedef struct lxlsx_fill {

    lxlsx_color_t fg_color;
    lxlsx_color_t bg_color;
    uint8_t pattern;

} lxlsx_fill;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_format *lxlsx_format_new(void);
void lxlsx_format_free(lxlsx_format *format);
int32_t lxlsx_format_get_xf_index(lxlsx_format *format);
int32_t lxlsx_format_get_dxf_index(lxlsx_format *format);
lxlsx_font *lxlsx_format_get_font_key(lxlsx_format *format);
lxlsx_border *lxlsx_format_get_border_key(lxlsx_format *format);
lxlsx_fill *lxlsx_format_get_fill_key(lxlsx_format *format);

/**
 * @brief Set the font used in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param font_name Cell font name.
 *
 * Specify the font used used in the cell format:
 *
 * @code
 *     lxlsx_format_set_font_name(format, "Avenir Black Oblique");
 * @endcode
 *
 * @image html lxlsx_format_set_font_name.png
 *
 * Excel can only display fonts that are installed on the system that it is
 * running on. Therefore it is generally best to use the fonts that come as
 * standard with Excel such as Calibri, Times New Roman and Courier New.
 *
 * The default font in Excel 2007, and later, is Calibri.
 */
void lxlsx_format_set_font_name(lxlsx_format *format, const char *font_name);

/**
 * @brief Set the size of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param size   The cell font size.
 *
 * Set the font size of the cell format:
 *
 * @code
 *     lxlsx_format_set_font_size(format, 30);
 * @endcode
 *
 * @image html lxlsx_format_font_size.png
 *
 * Excel adjusts the height of a row to accommodate the largest font
 * size in the row. You can also explicitly specify the height of a
 * row using the lxlsx_worksheet_set_row() function.
 */
void lxlsx_format_set_font_size(lxlsx_format *format, double size);

/**
 * @brief Set the color of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell font color.
 *
 *
 * Set the font color:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_font_color(format, LXLSX_COLOR_RED);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Wheelbarrow", format);
 * @endcode
 *
 * @image html lxlsx_format_font_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * @note
 * The lxlsx_format_set_font_color() method is used to set the font color in a
 * cell. To set the color of a cell background use the lxlsx_format_set_bg_color()
 * and lxlsx_format_set_pattern() methods.
 */
void lxlsx_format_set_font_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Turn on bold for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the bold property of the font:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_bold(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Bold Text", format);
 * @endcode
 *
 * @image html lxlsx_format_font_bold.png
 */
void lxlsx_format_set_bold(lxlsx_format *format);

/**
 * @brief Turn on italic for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the italic property of the font:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_italic(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Italic Text", format);
 * @endcode
 *
 * @image html lxlsx_format_font_italic.png
 */
void lxlsx_format_set_italic(lxlsx_format *format);

/**
 * @brief Turn on underline for the format:
 *
 * @param format Pointer to a Format instance.
 * @param style Underline style.
 *
 * Set the underline property of the format:
 *
 * @code
 *     lxlsx_format_set_underline(format, LXLSX_UNDERLINE_SINGLE);
 * @endcode
 *
 * @image html lxlsx_format_font_underlined.png
 *
 * The available underline styles are:
 *
 * - #LXLSX_UNDERLINE_SINGLE
 * - #LXLSX_UNDERLINE_DOUBLE
 * - #LXLSX_UNDERLINE_SINGLE_ACCOUNTING
 * - #LXLSX_UNDERLINE_DOUBLE_ACCOUNTING
 *
 */
void lxlsx_format_set_underline(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the strikeout property of the font.
 *
 * @param format Pointer to a Format instance.
 *
 * @image html lxlsx_format_font_strikeout.png
 *
 */
void lxlsx_format_set_font_strikeout(lxlsx_format *format);

/**
 * @brief Set the superscript/subscript property of the font.
 *
 * @param format Pointer to a Format instance.
 * @param style  Superscript or subscript style.
 *
 * Set the superscript o subscript property of the font.
 *
 * @image html lxlsx_format_font_script.png
 *
 * The available script styles are:
 *
 * - #LXLSX_FONT_SUPERSCRIPT
 * - #LXLSX_FONT_SUBSCRIPT
 */
void lxlsx_format_set_font_script(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the Format font family property.
 *
 * @param format Pointer to a Format instance.
 * @param value  The font family index.
 *
 * Set the font family. This is usually an integer in the range 1-4. This
 * function is implemented for completeness but is rarely used in practice.
 *
 * @code
 *     lxlsx_format_set_font_family(format, 178);
 * @endcode
 *
 */
void lxlsx_format_set_font_family(lxlsx_format *format, uint8_t value);

/**
 * @brief Set the Format font character set property.
 *
 * @param format Pointer to a Format instance.
 * @param value  The font character set.
 *
 * Set the font character set property. This function is implemented for
 * completeness but is rarely used in practice.
 *
 * @code
 *     lxlsx_format_set_font_charset(format, 178);
 * @endcode
 *
 */
void lxlsx_format_set_font_charset(lxlsx_format *format, uint8_t value);

/**
 * @brief Set the number format for a cell.
 *
 * @param format      Pointer to a Format instance.
 * @param num_format The cell number format string.
 *
 * This method is used to define the numerical format of a number in
 * Excel. It controls whether a number is displayed as an integer, a
 * floating point number, a date, a currency value or some other user
 * defined format.
 *
 * The numerical format of a cell can be specified by using a format
 * string:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_num_format(format, "d mmm yyyy");
 * @endcode
 *
 * Format strings can control any aspect of number formatting allowed by Excel:
 *
 * @dontinclude lxlsx_format_num_format.c
 * @skipline set_num_format
 * @until 1209
 *
 * @image html lxlsx_format_set_num_format.png
 *
 * To set a number format that matches an Excel format category such as "Date"
 * or "Currency" see @ref ww_formats_categories.
 *
 * The number system used for dates is described in @ref working_with_dates.
 *
 * For more information on number formats in Excel refer to the
 * [Microsoft documentation on cell formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
 */
void lxlsx_format_set_num_format(lxlsx_format *format, const char *num_format);

/**
 * @brief Set the Excel built-in number format for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param index  The built-in number format index for the cell.
 *
 * This function is similar to lxlsx_format_set_num_format() except that it takes an
 * index to a limited number of Excel's built-in number formats instead of a
 * user defined format string:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_num_format_index(format, 0x0F); // d-mmm-yy
 * @endcode
 *
 * @note
 * Unless you need to specifically access one of Excel's built-in number
 * formats the lxlsx_format_set_num_format() function above is a better
 * solution. The lxlsx_format_set_num_format_index() function is mainly included for
 * backward compatibility and completeness.
 *
 * The Excel built-in number formats as shown in the table below:
 *
 *   | Index | Index | Format String                                        |
 *   | ----- | ----- | ---------------------------------------------------- |
 *   | 0     | 0x00  | `General`                                            |
 *   | 1     | 0x01  | `0`                                                  |
 *   | 2     | 0x02  | `0.00`                                               |
 *   | 3     | 0x03  | `#,##0`                                              |
 *   | 4     | 0x04  | `#,##0.00`                                           |
 *   | 5     | 0x05  | `($#,##0_);($#,##0)`                                 |
 *   | 6     | 0x06  | `($#,##0_);[Red]($#,##0)`                            |
 *   | 7     | 0x07  | `($#,##0.00_);($#,##0.00)`                           |
 *   | 8     | 0x08  | `($#,##0.00_);[Red]($#,##0.00)`                      |
 *   | 9     | 0x09  | `0%`                                                 |
 *   | 10    | 0x0a  | `0.00%`                                              |
 *   | 11    | 0x0b  | `0.00E+00`                                           |
 *   | 12    | 0x0c  | `# ?/?`                                              |
 *   | 13    | 0x0d  | `# ??/??`                                            |
 *   | 14    | 0x0e  | `m/d/yy`                                             |
 *   | 15    | 0x0f  | `d-mmm-yy`                                           |
 *   | 16    | 0x10  | `d-mmm`                                              |
 *   | 17    | 0x11  | `mmm-yy`                                             |
 *   | 18    | 0x12  | `h:mm AM/PM`                                         |
 *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
 *   | 20    | 0x14  | `h:mm`                                               |
 *   | 21    | 0x15  | `h:mm:ss`                                            |
 *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
 *   | ...   | ...   | ...                                                  |
 *   | 37    | 0x25  | `(#,##0_);(#,##0)`                                   |
 *   | 38    | 0x26  | `(#,##0_);[Red](#,##0)`                              |
 *   | 39    | 0x27  | `(#,##0.00_);(#,##0.00)`                             |
 *   | 40    | 0x28  | `(#,##0.00_);[Red](#,##0.00)`                        |
 *   | 41    | 0x29  | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
 *   | 42    | 0x2a  | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
 *   | 43    | 0x2b  | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
 *   | 44    | 0x2c  | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
 *   | 45    | 0x2d  | `mm:ss`                                              |
 *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
 *   | 47    | 0x2f  | `mm:ss.0`                                            |
 *   | 48    | 0x30  | `##0.0E+0`                                           |
 *   | 49    | 0x31  | `@`                                                  |
 *
 * @note
 *  -  Numeric formats 23 to 36 are not documented by Microsoft and may differ
 *     in international versions. The listed date and currency formats may also
 *     vary depending on system settings.
 *  - The dollar sign in the above format appears as the defined local currency
 *    symbol.
 *  - These formats can also be set via lxlsx_format_set_num_format().
 *  - See also @ref ww_formats_categories.
 */
void lxlsx_format_set_num_format_index(lxlsx_format *format, uint8_t index);

/**
 * @brief Set the cell unlocked state.
 *
 * @param format Pointer to a Format instance.
 *
 * This property can be used to allow modification of a cell in a protected
 * worksheet. In Excel, cell locking is turned on by default for all
 * cells. However, it only has an effect if the worksheet has been protected
 * using the worksheet lxlsx_worksheet_protect() function:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_unlocked(format);
 *
 *     // Enable worksheet protection, without password or options.
 *     lxlsx_worksheet_protect(worksheet, NULL, NULL);
 *
 *     // This cell cannot be edited.
 *     lxlsx_worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL);
 *
 *     // This cell can be edited.
 *     lxlsx_worksheet_write_formula(worksheet, 1, 0, "=1+2", format);
 * @endcode
 */
void lxlsx_format_set_unlocked(lxlsx_format *format);

/**
 * @brief Hide formulas in a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This property is used to hide a formula while still displaying its
 * result. This is generally used to hide complex calculations from end users
 * who are only interested in the result. It only has an effect if the
 * worksheet has been protected using the worksheet lxlsx_worksheet_protect()
 * function:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_hidden(format);
 *
 *     // Enable worksheet protection, without password or options.
 *     lxlsx_worksheet_protect(worksheet, NULL, NULL);
 *
 *     // The formula in this cell isn't visible.
 *     lxlsx_worksheet_write_formula(worksheet, 0, 0, "=1+2", format);
 * @endcode
 */
void lxlsx_format_set_hidden(lxlsx_format *format);

/**
 * @brief Set the alignment for data in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param alignment The horizontal and or vertical alignment direction.
 *
 * This method is used to set the horizontal and vertical text alignment within a
 * cell. The following are the available horizontal alignments:
 *
 * - #LXLSX_ALIGN_LEFT
 * - #LXLSX_ALIGN_CENTER
 * - #LXLSX_ALIGN_RIGHT
 * - #LXLSX_ALIGN_FILL
 * - #LXLSX_ALIGN_JUSTIFY
 * - #LXLSX_ALIGN_CENTER_ACROSS
 * - #LXLSX_ALIGN_DISTRIBUTED
 *
 * The following are the available vertical alignments:
 *
 * - #LXLSX_ALIGN_VERTICAL_TOP
 * - #LXLSX_ALIGN_VERTICAL_BOTTOM
 * - #LXLSX_ALIGN_VERTICAL_CENTER
 * - #LXLSX_ALIGN_VERTICAL_JUSTIFY
 * - #LXLSX_ALIGN_VERTICAL_DISTRIBUTED
 *
 * As in Excel, vertical and horizontal alignments can be combined:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *
 *     lxlsx_format_set_align(format, LXLSX_ALIGN_CENTER);
 *     lxlsx_format_set_align(format, LXLSX_ALIGN_VERTICAL_CENTER);
 *
 *     lxlsx_worksheet_set_row(0, 30);
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Some Text", format);
 * @endcode
 *
 * @image html lxlsx_format_font_align.png
 *
 * Text can be aligned across two or more adjacent cells using the
 * center_across property. However, for genuine merged cells it is better to
 * use the lxlsx_worksheet_merge_range() worksheet method.
 *
 * The vertical justify option can be used to provide automatic text wrapping
 * in a cell. The height of the cell will be adjusted to accommodate the
 * wrapped text. To specify where the text wraps use the
 * lxlsx_format_set_text_wrap() method.
 */
void lxlsx_format_set_align(lxlsx_format *format, uint8_t alignment);

/**
 * @brief Wrap text in a cell.
 *
 * Turn text wrapping on for text in a cell.
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_text_wrap(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a cell", format);
 * @endcode
 *
 * If you wish to control where the text is wrapped you can add newline characters
 * to the string:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_text_wrap(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format);
 * @endcode
 *
 * @image html lxlsx_format_font_text_wrap.png
 *
 * Excel will adjust the height of the row to accommodate the wrapped text. A
 * similar effect can be obtained without newlines using the
 * lxlsx_format_set_align() function with #LXLSX_ALIGN_VERTICAL_JUSTIFY.
 */
void lxlsx_format_set_text_wrap(lxlsx_format *format);

/**
 * @brief Set the rotation of the text in a cell.
 *
 * @param format Pointer to a Format instance.
 * @param angle  Rotation angle in the range -90 to 90 and 270.
 *
 * Set the rotation of the text in a cell. The rotation can be any angle in the
 * range -90 to 90 degrees:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_rotation(format, 30);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "This text is rotated", format);
 * @endcode
 *
 * @image html lxlsx_format_font_text_rotated.png
 *
 * The angle 270 is also supported. This indicates text where the letters run from
 * top to bottom.
 */
void lxlsx_format_set_rotation(lxlsx_format *format, int16_t angle);

/**
 * @brief Set the cell text indentation level.
 *
 * @param format Pointer to a Format instance.
 * @param level  Indentation level.
 *
 * This method can be used to indent text in a cell. The argument, which should be
 * an integer, is taken as the level of indentation:
 *
 * @code
 *     format1 = lxlsx_workbook_add_format(workbook);
 *     format2 = lxlsx_workbook_add_format(workbook);
 *
 *     lxlsx_format_set_indent(format1, 1);
 *     lxlsx_format_set_indent(format2, 2);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",  format1);
 *     lxlsx_worksheet_write_string(worksheet, 1, 0, "This text is indented 2 levels", format2);
 * @endcode
 *
 * @image html text_indent.png
 *
 * @note
 * Indentation is a horizontal alignment property. It will override any other
 * horizontal properties but it can be used in conjunction with vertical
 * properties.
 */
void lxlsx_format_set_indent(lxlsx_format *format, uint8_t level);

/**
 * @brief Turn on the text "shrink to fit" for a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This method can be used to shrink text so that it fits in a cell:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_shrink(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!", format);
 * @endcode
 */
void lxlsx_format_set_shrink(lxlsx_format *format);

/**
 * @brief Set the background fill pattern for a cell
 *
 * @param format Pointer to a Format instance.
 * @param index  Pattern index.
 *
 * Set the background pattern for a cell.
 *
 * The most common pattern is a solid fill of the background color:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *
 *     lxlsx_format_set_pattern (format, LXLSX_PATTERN_SOLID);
 *     lxlsx_format_set_bg_color(format, LXLSX_COLOR_YELLOW);
 * @endcode
 *
 * The available fill patterns are:
 *
 *    Fill Type                     | Define
 *    ----------------------------- | -----------------------------
 *    Solid                         | #LXLSX_PATTERN_SOLID
 *    Medium gray                   | #LXLSX_PATTERN_MEDIUM_GRAY
 *    Dark gray                     | #LXLSX_PATTERN_DARK_GRAY
 *    Light gray                    | #LXLSX_PATTERN_LIGHT_GRAY
 *    Dark horizontal line          | #LXLSX_PATTERN_DARK_HORIZONTAL
 *    Dark vertical line            | #LXLSX_PATTERN_DARK_VERTICAL
 *    Dark diagonal stripe          | #LXLSX_PATTERN_DARK_DOWN
 *    Reverse dark diagonal stripe  | #LXLSX_PATTERN_DARK_UP
 *    Dark grid                     | #LXLSX_PATTERN_DARK_GRID
 *    Dark trellis                  | #LXLSX_PATTERN_DARK_TRELLIS
 *    Light horizontal line         | #LXLSX_PATTERN_LIGHT_HORIZONTAL
 *    Light vertical line           | #LXLSX_PATTERN_LIGHT_VERTICAL
 *    Light diagonal stripe         | #LXLSX_PATTERN_LIGHT_DOWN
 *    Reverse light diagonal stripe | #LXLSX_PATTERN_LIGHT_UP
 *    Light grid                    | #LXLSX_PATTERN_LIGHT_GRID
 *    Light trellis                 | #LXLSX_PATTERN_LIGHT_TRELLIS
 *    12.5% gray                    | #LXLSX_PATTERN_GRAY_125
 *    6.25% gray                    | #LXLSX_PATTERN_GRAY_0625
 *
 */
void lxlsx_format_set_pattern(lxlsx_format *format, uint8_t index);

/**
 * @brief Set the pattern background color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern background color.
 *
 * The lxlsx_format_set_bg_color() method can be used to set the background color of
 * a pattern. Patterns are defined via the lxlsx_format_set_pattern() method. If a
 * pattern hasn't been defined then a solid fill pattern is used as the
 * default.
 *
 * Here is an example of how to set up a solid fill in a cell:
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *
 *     lxlsx_format_set_pattern (format, LXLSX_PATTERN_SOLID);
 *     lxlsx_format_set_bg_color(format, LXLSX_COLOR_GREEN);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Ray", format);
 * @endcode
 *
 * @image html formats_set_bg_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
void lxlsx_format_set_bg_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the pattern foreground color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern foreground  color.
 *
 * The lxlsx_format_set_fg_color() method can be used to set the foreground color of
 * a pattern.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
void lxlsx_format_set_fg_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell border style:
 *
 * @code
 *     lxlsx_format_set_border(format, LXLSX_BORDER_THIN);
 * @endcode
 *
 * Individual border elements can be configured using the following functions with
 * the same parameters:
 *
 * - lxlsx_format_set_bottom()
 * - lxlsx_format_set_top()
 * - lxlsx_format_set_left()
 * - lxlsx_format_set_right()
 *
 * A cell border is comprised of a border on the bottom, top, left and right.
 * These can be set to the same value using lxlsx_format_set_border() or
 * individually using the relevant method calls shown above.
 *
 * The following border styles are available:
 *
 * - #LXLSX_BORDER_THIN
 * - #LXLSX_BORDER_MEDIUM
 * - #LXLSX_BORDER_DASHED
 * - #LXLSX_BORDER_DOTTED
 * - #LXLSX_BORDER_THICK
 * - #LXLSX_BORDER_DOUBLE
 * - #LXLSX_BORDER_HAIR
 * - #LXLSX_BORDER_MEDIUM_DASHED
 * - #LXLSX_BORDER_DASH_DOT
 * - #LXLSX_BORDER_MEDIUM_DASH_DOT
 * - #LXLSX_BORDER_DASH_DOT_DOT
 * - #LXLSX_BORDER_MEDIUM_DASH_DOT_DOT
 * - #LXLSX_BORDER_SLANT_DASH_DOT
 *
 *  The most commonly used style is the `thin` style.
 */
void lxlsx_format_set_border(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the cell bottom border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell bottom border style. See lxlsx_format_set_border() for details on the
 * border styles.
 */
void lxlsx_format_set_bottom(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the cell top border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell top border style. See lxlsx_format_set_border() for details on the border
 * styles.
 */
void lxlsx_format_set_top(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the cell left border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell left border style. See lxlsx_format_set_border() for details on the
 * border styles.
 */
void lxlsx_format_set_left(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the cell right border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell right border style. See lxlsx_format_set_border() for details on the
 * border styles.
 */
void lxlsx_format_set_right(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the color of the cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * Individual border elements can be configured using the following methods with
 * the same parameters:
 *
 * - lxlsx_format_set_bottom_color()
 * - lxlsx_format_set_top_color()
 * - lxlsx_format_set_left_color()
 * - lxlsx_format_set_right_color()
 *
 * Set the color of the cell borders. A cell border is comprised of a border
 * on the bottom, top, left and right. These can be set to the same color
 * using lxlsx_format_set_border_color() or individually using the relevant method
 * calls shown above.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
void lxlsx_format_set_border_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the color of the bottom cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See lxlsx_format_set_border_color() for details on the border colors.
 */
void lxlsx_format_set_bottom_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the color of the top cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See lxlsx_format_set_border_color() for details on the border colors.
 */
void lxlsx_format_set_top_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the color of the left cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See lxlsx_format_set_border_color() for details on the border colors.
 */
void lxlsx_format_set_left_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the color of the right cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See lxlsx_format_set_border_color() for details on the border colors.
 */
void lxlsx_format_set_right_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Set the diagonal cell border type.
 *
 * @param format Pointer to a Format instance.
 * @param type   The #lxlsx_format_diagonal_types diagonal border type.
 *
 * Set the diagonal cell border type:
 *
 * @code
 *     lxlsx_format *format1 = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_diag_type(  format1, LXLSX_DIAGONAL_BORDER_UP);
 *
 *     lxlsx_format *format2 = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_diag_type(  format2, LXLSX_DIAGONAL_BORDER_DOWN);
 *
 *     lxlsx_format *format3 = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_diag_type(  format3, LXLSX_DIAGONAL_BORDER_UP_DOWN);
 *
 *     lxlsx_format *format4 = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_diag_type(  format4, LXLSX_DIAGONAL_BORDER_UP_DOWN);
 *     lxlsx_format_set_diag_border(format4, LXLSX_BORDER_HAIR);
 *     lxlsx_format_set_diag_color( format4, LXLSX_COLOR_RED);
 *
 *     lxlsx_worksheet_write_string(worksheet, CELL("B3"),  "Text", format1);
 *     lxlsx_worksheet_write_string(worksheet, CELL("B6"),  "Text", format2);
 *     lxlsx_worksheet_write_string(worksheet, CELL("B9"),  "Text", format3);
 *     lxlsx_worksheet_write_string(worksheet, CELL("B12"), "Text", format4);
 * @endcode
 *
 * @image html diagonal_border.png
 *
 * The allowable border types are defined in #lxlsx_format_diagonal_types:
 *
 * - #LXLSX_DIAGONAL_BORDER_UP: Cell diagonal border from bottom left to top
 *   right.
 *
 * - #LXLSX_DIAGONAL_BORDER_DOWN: Cell diagonal border from top left to bottom
 *   right.
 *
 * - #LXLSX_DIAGONAL_BORDER_UP_DOWN: Cell diagonal border from top left to
 *   bottom right. A combination of the 2 previous types.
 *
 * If the border style isn't specified with `lxlsx_format_set_diag_border()` then it
 * will default to #LXLSX_BORDER_THIN.
 */
void lxlsx_format_set_diag_type(lxlsx_format *format, uint8_t type);

/**
 * @brief Set the diagonal cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  The #lxlsx_format_borders style.
 *
 * Set the diagonal border style. This should be a #lxlsx_format_borders value.
 * See the example above.
 *
 */
void lxlsx_format_set_diag_border(lxlsx_format *format, uint8_t style);

/**
 * @brief Set the diagonal cell border color.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell diagonal border color.
 *
 * Set the diagonal border color. The color should be an RGB integer value,
 * see @ref working_with_colors and the above example.
 */
void lxlsx_format_set_diag_color(lxlsx_format *format, lxlsx_color_t color);

/**
 * @brief Turn on quote prefix for the format.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the quote prefix property of a format to ensure a string is treated
 * as a string after editing. This is the same as prefixing the string with
 * a single quote in Excel. You don't need to add the quote to the
 * string but you do need to add the format.
 *
 * @code
 *     format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_quote_prefix(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "=Foo", format);
 * @endcode
 *
 */
void lxlsx_format_set_quote_prefix(lxlsx_format *format);

void lxlsx_format_set_font_outline(lxlsx_format *format);
void lxlsx_format_set_font_shadow(lxlsx_format *format);
void lxlsx_format_set_font_scheme(lxlsx_format *format, const char *font_scheme);
void lxlsx_format_set_font_condense(lxlsx_format *format);
void lxlsx_format_set_font_extend(lxlsx_format *format);
void lxlsx_format_set_reading_order(lxlsx_format *format, uint8_t value);
void lxlsx_format_set_theme(lxlsx_format *format, uint8_t value);
void lxlsx_format_set_hyperlink(lxlsx_format *format);
void lxlsx_format_set_color_indexed(lxlsx_format *format, uint8_t value);
void lxlsx_format_set_font_only(lxlsx_format *format);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_FORMAT_H__ */
