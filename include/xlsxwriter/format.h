/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page format_page The Format object
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
 * Formats in `libxlsxwriter` are accessed via the lxw_format
 * struct. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the workbook_add_format() method as
 * follows:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 * @endcode
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in this section.
 * For example:
 *
 * @code
 *    // Create the Format.
 *    lxw_format *format = workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    format_set_bold(format);
 *    format_set_font_color(format, LXW_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 * The full range of formatting options that can be applied using
 * `libxlsxwriter` are shown below.
 *
 */
#ifndef __LXW_FORMAT_H__
#define __LXW_FORMAT_H__

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
typedef int32_t lxw_color_t;

#define LXW_FORMAT_FIELD_LEN            128
#define LXW_DEFAULT_FONT_NAME           "Calibri"
#define LXW_DEFAULT_FONT_FAMILY         2
#define LXW_DEFAULT_FONT_THEME          1
#define LXW_PROPERTY_UNSET              -1
#define LXW_COLOR_UNSET                 -1
#define LXW_COLOR_MASK                  0xFFFFFF
#define LXW_MIN_FONT_SIZE               1.0
#define LXW_MAX_FONT_SIZE               409.0

#define LXW_FORMAT_FIELD_COPY(dst, src)             \
    do{                                             \
        strncpy(dst, src, LXW_FORMAT_FIELD_LEN -1); \
        dst[LXW_FORMAT_FIELD_LEN - 1] = '\0';       \
    } while (0)

/** Format underline values for format_set_underline(). */
enum lxw_format_underlines {
    /** Single underline */
    LXW_UNDERLINE_SINGLE = 1,

    /** Double underline */
    LXW_UNDERLINE_DOUBLE,

    /** Single accounting underline */
    LXW_UNDERLINE_SINGLE_ACCOUNTING,

    /** Double accounting underline */
    LXW_UNDERLINE_DOUBLE_ACCOUNTING
};

/** Superscript and subscript values for format_set_font_script(). */
enum lxw_format_scripts {

    /** Superscript font */
    LXW_FONT_SUPERSCRIPT = 1,

    /** Subscript font */
    LXW_FONT_SUBSCRIPT
};

/** Alignment values for format_set_align(). */
enum lxw_format_alignments {
    /** No alignment. Cell will use Excel's default for the data type */
    LXW_ALIGN_NONE = 0,

    /** Left horizontal alignment */
    LXW_ALIGN_LEFT,

    /** Center horizontal alignment */
    LXW_ALIGN_CENTER,

    /** Right horizontal alignment */
    LXW_ALIGN_RIGHT,

    /** Cell fill horizontal alignment */
    LXW_ALIGN_FILL,

    /** Justify horizontal alignment */
    LXW_ALIGN_JUSTIFY,

    /** Center Across horizontal alignment */
    LXW_ALIGN_CENTER_ACROSS,

    /** Left horizontal alignment */
    LXW_ALIGN_DISTRIBUTED,

    /** Top vertical alignment */
    LXW_ALIGN_VERTICAL_TOP,

    /** Bottom vertical alignment */
    LXW_ALIGN_VERTICAL_BOTTOM,

    /** Center vertical alignment */
    LXW_ALIGN_VERTICAL_CENTER,

    /** Justify vertical alignment */
    LXW_ALIGN_VERTICAL_JUSTIFY,

    /** Distributed vertical alignment */
    LXW_ALIGN_VERTICAL_DISTRIBUTED
};

enum lxw_format_diagonal_types {
    LXW_DIAGONAL_BORDER_UP = 1,
    LXW_DIAGONAL_BORDER_DOWN,
    LXW_DIAGONAL_BORDER_UP_DOWN
};

/** Predefined values for common colors. */
enum lxw_defined_colors {
    /** Black */
    LXW_COLOR_BLACK = 0x1000000,

    /** Blue */
    LXW_COLOR_BLUE = 0x0000FF,

    /** Brown */
    LXW_COLOR_BROWN = 0x800000,

    /** Cyan */
    LXW_COLOR_CYAN = 0x00FFFF,

    /** Gray */
    LXW_COLOR_GRAY = 0x808080,

    /** Green */
    LXW_COLOR_GREEN = 0x008000,

    /** Lime */
    LXW_COLOR_LIME = 0x00FF00,

    /** Magenta */
    LXW_COLOR_MAGENTA = 0xFF00FF,

    /** Navy */
    LXW_COLOR_NAVY = 0x000080,

    /** Orange */
    LXW_COLOR_ORANGE = 0xFF6600,

    /** Pink */
    LXW_COLOR_PINK = 0xFF00FF,

    /** Purple */
    LXW_COLOR_PURPLE = 0x800080,

    /** Red */
    LXW_COLOR_RED = 0xFF0000,

    /** Silver */
    LXW_COLOR_SILVER = 0xC0C0C0,

    /** White */
    LXW_COLOR_WHITE = 0xFFFFFF,

    /** Yellow */
    LXW_COLOR_YELLOW = 0xFFFF00
};

/** Pattern value for use with format_set_pattern(). */
enum lxw_format_patterns {
    /** Empty pattern */
    LXW_PATTERN_NONE = 0,

    /** Solid pattern */
    LXW_PATTERN_SOLID,

    /** Medium gray pattern */
    LXW_PATTERN_MEDIUM_GRAY,

    /** Dark gray pattern */
    LXW_PATTERN_DARK_GRAY,

    /** Light gray pattern */
    LXW_PATTERN_LIGHT_GRAY,

    /** Dark horizontal line pattern */
    LXW_PATTERN_DARK_HORIZONTAL,

    /** Dark vertical line pattern */
    LXW_PATTERN_DARK_VERTICAL,

    /** Dark diagonal stripe pattern */
    LXW_PATTERN_DARK_DOWN,

    /** Reverse dark diagonal stripe pattern */
    LXW_PATTERN_DARK_UP,

    /** Dark grid pattern */
    LXW_PATTERN_DARK_GRID,

    /** Dark trellis pattern */
    LXW_PATTERN_DARK_TRELLIS,

    /** Light horizontal Line pattern */
    LXW_PATTERN_LIGHT_HORIZONTAL,

    /** Light vertical line pattern */
    LXW_PATTERN_LIGHT_VERTICAL,

    /** Light diagonal stripe pattern */
    LXW_PATTERN_LIGHT_DOWN,

    /** Reverse light diagonal stripe pattern */
    LXW_PATTERN_LIGHT_UP,

    /** Light grid pattern */
    LXW_PATTERN_LIGHT_GRID,

    /** Light trellis pattern */
    LXW_PATTERN_LIGHT_TRELLIS,

    /** 12.5% gray pattern */
    LXW_PATTERN_GRAY_125,

    /** 6.25% gray pattern */
    LXW_PATTERN_GRAY_0625
};

/** Cell border styles for use with format_set_border(). */
enum lxw_format_borders {
    /** No border */
    LXW_BORDER_NONE,

    /** Thin border style */
    LXW_BORDER_THIN,

    /** Medium border style */
    LXW_BORDER_MEDIUM,

    /** Dashed border style */
    LXW_BORDER_DASHED,

    /** Dotted border style */
    LXW_BORDER_DOTTED,

    /** Thick border style */
    LXW_BORDER_THICK,

    /** Double border style */
    LXW_BORDER_DOUBLE,

    /** Hair border style */
    LXW_BORDER_HAIR,

    /** Medium dashed border style */
    LXW_BORDER_MEDIUM_DASHED,

    /** Dash-dot border style */
    LXW_BORDER_DASH_DOT,

    /** Medium dash-dot border style */
    LXW_BORDER_MEDIUM_DASH_DOT,

    /** Dash-dot-dot border style */
    LXW_BORDER_DASH_DOT_DOT,

    /** Medium dash-dot-dot border style */
    LXW_BORDER_MEDIUM_DASH_DOT_DOT,

    /** Slant dash-dot border style */
    LXW_BORDER_SLANT_DASH_DOT
};

/**
 * @brief Struct to represent the formatting properties of an Excel format.
 *
 * Formats in `libxlsxwriter` are accessed via this struct.
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in format.h.
 *
 * For example:
 *
 * @code
 *    // Create the Format.
 *    lxw_format *format = workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    format_set_bold(format);
 *    format_set_font_color(format, LXW_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
 *
 * @endcode
 *
 */
typedef struct lxw_format {

    FILE *file;

    lxw_hash_table *xf_format_indices;
    uint16_t *num_xf_formats;

    int32_t xf_index;
    int32_t dxf_index;

    char num_format[LXW_FORMAT_FIELD_LEN];
    char font_name[LXW_FORMAT_FIELD_LEN];
    char font_scheme[LXW_FORMAT_FIELD_LEN];
    uint16_t num_format_index;
    uint16_t font_index;
    uint8_t has_font;
    uint8_t has_dxf_font;
    double font_size;
    uint8_t bold;
    uint8_t italic;
    lxw_color_t font_color;
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

    lxw_color_t fg_color;
    lxw_color_t bg_color;
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
    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;

    uint8_t indent;
    uint8_t shrink;
    uint8_t merge_range;
    uint8_t reading_order;
    uint8_t just_distrib;
    uint8_t color_indexed;
    uint8_t font_only;

    STAILQ_ENTRY (lxw_format) list_pointers;
} lxw_format;

/*
 * Struct to represent the font component of a format.
 */
typedef struct lxw_font {

    char font_name[LXW_FORMAT_FIELD_LEN];
    double font_size;
    uint8_t bold;
    uint8_t italic;
    uint8_t underline;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    lxw_color_t font_color;
} lxw_font;

/*
 * Struct to represent the border component of a format.
 */
typedef struct lxw_border {

    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;

    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;

} lxw_border;

/*
 * Struct to represent the fill component of a format.
 */
typedef struct lxw_fill {

    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t pattern;

} lxw_fill;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_format *lxw_format_new();
void lxw_format_free(lxw_format *format);
int32_t lxw_format_get_xf_index(lxw_format *format);
lxw_font *lxw_format_get_font_key(lxw_format *format);
lxw_border *lxw_format_get_border_key(lxw_format *format);
lxw_fill *lxw_format_get_fill_key(lxw_format *format);

lxw_color_t lxw_format_check_color(lxw_color_t color);

/**
 * @brief Set the font used in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param font_name Cell font name.
 *
 * Specify the font used used in the cell format:
 *
 * @code
 *     format_set_font_name(format, "Avenir Black Oblique");
 * @endcode
 *
 * @image html format_set_font_name.png
 *
 * Excel can only display fonts that are installed on the system that it is
 * running on. Therefore it is generally best to use the fonts that come as
 * standard with Excel such as Calibri, Times New Roman and Courier New.
 *
 * The default font in Excel 2007, and later, is Calibri.
 */
void format_set_font_name(lxw_format *format, const char *font_name);

/**
 * @brief Set the size of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param size   The cell font size.
 *
 * Set the font size of the cell format:
 *
 * @code
 *     format_set_font_size(format, 30);
 * @endcode
 *
 * @image html format_font_size.png
 *
 * Excel adjusts the height of a row to accommodate the largest font
 * size in the row. You can also explicitly specify the height of a
 * row using the worksheet_set_row() function.
 */
void format_set_font_size(lxw_format *format, double size);

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
 *     format = workbook_add_format(workbook);
 *     format_set_font_color(format, LXW_COLOR_RED);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Wheelbarrow", format);
 * @endcode
 *
 * @image html format_font_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * @note
 * The format_set_font_color() method is used to set the font color in a
 * cell. To set the color of a cell background use the format_set_bg_color()
 * and format_set_pattern() methods.
 */
void format_set_font_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Turn on bold for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the bold property of the font:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_bold(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Bold Text", format);
 * @endcode
 *
 * @image html format_font_bold.png
 */
void format_set_bold(lxw_format *format);

/**
 * @brief Turn on italic for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the italic property of the font:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_italic(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Italic Text", format);
 * @endcode
 *
 * @image html format_font_italic.png
 */
void format_set_italic(lxw_format *format);

/**
 * @brief Turn on underline for the format:
 *
 * @param format Pointer to a Format instance.
 * @param style Underline style.
 *
 * Set the underline property of the format:
 *
 * @code
 *     format_set_underline(format, LXW_UNDERLINE_SINGLE);
 * @endcode
 *
 * @image html format_font_underlined.png
 *
 * The available underline styles are:
 *
 * - #LXW_UNDERLINE_SINGLE
 * - #LXW_UNDERLINE_DOUBLE
 * - #LXW_UNDERLINE_SINGLE_ACCOUNTING
 * - #LXW_UNDERLINE_DOUBLE_ACCOUNTING
 *
 */
void format_set_underline(lxw_format *format, uint8_t style);

/**
 * @brief Set the strikeout property of the font.
 *
 * @param format Pointer to a Format instance.
 *
 * @image html format_font_strikeout.png
 *
 */
void format_set_font_strikeout(lxw_format *format);

/**
 * @brief Set the superscript/subscript property of the font.
 *
 * @param format Pointer to a Format instance.
 * @param style  Superscript or subscript style.
 *
 * Set the superscript o subscript property of the font.
 *
 * @image html format_font_script.png
 *
 * The available script styles are:
 *
 * - #LXW_FONT_SUPERSCRIPT
 * - #LXW_FONT_SUBSCRIPT
 */
void format_set_font_script(lxw_format *format, uint8_t style);

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
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, "d mmm yyyy");
 * @endcode
 *
 * Format strings can control any aspect of number formatting allowed by Excel:
 *
 * @dontinclude format_num_format.c
 * @skipline set_num_format
 * @until 1209
 *
 * @image html format_set_num_format.png
 *
 * The number system used for dates is described in @ref working_with_dates.
 *
 * For more information on number formats in Excel refer to the
 * [Microsoft documentation on cell formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
 */
void format_set_num_format(lxw_format *format, const char *num_format);

/**
 * @brief Set the Excel built-in number format for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param index  The built-in number format index for the cell.
 *
 * This function is similar to format_set_num_format() except that it takes an
 * index to a limited number of Excel's built-in number formats instead of a
 * user defined format string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, 0x0F);     // d-mmm-yy
 * @endcode
 *
 * @note
 * Unless you need to specifically access one of Excel's built-in number
 * formats the format_set_num_format() function above is a better
 * solution. The format_set_num_format_index() function is mainly included for
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
 *  - These formats can also be set via format_set_num_format().
 */
void format_set_num_format_index(lxw_format *format, uint8_t index);

/**
 * @brief Set the cell unlocked state.
 *
 * @param format Pointer to a Format instance.
 *
 * This property can be used to allow modification of a cell in a protected
 * worksheet. In Excel, cell locking is turned on by default for all
 * cells. However, it only has an effect if the worksheet has been protected
 * using the worksheet worksheet_protect() function:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_unlocked(format);
 *
 *     // Enable worksheet protection, without password or options.
 *     worksheet_protect(worksheet, NULL, NULL);
 *
 *     // This cell cannot be edited.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL);
 *
 *     // This cell can be edited.
 *     worksheet_write_formula(worksheet, 1, 0, "=1+2", format);
 * @endcode
 */
void format_set_unlocked(lxw_format *format);

/**
 * @brief Hide formulas in a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This property is used to hide a formula while still displaying its
 * result. This is generally used to hide complex calculations from end users
 * who are only interested in the result. It only has an effect if the
 * worksheet has been protected using the worksheet worksheet_protect()
 * function:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_hidden(format);
 *
 *     // Enable worksheet protection, without password or options.
 *     worksheet_protect(worksheet, NULL, NULL);
 *
 *     // The formula in this cell isn't visible.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", format);
 * @endcode
 */
void format_set_hidden(lxw_format *format);

/**
 * @brief Set the alignment for data in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param alignment The horizontal and or vertical alignment direction.
 *
 * This method is used to set the horizontal and vertical text alignment within a
 * cell. The following are the available horizontal alignments:
 *
 * - #LXW_ALIGN_LEFT
 * - #LXW_ALIGN_CENTER
 * - #LXW_ALIGN_RIGHT
 * - #LXW_ALIGN_FILL
 * - #LXW_ALIGN_JUSTIFY
 * - #LXW_ALIGN_CENTER_ACROSS
 * - #LXW_ALIGN_DISTRIBUTED
 *
 * The following are the available vertical alignments:
 *
 * - #LXW_ALIGN_VERTICAL_TOP
 * - #LXW_ALIGN_VERTICAL_BOTTOM
 * - #LXW_ALIGN_VERTICAL_CENTER
 * - #LXW_ALIGN_VERTICAL_JUSTIFY
 * - #LXW_ALIGN_VERTICAL_DISTRIBUTED
 *
 * As in Excel, vertical and horizontal alignments can be combined:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_align(format, LXW_ALIGN_CENTER);
 *     format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
 *
 *     worksheet_set_row(0, 30);
 *     worksheet_write_string(worksheet, 0, 0, "Some Text", format);
 * @endcode
 *
 * @image html format_font_align.png
 *
 * Text can be aligned across two or more adjacent cells using the
 * center_across property. However, for genuine merged cells it is better to
 * use the worksheet_merge_range() worksheet method.
 *
 * The vertical justify option can be used to provide automatic text wrapping
 * in a cell. The height of the cell will be adjusted to accommodate the
 * wrapped text. To specify where the text wraps use the
 * format_set_text_wrap() method.
 */
void format_set_align(lxw_format *format, uint8_t alignment);

/**
 * @brief Wrap text in a cell.
 *
 * Turn text wrapping on for text in a cell.
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a cell", format);
 * @endcode
 *
 * If you wish to control where the text is wrapped you can add newline characters
 * to the string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format);
 * @endcode
 *
 * @image html format_font_text_wrap.png
 *
 * Excel will adjust the height of the row to accommodate the wrapped text. A
 * similar effect can be obtained without newlines using the
 * format_set_align() function with #LXW_ALIGN_VERTICAL_JUSTIFY.
 */
void format_set_text_wrap(lxw_format *format);

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
 *     format = workbook_add_format(workbook);
 *     format_set_rotation(format, 30);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is rotated", format);
 * @endcode
 *
 * @image html format_font_text_rotated.png
 *
 * The angle 270 is also supported. This indicates text where the letters run from
 * top to bottom.
 */
void format_set_rotation(lxw_format *format, int16_t angle);

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
 *     format1 = workbook_add_format(workbook);
 *     format2 = workbook_add_format(workbook);
 *
 *     format_set_indent(format1, 1);
 *     format_set_indent(format2, 2);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",  format1);
 *     worksheet_write_string(worksheet, 1, 0, "This text is indented 2 levels", format2);
 * @endcode
 *
 * @image html text_indent.png
 *
 * @note
 * Indentation is a horizontal alignment property. It will override any other
 * horizontal properties but it can be used in conjunction with vertical
 * properties.
 */
void format_set_indent(lxw_format *format, uint8_t level);

/**
 * @brief Turn on the text "shrink to fit" for a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This method can be used to shrink text so that it fits in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_shrink(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!", format);
 * @endcode
 */
void format_set_shrink(lxw_format *format);

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
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_YELLOW);
 * @endcode
 *
 * The available fill patterns are:
 *
 *    Fill Type                     | Define
 *    ----------------------------- | -----------------------------
 *    Solid                         | #LXW_PATTERN_SOLID
 *    Medium gray                   | #LXW_PATTERN_MEDIUM_GRAY
 *    Dark gray                     | #LXW_PATTERN_DARK_GRAY
 *    Light gray                    | #LXW_PATTERN_LIGHT_GRAY
 *    Dark horizontal line          | #LXW_PATTERN_DARK_HORIZONTAL
 *    Dark vertical line            | #LXW_PATTERN_DARK_VERTICAL
 *    Dark diagonal stripe          | #LXW_PATTERN_DARK_DOWN
 *    Reverse dark diagonal stripe  | #LXW_PATTERN_DARK_UP
 *    Dark grid                     | #LXW_PATTERN_DARK_GRID
 *    Dark trellis                  | #LXW_PATTERN_DARK_TRELLIS
 *    Light horizontal line         | #LXW_PATTERN_LIGHT_HORIZONTAL
 *    Light vertical line           | #LXW_PATTERN_LIGHT_VERTICAL
 *    Light diagonal stripe         | #LXW_PATTERN_LIGHT_DOWN
 *    Reverse light diagonal stripe | #LXW_PATTERN_LIGHT_UP
 *    Light grid                    | #LXW_PATTERN_LIGHT_GRID
 *    Light trellis                 | #LXW_PATTERN_LIGHT_TRELLIS
 *    12.5% gray                    | #LXW_PATTERN_GRAY_125
 *    6.25% gray                    | #LXW_PATTERN_GRAY_0625
 *
 */
void format_set_pattern(lxw_format *format, uint8_t index);

/**
 * @brief Set the pattern background color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern background color.
 *
 * The format_set_bg_color() method can be used to set the background color of
 * a pattern. Patterns are defined via the format_set_pattern() method. If a
 * pattern hasn't been defined then a solid fill pattern is used as the
 * default.
 *
 * Here is an example of how to set up a solid fill in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_GREEN);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Ray", format);
 * @endcode
 *
 * @image html formats_set_bg_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
void format_set_bg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the pattern foreground color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern foreground  color.
 *
 * The format_set_fg_color() method can be used to set the foreground color of
 * a pattern.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
void format_set_fg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell border style:
 *
 * @code
 *     format_set_border(format, LXW_BORDER_THIN);
 * @endcode
 *
 * Individual border elements can be configured using the following functions with
 * the same parameters:
 *
 * - format_set_bottom()
 * - format_set_top()
 * - format_set_left()
 * - format_set_right()
 *
 * A cell border is comprised of a border on the bottom, top, left and right.
 * These can be set to the same value using format_set_border() or
 * individually using the relevant method calls shown above.
 *
 * The following border styles are available:
 *
 * - #LXW_BORDER_THIN
 * - #LXW_BORDER_MEDIUM
 * - #LXW_BORDER_DASHED
 * - #LXW_BORDER_DOTTED
 * - #LXW_BORDER_THICK
 * - #LXW_BORDER_DOUBLE
 * - #LXW_BORDER_HAIR
 * - #LXW_BORDER_MEDIUM_DASHED
 * - #LXW_BORDER_DASH_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT
 * - #LXW_BORDER_DASH_DOT_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT_DOT
 * - #LXW_BORDER_SLANT_DASH_DOT
 *
 *  The most commonly used style is the `thin` style.
 */
void format_set_border(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell bottom border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell bottom border style. See format_set_border() for details on the
 * border styles.
 */
void format_set_bottom(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell top border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell top border style. See format_set_border() for details on the border
 * styles.
 */
void format_set_top(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell left border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell left border style. See format_set_border() for details on the
 * border styles.
 */
void format_set_left(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell right border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell right border style. See format_set_border() for details on the
 * border styles.
 */
void format_set_right(lxw_format *format, uint8_t style);

/**
 * @brief Set the color of the cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * Individual border elements can be configured using the following methods with
 * the same parameters:
 *
 * - format_set_bottom_color()
 * - format_set_top_color()
 * - format_set_left_color()
 * - format_set_right_color()
 *
 * Set the color of the cell borders. A cell border is comprised of a border
 * on the bottom, top, left and right. These can be set to the same color
 * using format_set_border_color() or individually using the relevant method
 * calls shown above.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
void format_set_border_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the bottom cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
void format_set_bottom_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the top cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
void format_set_top_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the left cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
void format_set_left_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the right cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
void format_set_right_color(lxw_format *format, lxw_color_t color);

void format_set_diag_type(lxw_format *format, uint8_t value);
void format_set_diag_color(lxw_format *format, lxw_color_t color);
void format_set_diag_border(lxw_format *format, uint8_t value);
void format_set_font_outline(lxw_format *format);
void format_set_font_shadow(lxw_format *format);
void format_set_font_family(lxw_format *format, uint8_t value);
void format_set_font_charset(lxw_format *format, uint8_t value);
void format_set_font_scheme(lxw_format *format, const char *font_scheme);
void format_set_font_condense(lxw_format *format);
void format_set_font_extend(lxw_format *format);
void format_set_reading_order(lxw_format *format, uint8_t value);
void format_set_theme(lxw_format *format, uint8_t value);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_FORMAT_H__ */
