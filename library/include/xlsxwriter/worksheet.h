/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page worksheet_page The Worksheet object
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * See @ref worksheet.h for full details of the functionality.
 *
 * @file worksheet.h
 *
 * @brief Functions related to adding data and formatting to a worksheet.
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * A Worksheet object isn't created directly. Instead a worksheet is
 * created by calling the workbook_add_worksheet() function from a
 * Workbook object:
 *
 * @code
 *     #include "xlsxwriter.h"
 *
 *     int main() {
 *
 *         lxw_workbook  *workbook  = workbook_new("filename.xlsx");
 *         lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
 *
 *         worksheet_write_string(worksheet, 0, 0, "Hello Excel", NULL);
 *
 *         return workbook_close(workbook);
 *     }
 * @endcode
 *
 */
#ifndef __LXW_WORKSHEET_H__
#define __LXW_WORKSHEET_H__

#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include <string.h>

#include "shared_strings.h"
#include "chart.h"
#include "drawing.h"
#include "common.h"
#include "format.h"
#include "utility.h"

#define LXW_ROW_MAX 1048576
#define LXW_COL_MAX 16384
#define LXW_COL_META_MAX 128
#define LXW_HEADER_FOOTER_MAX 255
#define LXW_MAX_NUMBER_URLS 65530
#define LXW_PANE_NAME_LENGTH 12 /* bottomRight + 1 */

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
#define LXW_BREAKS_MAX 1023

/** Default column width in Excel */
#define LXW_DEF_COL_WIDTH (double)8.43

/** Default row height in Excel */
#define LXW_DEF_ROW_HEIGHT (double)15.0

/** Gridline options using in `worksheet_gridlines()`. */
enum lxw_gridlines {
    /** Hide screen and print gridlines. */
    LXW_HIDE_ALL_GRIDLINES = 0,

    /** Show screen gridlines. */
    LXW_SHOW_SCREEN_GRIDLINES,

    /** Show print gridlines. */
    LXW_SHOW_PRINT_GRIDLINES,

    /** Show screen and print gridlines. */
    LXW_SHOW_ALL_GRIDLINES
};

/** Data validation property values. */
enum lxw_validation_boolean {
    LXW_VALIDATION_DEFAULT,

    /** Turn a data validation property off. */
    LXW_VALIDATION_OFF,

    /** Turn a data validation property on. Data validation properties are
     * generally on by default. */
    LXW_VALIDATION_ON
};

/** Data validation types. */
enum lxw_validation_types {
    LXW_VALIDATION_TYPE_NONE,

    /** Restrict cell input to whole/integer numbers only. */
    LXW_VALIDATION_TYPE_INTEGER,

    /** Restrict cell input to whole/integer numbers only, using a cell
     *  reference. */
    LXW_VALIDATION_TYPE_INTEGER_FORMULA,

    /** Restrict cell input to decimal numbers only. */
    LXW_VALIDATION_TYPE_DECIMAL,

    /** Restrict cell input to decimal numbers only, using a cell
     * reference. */
    LXW_VALIDATION_TYPE_DECIMAL_FORMULA,

    /** Restrict cell input to a list of strings in a dropdown. */
    LXW_VALIDATION_TYPE_LIST,

    /** Restrict cell input to a list of strings in a dropdown, using a
     * cell range. */
    LXW_VALIDATION_TYPE_LIST_FORMULA,

    /** Restrict cell input to date values only, using a lxw_datetime type. */
    LXW_VALIDATION_TYPE_DATE,

    /** Restrict cell input to date values only, using a cell reference. */
    LXW_VALIDATION_TYPE_DATE_FORMULA,

    /* Restrict cell input to date values only, as a serial number.
     * Undocumented. */
    LXW_VALIDATION_TYPE_DATE_NUMBER,

    /** Restrict cell input to time values only, using a lxw_datetime type. */
    LXW_VALIDATION_TYPE_TIME,

    /** Restrict cell input to time values only, using a cell reference. */
    LXW_VALIDATION_TYPE_TIME_FORMULA,

    /* Restrict cell input to time values only, as a serial number.
     * Undocumented. */
    LXW_VALIDATION_TYPE_TIME_NUMBER,

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    LXW_VALIDATION_TYPE_LENGTH,

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    LXW_VALIDATION_TYPE_LENGTH_FORMULA,

    /** Restrict cell to input controlled by a custom formula that returns
     * `TRUE/FALSE`. */
    LXW_VALIDATION_TYPE_CUSTOM_FORMULA,

    /** Allow any type of input. Mainly only useful for pop-up messages. */
    LXW_VALIDATION_TYPE_ANY
};

/** Data validation criteria uses to control the selection of data. */
enum lxw_validation_criteria {
    LXW_VALIDATION_CRITERIA_NONE,

    /** Select data between two values. */
    LXW_VALIDATION_CRITERIA_BETWEEN,

    /** Select data that is not between two values. */
    LXW_VALIDATION_CRITERIA_NOT_BETWEEN,

    /** Select data equal to a value. */
    LXW_VALIDATION_CRITERIA_EQUAL_TO,

    /** Select data not equal to a value. */
    LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,

    /** Select data greater than a value. */
    LXW_VALIDATION_CRITERIA_GREATER_THAN,

    /** Select data less than a value. */
    LXW_VALIDATION_CRITERIA_LESS_THAN,

    /** Select data greater than or equal to a value. */
    LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,

    /** Select data less than or equal to a value. */
    LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO
};

/** Data validation error types for pop-up messages. */
enum lxw_validation_error_types {
    /** Show a "Stop" data validation pop-up message. This is the default. */
    LXW_VALIDATION_ERROR_TYPE_STOP,

    /** Show an "Error" data validation pop-up message. */
    LXW_VALIDATION_ERROR_TYPE_WARNING,

    /** Show an "Information" data validation pop-up message. */
    LXW_VALIDATION_ERROR_TYPE_INFORMATION
};

enum cell_types {
    NUMBER_CELL = 1,
    STRING_CELL,
    INLINE_STRING_CELL,
    FORMULA_CELL,
    ARRAY_FORMULA_CELL,
    BLANK_CELL,
    BOOLEAN_CELL,
    HYPERLINK_URL,
    HYPERLINK_INTERNAL,
    HYPERLINK_EXTERNAL
};

enum pane_types {
    NO_PANES = 0,
    FREEZE_PANES,
    SPLIT_PANES,
    FREEZE_SPLIT_PANES
};

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxw_table_cells, lxw_cell);

/* Define a RB_TREE struct manually to add extra members. */
struct lxw_table_rows {
    struct lxw_row *rbh_root;
    struct lxw_row *cached_row;
    lxw_row_t cached_row_num;
};

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXW_RB_GENERATE_ROW(name, type, field, cmp)       \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxw_rb_generate_row{int unused;}

#define LXW_RB_GENERATE_CELL(name, type, field, cmp)      \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxw_rb_generate_cell{int unused;}

STAILQ_HEAD(lxw_merged_ranges, lxw_merged_range);
STAILQ_HEAD(lxw_selections, lxw_selection);
STAILQ_HEAD(lxw_data_validations, lxw_data_validation);
STAILQ_HEAD(lxw_image_data, lxw_image_options);
STAILQ_HEAD(lxw_chart_data, lxw_image_options);

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the worksheet_set_column() and worksheet_set_row()
 * functions.
 *
 * It has the following members:
 *
 * * `hidden`
 * * `level`
 * * `collapsed`
 *
 * The members of this struct are explained in @ref ww_outlines_grouping.
 *
 */
typedef struct lxw_row_col_options {
    /** Hide the row/column */
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxw_row_col_options;

typedef struct lxw_col_options {
    lxw_col_t firstcol;
    lxw_col_t lastcol;
    double width;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxw_col_options;

typedef struct lxw_merged_range {
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;

    STAILQ_ENTRY (lxw_merged_range) list_pointers;
} lxw_merged_range;

typedef struct lxw_repeat_rows {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
} lxw_repeat_rows;

typedef struct lxw_repeat_cols {
    uint8_t in_use;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_repeat_cols;

typedef struct lxw_print_area {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_print_area;

typedef struct lxw_autofilter {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
} lxw_autofilter;

typedef struct lxw_panes {
    uint8_t type;
    lxw_row_t first_row;
    lxw_col_t first_col;
    lxw_row_t top_row;
    lxw_col_t left_col;
    double x_split;
    double y_split;
} lxw_panes;

typedef struct lxw_selection {
    char pane[LXW_PANE_NAME_LENGTH];
    char active_cell[LXW_MAX_CELL_RANGE_LENGTH];
    char sqref[LXW_MAX_CELL_RANGE_LENGTH];

    STAILQ_ENTRY (lxw_selection) list_pointers;

} lxw_selection;

/**
 * @brief Worksheet data validation options.
 */
typedef struct lxw_data_validation {

    /**
     * Set the validation type. Should be a #lxw_validation_types value.
     */
    uint8_t validate;

    /**
     * Set the validation criteria type to select the data. Should be a
     * #lxw_validation_criteria value.
     */
    uint8_t criteria;

    /** Controls whether a data validation is not applied to blank data in the
     * cell. Should be a #lxw_validation_boolean value. It is on by
     * default.
     */
    uint8_t ignore_blank;

    /**
     * This parameter is used to toggle on and off the 'Show input message
     * when cell is selected' option in the Excel data validation dialog. When
     * the option is off an input message is not displayed even if it has been
     * set using input_message. Should be a #lxw_validation_boolean value. It
     * is on by default.
     */
    uint8_t show_input;

    /**
     * This parameter is used to toggle on and off the 'Show error alert
     * after invalid data is entered' option in the Excel data validation
     * dialog. When the option is off an error message is not displayed even
     * if it has been set using error_message. Should be a
     * #lxw_validation_boolean value. It is on by default.
     */
    uint8_t show_error;

    /**
     * This parameter is used to specify the type of error dialog that is
     * displayed. Should be a #lxw_validation_error_types value.
     */
    uint8_t error_type;

    /**
     * This parameter is used to toggle on and off the 'In-cell dropdown'
     * option in the Excel data validation dialog. When the option is on a
     * dropdown list will be shown for list validations. Should be a
     * #lxw_validation_boolean value. It is on by default.
     */
    uint8_t dropdown;

    uint8_t is_between;

    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a whole or decimal number.
     */
    double value_number;

    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a cell reference. It is valid for any of the
     * `_FORMULA` validation types.
     */
    char *value_formula;

    /**
     * This parameter is used to set a list of strings for a drop down list.
     * The list should be a `NULL` terminated array of char* strings:
     *
     * @code
     *    char *list[] = {"open", "high", "close", NULL};
     *
     *    data_validation->validate   = LXW_VALIDATION_TYPE_LIST;
     *    data_validation->value_list = list;
     * @endcode
     *
     * The `value_formula` parameter can also be used to specify a list from
     * an Excel cell range.
     *
     * Note, the string list is restricted by Excel to 255 characters,
     * including comma separators.
     */
    char **value_list;

    /**
     * This parameter is used to set the limiting value to which the date or
     * time criteria is applied using a #lxw_datetime struct.
     */
    lxw_datetime value_datetime;

    /**
     * This parameter is the same as `value_number` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    double minimum_number;

    /**
     * This parameter is the same as `value_formula` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    char *minimum_formula;

    /**
     * This parameter is the same as `value_datetime` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    lxw_datetime minimum_datetime;

    /**
     * This parameter is the same as `value_number` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    double maximum_number;

    /**
     * This parameter is the same as `value_formula` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    char *maximum_formula;

    /**
     * This parameter is the same as `value_datetime` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    lxw_datetime maximum_datetime;

    /**
     * The input_title parameter is used to set the title of the input message
     * that is displayed when a cell is entered. It has no default value and
     * is only displayed if the input message is displayed. See the
     * `input_message` parameter below.
     *
     * The maximum title length is 32 characters.
     */
    char *input_title;

    /**
     * The input_message parameter is used to set the input message that is
     * displayed when a cell is entered. It has no default value.
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    char *input_message;

    /**
     * The error_title parameter is used to set the title of the error message
     * that is displayed when the data validation criteria is not met. The
     * default error title is 'Microsoft Excel'. The maximum title length is
     * 32 characters.
     */
    char *error_title;

    /**
     * The error_message parameter is used to set the error message that is
     * displayed when a cell is entered. The default error message is "The
     * value you entered is not valid. A user has restricted values that can
     * be entered into the cell".
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    char *error_message;

    char sqref[LXW_MAX_CELL_RANGE_LENGTH];

    STAILQ_ENTRY (lxw_data_validation) list_pointers;

} lxw_data_validation;

/**
 * @brief Options for inserted images
 *
 * Options for modifying images inserted via `worksheet_insert_image_opt()`.
 *
 */
typedef struct lxw_image_options {

    /** Offset from the left of the cell in pixels. */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. */
    int32_t y_offset;

    /** X scale of the image as a decimal. */
    double x_scale;

    /** Y scale of the image as a decimal. */
    double y_scale;

    lxw_row_t row;
    lxw_col_t col;
    char *filename;
    char *url;
    char *tip;
    uint8_t anchor;

    /* Internal metadata. */
    FILE *stream;
    uint8_t image_type;
    double width;
    double height;
    char *short_name;
    char *extension;
    double x_dpi;
    double y_dpi;
    lxw_chart *chart;

    STAILQ_ENTRY (lxw_image_options) list_pointers;

} lxw_image_options;

/**
 * @brief Header and footer options.
 *
 * Optional parameters used in the worksheet_set_header_opt() and
 * worksheet_set_footer_opt() functions.
 *
 */
typedef struct lxw_header_footer_options {
    /** Header or footer margin in inches. Excel default is 0.3. */
    double margin;
} lxw_header_footer_options;

/**
 * @brief Worksheet protection options.
 */
typedef struct lxw_protection {
    /** Turn off selection of locked cells. This in on in Excel by default.*/
    uint8_t no_select_locked_cells;

    /** Turn off selection of unlocked cells. This in on in Excel by default.*/
    uint8_t no_select_unlocked_cells;

    /** Prevent formatting of cells. */
    uint8_t format_cells;

    /** Prevent formatting of columns. */
    uint8_t format_columns;

    /** Prevent formatting of rows. */
    uint8_t format_rows;

    /** Prevent insertion of columns. */
    uint8_t insert_columns;

    /** Prevent insertion of rows. */
    uint8_t insert_rows;

    /** Prevent insertion of hyperlinks. */
    uint8_t insert_hyperlinks;

    /** Prevent deletion of columns. */
    uint8_t delete_columns;

    /** Prevent deletion of rows. */
    uint8_t delete_rows;

    /** Prevent sorting data. */
    uint8_t sort;

    /** Prevent filtering data. */
    uint8_t autofilter;

    /** Prevent insertion of pivot tables. */
    uint8_t pivot_tables;

    /** Protect scenarios. */
    uint8_t scenarios;

    /** Protect drawing objects. */
    uint8_t objects;

    uint8_t no_sheet;
    uint8_t content;
    uint8_t is_configured;
    char hash[5];
} lxw_protection;

/**
 * @brief Struct to represent an Excel worksheet.
 *
 * The members of the lxw_worksheet struct aren't modified directly. Instead
 * the worksheet properties are set by calling the functions shown in
 * worksheet.h.
 */
typedef struct lxw_worksheet {

    FILE *file;
    FILE *optimize_tmpfile;
    struct lxw_table_rows *table;
    struct lxw_table_rows *hyperlinks;
    struct lxw_cell **array;
    struct lxw_merged_ranges *merged_ranges;
    struct lxw_selections *selections;
    struct lxw_data_validations *data_validations;
    struct lxw_image_data *image_data;
    struct lxw_chart_data *chart_data;

    lxw_row_t dim_rowmin;
    lxw_row_t dim_rowmax;
    lxw_col_t dim_colmin;
    lxw_col_t dim_colmax;

    lxw_sst *sst;
    char *name;
    char *quoted_name;
    char *tmpdir;

    uint32_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;

    lxw_col_options **col_options;
    uint16_t col_options_max;

    double *col_sizes;
    uint16_t col_sizes_max;

    lxw_format **col_formats;
    uint16_t col_formats_max;

    uint8_t col_size_changed;
    uint8_t row_size_changed;
    uint8_t optimize;
    struct lxw_row *optimize_row;

    uint16_t fit_height;
    uint16_t fit_width;
    uint16_t horizontal_dpi;
    uint16_t hlink_count;
    uint16_t page_start;
    uint16_t print_scale;
    uint16_t rel_count;
    uint16_t vertical_dpi;
    uint16_t zoom;
    uint8_t filter_on;
    uint8_t fit_page;
    uint8_t hcenter;
    uint8_t orientation;
    uint8_t outline_changed;
    uint8_t outline_on;
    uint8_t outline_style;
    uint8_t outline_below;
    uint8_t outline_right;
    uint8_t page_order;
    uint8_t page_setup_changed;
    uint8_t page_view;
    uint8_t paper_size;
    uint8_t print_gridlines;
    uint8_t print_headers;
    uint8_t print_options_changed;
    uint8_t right_to_left;
    uint8_t screen_gridlines;
    uint8_t show_zeros;
    uint8_t vba_codename;
    uint8_t vcenter;
    uint8_t zoom_scale_normal;
    uint8_t num_validations;

    lxw_color_t tab_color;

    double margin_left;
    double margin_right;
    double margin_top;
    double margin_bottom;
    double margin_header;
    double margin_footer;

    double default_row_height;
    uint32_t default_row_pixels;
    uint32_t default_col_pixels;
    uint8_t default_row_zeroed;
    uint8_t default_row_set;
    uint8_t outline_row_level;
    uint8_t outline_col_level;

    uint8_t header_footer_changed;
    char header[LXW_HEADER_FOOTER_MAX];
    char footer[LXW_HEADER_FOOTER_MAX];

    struct lxw_repeat_rows repeat_rows;
    struct lxw_repeat_cols repeat_cols;
    struct lxw_print_area print_area;
    struct lxw_autofilter autofilter;

    uint16_t merged_range_count;

    lxw_row_t *hbreaks;
    lxw_col_t *vbreaks;
    uint16_t hbreaks_count;
    uint16_t vbreaks_count;

    struct lxw_rel_tuples *external_hyperlinks;
    struct lxw_rel_tuples *external_drawing_links;
    struct lxw_rel_tuples *drawing_links;

    struct lxw_panes panes;

    struct lxw_protection protection;

    lxw_drawing *drawing;

    STAILQ_ENTRY (lxw_worksheet) list_pointers;

} lxw_worksheet;

/*
 * Worksheet initialization data.
 */
typedef struct lxw_worksheet_init_data {
    uint32_t index;
    uint8_t hidden;
    uint8_t optimize;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    lxw_sst *sst;
    char *name;
    char *quoted_name;
    char *tmpdir;

} lxw_worksheet_init_data;

/* Struct to represent a worksheet row. */
typedef struct lxw_row {
    lxw_row_t row_num;
    double height;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
    uint8_t row_changed;
    uint8_t data_changed;
    uint8_t height_changed;
    struct lxw_table_cells *cells;

    /* tree management pointers for tree.h. */
    RB_ENTRY (lxw_row) tree_pointers;
} lxw_row;

/* Struct to represent a worksheet cell. */
typedef struct lxw_cell {
    lxw_row_t row_num;
    lxw_col_t col_num;
    enum cell_types type;
    lxw_format *format;

    union {
        double number;
        int32_t string_id;
        char *string;
    } u;

    double formula_result;
    char *user_data1;
    char *user_data2;
    char *sst_string;

    /* List pointers for tree.h. */
    RB_ENTRY (lxw_cell) tree_pointers;
} lxw_cell;

/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Write a number to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param number    The number to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * The `worksheet_write_number()` function writes numeric types to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_number(worksheet, 0, 0, 123456, NULL);
 *     worksheet_write_number(worksheet, 1, 0, 2.3451, NULL);
 * @endcode
 *
 * @image html write_number01.png
 *
 * The native data type for all numbers in Excel is a IEEE-754 64-bit
 * double-precision floating point, which is also the default type used by
 * `%worksheet_write_number`.
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_num_format(format, "$#,##0.00");
 *
 *     worksheet_write_number(worksheet, 0, 0, 1234.567, format);
 * @endcode
 *
 * @image html write_number02.png
 *
 * @note Excel doesn't support `NaN`, `Inf` or `-Inf` as a number value. If
 * you are writing data that contains these values then your application
 * should convert them to a string or handle them in some other way.
 *
 */
lxw_error worksheet_write_number(lxw_worksheet *worksheet,
                                 lxw_row_t row,
                                 lxw_col_t col, double number,
                                 lxw_format *format);
/**
 * @brief Write a string to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param string    String to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_write_string()` function writes a string to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is English!", NULL);
 * @endcode
 *
 * @image html write_string01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_bold(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is Bold!", format);
 * @endcode
 *
 * @image html write_string02.png
 *
 * Unicode strings are supported in UTF-8 encoding. This generally requires
 * that your source file is UTF-8 encoded or that the data has been read from
 * a UTF-8 source:
 *
 * @code
 *    worksheet_write_string(worksheet, 0, 0, "Это фраза на русском!", NULL);
 * @endcode
 *
 * @image html write_string03.png
 *
 */
lxw_error worksheet_write_string(lxw_worksheet *worksheet,
                                 lxw_row_t row,
                                 lxw_col_t col, const char *string,
                                 lxw_format *format);
/**
 * @brief Write a formula to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_write_formula()` function writes a formula or function to
 * the cell specified by `row` and `column`:
 *
 * @code
 *  worksheet_write_formula(worksheet, 0, 0, "=B3 + 6",                    NULL);
 *  worksheet_write_formula(worksheet, 1, 0, "=SIN(PI()/4)",               NULL);
 *  worksheet_write_formula(worksheet, 2, 0, "=SUM(A1:A2)",                NULL);
 *  worksheet_write_formula(worksheet, 3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL);
 *  worksheet_write_formula(worksheet, 4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL);
 *  worksheet_write_formula(worksheet, 5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL);
 * @endcode
 *
 * @image html write_formula01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores a
 * default value of `0`. The correct formula result is displayed in Excel, as
 * shown in the example above, since it recalculates the formulas when it loads
 * the file. For cases where this is an issue see the
 * `worksheet_write_formula_num()` function and the discussion in that section.
 *
 * Formulas must be written with the US style separator/range operator which
 * is a comma (not semi-colon). Therefore a formula with multiple values
 * should be written as follows:
 *
 * @code
 *     // OK.
 *     worksheet_write_formula(worksheet, 0, 0, "=SUM(1, 2, 3)", NULL);
 *
 *     // NO. Error on load.
 *     worksheet_write_formula(worksheet, 1, 0, "=SUM(1; 2; 3)", NULL);
 * @endcode
 *
 * See also @ref working_with_formulas.
 */
lxw_error worksheet_write_formula(lxw_worksheet *worksheet,
                                  lxw_row_t row,
                                  lxw_col_t col, const char *formula,
                                  lxw_format *format);
/**
 * @brief Write an array formula to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 * @param formula     Array formula to write to cell.
 * @param format      A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
  * The `%worksheet_write_array_formula()` function writes an array formula to
 * a cell range. In Excel an array formula is a formula that performs a
 * calculation on a set of values.
 *
 * In Excel an array formula is indicated by a pair of braces around the
 * formula: `{=SUM(A1:B1*A2:B2)}`.
 *
 * Array formulas can return a single value or a range or values. For array
 * formulas that return a range of values you must specify the range that the
 * return values will be written to. This is why this function has `first_`
 * and `last_` row/column parameters. The RANGE() macro can also be used to
 * specify the range:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_write_array_formula(worksheet, RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL);
 * @endcode
 *
 * If the array formula returns a single value then the `first_` and `last_`
 * parameters should be the same:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL);
 *     worksheet_write_array_formula(worksheet, RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL);
 * @endcode
 *
 */
lxw_error worksheet_write_array_formula(lxw_worksheet *worksheet,
                                        lxw_row_t first_row,
                                        lxw_col_t first_col,
                                        lxw_row_t last_row,
                                        lxw_col_t last_col,
                                        const char *formula,
                                        lxw_format *format);

lxw_error worksheet_write_array_formula_num(lxw_worksheet *worksheet,
                                            lxw_row_t first_row,
                                            lxw_col_t first_col,
                                            lxw_row_t last_row,
                                            lxw_col_t last_col,
                                            const char *formula,
                                            lxw_format *format,
                                            double result);

/**
 * @brief Write a date or time to a worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param datetime  The datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * The `worksheet_write_datetime()` function can be used to write a date or
 * time to the cell specified by `row` and `column`:
 *
 * @dontinclude dates_and_times02.c
 * @skip include
 * @until num_format
 * @skip Feb
 * @until }
 *
 * The `format` parameter should be used to apply formatting to the cell using
 * a @ref format.h "Format" object as shown above. Without a date format the
 * datetime will appear as a number only.
 *
 * See @ref working_with_dates for more information about handling dates and
 * times in libxlsxwriter.
 */
lxw_error worksheet_write_datetime(lxw_worksheet *worksheet,
                                   lxw_row_t row,
                                   lxw_col_t col, lxw_datetime *datetime,
                                   lxw_format *format);

lxw_error worksheet_write_url_opt(lxw_worksheet *worksheet,
                                  lxw_row_t row_num,
                                  lxw_col_t col_num, const char *url,
                                  lxw_format *format, const char *string,
                                  const char *tooltip);
/**
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param url       The url to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 *
 * The `%worksheet_write_url()` function is used to write a URL/hyperlink to a
 * worksheet cell specified by `row` and `column`.
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", url_format);
 * @endcode
 *
 * @image html hyperlinks_short.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a @ref
 * format.h "Format" object. The typical worksheet format for a hyperlink is a
 * blue underline:
 *
 * @code
 *    lxw_format *url_format   = workbook_add_format(workbook);
 *
 *    format_set_underline (url_format, LXW_UNDERLINE_SINGLE);
 *    format_set_font_color(url_format, LXW_COLOR_BLUE);
 *
 * @endcode
 *
 * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
 * and `mailto:` :
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "ftp://www.python.org/",     url_format);
 *     worksheet_write_url(worksheet, 1, 0, "http://www.python.org/",    url_format);
 *     worksheet_write_url(worksheet, 2, 0, "https://www.python.org/",   url_format);
 *     worksheet_write_url(worksheet, 3, 0, "mailto:jmcnamara@cpan.org", url_format);
 *
 * @endcode
 *
 * An Excel hyperlink is comprised of two elements: the displayed string and
 * the non-displayed link. By default the displayed string is the same as the
 * link. However, it is possible to overwrite it with any other
 * `libxlsxwriter` type using the appropriate `worksheet_write_*()`
 * function. The most common case is to overwrite the displayed link text with
 * another string:
 *
 * @code
 *  // Write a hyperlink but overwrite the displayed string.
 *  worksheet_write_url   (worksheet, 2, 0, "http://libxlsxwriter.github.io", url_format);
 *  worksheet_write_string(worksheet, 2, 0, "Read the documentation.",        url_format);
 *
 * @endcode
 *
 * @image html hyperlinks_short2.png
 *
 * Two local URIs are supported: `internal:` and `external:`. These are used
 * for hyperlinks to internal worksheet references or external workbook and
 * worksheet references:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:Sheet2!A1",                url_format);
 *     worksheet_write_url(worksheet, 1, 0, "internal:Sheet2!B2",                url_format);
 *     worksheet_write_url(worksheet, 2, 0, "internal:Sheet2!A1:B2",             url_format);
 *     worksheet_write_url(worksheet, 3, 0, "internal:'Sales Data'!A1",          url_format);
 *     worksheet_write_url(worksheet, 4, 0, "external:c:\\temp\\foo.xlsx",       url_format);
 *     worksheet_write_url(worksheet, 5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 6, 0, "external:..\\foo.xlsx",             url_format);
 *     worksheet_write_url(worksheet, 7, 0, "external:..\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 8, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 *
 * @endcode
 *
 * Worksheet references are typically of the form `Sheet1!A1`. You can also
 * link to a worksheet range using the standard Excel notation:
 * `Sheet1!A1:B2`.
 *
 * In external links the workbook and worksheet name must be separated by the
 * `#` character:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 * @endcode
 *
 * You can also link to a named range in the target worksheet: For example say
 * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
 * you could link to it as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\temp\\foo.xlsx#my_name", url_format);
 *
 * @endcode
 *
 * Excel requires that worksheet names containing spaces or non alphanumeric
 * characters are single quoted as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:'Sales Data'!A1", url_format);
 * @endcode
 *
 * Links to network files are also supported. Network files normally begin
 * with two back slashes as follows `\\NETWORK\etc`. In order to represent
 * this in a C string literal the backslashes should be escaped:
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 * @endcode
 *
 *
 * Alternatively, you can use Unix style forward slashes. These are
 * translated internally to backslashes:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:/temp/foo.xlsx",     url_format);
 *     worksheet_write_url(worksheet, 1, 0, "external://NET/share/foo.xlsx", url_format);
 *
 * @endcode
 *
 *
 * **Note:**
 *
 *    libxlsxwriter will escape the following characters in URLs as required
 *    by Excel: `\s " < > \ [ ]  ^ { }` unless the URL already contains `%%xx`
 *    style escapes. In which case it is assumed that the URL was escaped
 *    correctly by the user and will by passed directly to Excel.
 *
 */
lxw_error worksheet_write_url(lxw_worksheet *worksheet,
                              lxw_row_t row,
                              lxw_col_t col, const char *url,
                              lxw_format *format);

/**
 * @brief Write a formatted boolean worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param value     The boolean value to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * Write an Excel boolean to the cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_boolean(worksheet, 2, 2, 0, my_format);
 * @endcode
 *
 */
lxw_error worksheet_write_boolean(lxw_worksheet *worksheet,
                                  lxw_row_t row, lxw_col_t col,
                                  int value, lxw_format *format);

/**
 * @brief Write a formatted blank worksheet cell.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * Write a blank cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_blank(worksheet, 1, 1, border_format);
 * @endcode
 *
 * This function is used to add formatting to a cell which doesn't contain a
 * string or number value.
 *
 * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
 * cell is a cell which doesn't contain data or formatting whilst a Blank cell
 * doesn't contain data but does contain formatting. Excel stores Blank cells
 * but ignores Empty cells.
 *
 * As such, if you write an empty cell without formatting it is ignored.
 *
 */
lxw_error worksheet_write_blank(lxw_worksheet *worksheet,
                                lxw_row_t row, lxw_col_t col,
                                lxw_format *format);

/**
 * @brief Write a formula to a worksheet cell with a user defined result.
 *
 * @param worksheet pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined result for a formula.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_write_formula_num()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * result:
 *
 * @code
 *     // Required as a workaround only.
 *     worksheet_write_formula_num(worksheet, 0, 0, "=1 + 2", NULL, 3);
 * @endcode
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores
 * the value `0` as the formula result. It then sets a global flag in the XLSX
 * file to say that all formulas and functions should be recalculated when the
 * file is opened.
 *
 * This is the method recommended in the Excel documentation and in general it
 * works fine with spreadsheet applications.
 *
 * However, applications that don't have a facility to calculate formulas,
 * such as Excel Viewer, or some mobile applications will only display the `0`
 * results.
 *
 * If required, the `%worksheet_write_formula_num()` function can be used to
 * specify a formula and its result.
 *
 * This function is rarely required and is only provided for compatibility
 * with some third party applications. For most applications the
 * worksheet_write_formula() function is the recommended way of writing
 * formulas.
 *
 * See also @ref working_with_formulas.
 */
lxw_error worksheet_write_formula_num(lxw_worksheet *worksheet,
                                      lxw_row_t row,
                                      lxw_col_t col,
                                      const char *formula,
                                      lxw_format *format, double result);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 *
 * The `%worksheet_set_row()` function is used to change the default
 * properties of a row. The most common use for this function is to change the
 * height of a row:
 *
 * @code
 *     // Set the height of Row 1 to 20.
 *     worksheet_set_row(worksheet, 0, 20, NULL);
 * @endcode
 *
 * The other common use for `%worksheet_set_row()` is to set the a @ref
 * format.h "Format" for all cells in the row:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the header row to bold.
 *     worksheet_set_row(worksheet, 0, 15, bold);
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height of #LXW_DEF_ROW_HEIGHT = 15:
 *
 * @code
 *     worksheet_set_row(worksheet, 0, LXW_DEF_ROW_HEIGHT, format);
 *     worksheet_set_row(worksheet, 0, 15, format); // Same as above.
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the row that don't
 * have a format. As with Excel the row format is overridden by an explicit
 * cell format. For example:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1);
 *
 *     // Cell A1 in Row 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell B1 in Row 1 keeps format2.
 *     worksheet_write_string(worksheet, 0, 1, "Hello", format2);
 * @endcode
 *
 */
lxw_error worksheet_set_row(lxw_worksheet *worksheet,
                            lxw_row_t row, double height, lxw_format *format);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_row_opt()` function  is the same as
 *  `worksheet_set_row()` with an additional `options` parameter.
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a row. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options1 = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     // Hide the fourth and fifth (zero indexed) rows.
 *     worksheet_set_row_opt(worksheet, 3,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 4,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *
 * @endcode
 *
 * @image html hide_row_col2.png
 *
 * The `"hidden"`, `"level"`,  and `"collapsed"`, options can also be used to
 * create Outlines and Grouping. See @ref working_with_outlines.
 *
 * @code
 *     // The option structs with the outline level set.
 *     lxw_row_col_options options1 = {.hidden = 0, .level = 2, .collapsed = 0};
 *     lxw_row_col_options options2 = {.hidden = 0, .level = 1, .collapsed = 0};
 *
 *
 *     // Set the row options with the outline level.
 *     worksheet_set_row_opt(worksheet, 1,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 2,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 3,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 4,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 5,  LXW_DEF_ROW_HEIGHT, NULL, &options2);
 *
 *     worksheet_set_row_opt(worksheet, 6,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 7,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 8,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 9,  LXW_DEF_ROW_HEIGHT, NULL, &options1);
 *     worksheet_set_row_opt(worksheet, 10, LXW_DEF_ROW_HEIGHT, NULL, &options2);
 * @endcode
 *
 * @image html outline1.png
 *
 */
lxw_error worksheet_set_row_opt(lxw_worksheet *worksheet,
                                lxw_row_t row,
                                double height,
                                lxw_format *format,
                                lxw_row_col_options *options);

/**
 * @brief Set the properties for one or more columns of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 *
 * The `%worksheet_set_column()` function can be used to change the default
 * properties of a single column or a range of columns:
 *
 * @code
 *     // Width of columns B:D set to 30.
 *     worksheet_set_column(worksheet, 1, 3, 30, NULL);
 *
 * @endcode
 *
 * If `%worksheet_set_column()` is applied to a single column the value of
 * `first_col` and `last_col` should be the same:
 *
 * @code
 *     // Width of column B set to 30.
 *     worksheet_set_column(worksheet, 1, 1, 30, NULL);
 *
 * @endcode
 *
 * It is also possible, and generally clearer, to specify a column range using
 * the form of `COLS()` macro:
 *
 * @code
 *     worksheet_set_column(worksheet, 4, 4, 20, NULL);
 *     worksheet_set_column(worksheet, 5, 8, 30, NULL);
 *
 *     // Same as the examples above but clearer.
 *     worksheet_set_column(worksheet, COLS("E:E"), 20, NULL);
 *     worksheet_set_column(worksheet, COLS("F:H"), 30, NULL);
 *
 * @endcode
 *
 * The `width` parameter sets the column width in the same units used by Excel
 * which is: the number of characters in the default font. The default width
 * is 8.43 in the default font of Calibri 11. The actual relationship between
 * a string width and a column width in Excel is complex. See the
 * [following explanation of column widths](https://support.microsoft.com/en-us/kb/214123)
 * from the Microsoft support documentation for more details.
 *
 * There is no way to specify "AutoFit" for a column in the Excel file
 * format. This feature is only available at runtime from within Excel. It is
 * possible to simulate "AutoFit" in your application by tracking the maximum
 * width of the data in the column as your write it and then adjusting the
 * column width at the end.
 *
 * As usual the @ref format.h `format` parameter is optional. If you wish to
 * set the format without changing the width you can pass a default column
 * width of #LXW_DEF_COL_WIDTH = 8.43:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the first column to bold.
 *     worksheet_set_column(worksheet, 0, 0, LXW_DEF_COL_HEIGHT, bold);
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the column that
 * don't have a format. For example:
 *
 * @code
 *     // Column 1 has format1.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format1);
 *
 *     // Cell A1 in column 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell A2 in column 1 keeps format2.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", format2);
 * @endcode
 *
 * As in Excel a row format takes precedence over a default column format:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1);
 *
 *     // Col 1 has format2.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format2);
 *
 *     // Cell A1 defaults to format1, the row format.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *    // Cell A2 keeps format2, the column format.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", NULL);
 * @endcode
 */
lxw_error worksheet_set_column(lxw_worksheet *worksheet,
                               lxw_col_t first_col,
                               lxw_col_t last_col,
                               double width, lxw_format *format);

/**
 * @brief Set the properties for one or more columns of cells with options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_column_opt()` function  is the same as
 * `worksheet_set_column()` with an additional `options` parameter.
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a column. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options1 = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     worksheet_set_column_opt(worksheet, COLS("D:E"),  LXW_DEF_COL_WIDTH, NULL, &options1);
 * @endcode
 *
 * @image html hide_row_col3.png
 *
 * The `"hidden"`, `"level"`,  and `"collapsed"`, options can also be used to
 * create Outlines and Grouping. See @ref working_with_outlines.
 *
 * @code
 *     lxw_row_col_options options1 = {.hidden = 0, .level = 1, .collapsed = 0};
 *
 *     worksheet_set_column_opt(worksheet, COLS("B:G"),  5, NULL, &options1);
 * @endcode
 *
 * @image html outline8.png
 */
lxw_error worksheet_set_column_opt(lxw_worksheet *worksheet,
                                   lxw_col_t first_col,
                                   lxw_col_t last_col,
                                   double width,
                                   lxw_format *format,
                                   lxw_row_col_options *options);

/**
 * @brief Insert an image in a worksheet cell.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 *
 * @return A #lxw_error code.
 *
 * This function can be used to insert a image into a worksheet. The image can
 * be in PNG, JPEG or BMP format:
 *
 * @code
 *     worksheet_insert_image(worksheet, 2, 1, "logo.png");
 * @endcode
 *
 * @image html insert_image.png
 *
 * The `worksheet_insert_image_opt()` function takes additional optional
 * parameters to position and scale the image, see below.
 *
 * **Note**:
 * The scaling of a image may be affected if is crosses a row that has its
 * default height changed due to a font that is larger than the default font
 * size or that has text wrapping turned on. To avoid this you should
 * explicitly set the height of the row using `worksheet_set_row()` if it
 * crosses an inserted image.
 *
 * BMP images are only supported for backward compatibility. In general it is
 * best to avoid BMP images since they aren't compressed. If used, BMP images
 * must be 24 bit, true color, bitmaps.
 */
lxw_error worksheet_insert_image(lxw_worksheet *worksheet,
                                 lxw_row_t row, lxw_col_t col,
                                 const char *filename);

/**
 * @brief Insert an image in a worksheet cell, with options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 * @param options   Optional image parameters.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_insert_image_opt()` function is like
 * `worksheet_insert_image()` function except that it takes an optional
 * #lxw_image_options struct to scale and position the image:
 *
 * @code
 *    lxw_image_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.5};
 *
 *    worksheet_insert_image_opt(worksheet, 2, 1, "logo.png", &options);
 *
 * @endcode
 *
 * @image html insert_image_opt.png
 *
 * @note See the notes about row scaling and BMP images in
 * `worksheet_insert_image()` above.
 */
lxw_error worksheet_insert_image_opt(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     const char *filename,
                                     lxw_image_options *options);
/**
 * @brief Insert a chart object into a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param chart     A #lxw_chart object created via workbook_add_chart().
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_insert_chart()` can be used to insert a chart into a
 * worksheet. The chart object must be created first using the
 * `workbook_add_chart()` function and configured using the @ref chart.h
 * functions.
 *
 * @code
 *     // Create a chart object.
 *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_LINE);
 *
 *     // Add a data series to the chart.
 *     chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$6");
 *
 *     // Insert the chart into the worksheet
 *     worksheet_insert_chart(worksheet, 0, 2, chart);
 * @endcode
 *
 * @image html chart_working.png
 *
 *
 * **Note:**
 *
 * A chart may only be inserted into a worksheet once. If several similar
 * charts are required then each one must be created separately with
 * `%worksheet_insert_chart()`.
 *
 */
lxw_error worksheet_insert_chart(lxw_worksheet *worksheet,
                                 lxw_row_t row, lxw_col_t col,
                                 lxw_chart *chart);

/**
 * @brief Insert a chart object into a worksheet, with options.
 *
 * @param worksheet    Pointer to a lxw_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param chart        A #lxw_chart object created via workbook_add_chart().
 * @param user_options Optional chart parameters.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_insert_chart_opt()` function is like
 * `worksheet_insert_chart()` function except that it takes an optional
 * #lxw_image_options struct to scale and position the image of the chart:
 *
 * @code
 *    lxw_image_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.75};
 *
 *    worksheet_insert_chart_opt(worksheet, 0, 2, chart, &options);
 *
 * @endcode
 *
 * @image html chart_line_opt.png
 *
 * The #lxw_image_options struct is the same struct used in
 * `worksheet_insert_image_opt()` to position and scale images.
 *
 */
lxw_error worksheet_insert_chart_opt(lxw_worksheet *worksheet,
                                     lxw_row_t row, lxw_col_t col,
                                     lxw_chart *chart,
                                     lxw_image_options *user_options);

/**
 * @brief Merge a range of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param string    String to write to the merged range.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_merge_range()` function allows cells to be merged together
 * so that they act as a single area.
 *
 * Excel generally merges and centers cells at same time. To get similar
 * behavior with libxlsxwriter you need to apply a @ref format.h "Format"
 * object with the appropriate alignment:
 *
 * @code
 *     lxw_format *merge_format = workbook_add_format(workbook);
 *     format_set_align(merge_format, LXW_ALIGN_CENTER);
 *
 *     worksheet_merge_range(worksheet, 1, 1, 1, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * It is possible to apply other formatting to the merged cells as well:
 *
 * @code
 *    format_set_align   (merge_format, LXW_ALIGN_CENTER);
 *    format_set_align   (merge_format, LXW_ALIGN_VERTICAL_CENTER);
 *    format_set_border  (merge_format, LXW_BORDER_DOUBLE);
 *    format_set_bold    (merge_format);
 *    format_set_bg_color(merge_format, 0xD7E4BC);
 *
 *    worksheet_merge_range(worksheet, 2, 1, 3, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * @image html merge.png
 *
 * The `%worksheet_merge_range()` function writes a `char*` string using
 * `worksheet_write_string()`. In order to write other data types, such as a
 * number or a formula, you can overwrite the first cell with a call to one of
 * the other write functions. The same Format should be used as was used in
 * the merged range.
 *
 * @code
 *    // First write a range with a blank string.
 *    worksheet_merge_range (worksheet, 1, 1, 1, 3, "", format);
 *
 *    // Then overwrite the first cell with a number.
 *    worksheet_write_number(worksheet, 1, 1, 123, format);
 * @endcode
 *
 * @note Merged ranges generally don’t work in libxlsxwriter when the Workbook
 * #lxw_workbook_options `constant_memory` mode is enabled.
 */
lxw_error worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row,
                                lxw_col_t first_col, lxw_row_t last_row,
                                lxw_col_t last_col, const char *string,
                                lxw_format *format);

/**
 * @brief Set the autofilter area in the worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_autofilter()` function allows an autofilter to be added to
 * a worksheet.
 *
 * An autofilter is a way of adding drop down lists to the headers of a 2D
 * range of worksheet data. This allows users to filter the data based on
 * simple criteria so that some data is shown and some is hidden.
 *
 * @image html autofilter.png
 *
 * To add an autofilter to a worksheet:
 *
 * @code
 *     worksheet_autofilter(worksheet, 0, 0, 50, 3);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_autofilter(worksheet, RANGE("A1:D51"));
 * @endcode
 *
 * Note: it isn't currently possible to apply filter conditions to the
 * autofilter.
 */
lxw_error worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row,
                               lxw_col_t first_col, lxw_row_t last_row,
                               lxw_col_t last_col);

/**
 * @brief Add a data validation to a cell.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param row        The zero indexed row number.
 * @param col        The zero indexed column number.
 * @param validation A #lxw_data_validation object to control the validation.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_data_validation_cell()` function is used to construct an
 * Excel data validation or to limit the user input to a dropdown list of
 * values:
 *
 * @code
 *
 *    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
 *
 *    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER;
 *    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN;
 *    data_validation->minimum_number = 1;
 *    data_validation->maximum_number = 10;
 *
 *    worksheet_data_validation_cell(worksheet, 2, 1, data_validation);
 *
 *    // Same as above with the CELL() macro.
 *    worksheet_data_validation_cell(worksheet, CELL("B3"), data_validation);
 *
 * @endcode
 *
 * @image html data_validate4.png
 *
 * Data validation and the various options of #lxw_data_validation are
 * described in more detail in @ref working_with_data_validation.
 */
lxw_error worksheet_data_validation_cell(lxw_worksheet *worksheet,
                                         lxw_row_t row, lxw_col_t col,
                                         lxw_data_validation *validation);

/**
 * @brief Add a data validation to a range cell.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param first_row  The first row of the range. (All zero indexed.)
 * @param first_col  The first column of the range.
 * @param last_row   The last row of the range.
 * @param last_col   The last col of the range.
 * @param validation A #lxw_data_validation object to control the validation.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_data_validation_range()` function is the same as the
 * `%worksheet_data_validation_cell()`, see above,  except the data validation
 * is applied to a range of cells:
 *
 * @code
 *
 *    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
 *
 *    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER;
 *    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN;
 *    data_validation->minimum_number = 1;
 *    data_validation->maximum_number = 10;
 *
 *    worksheet_data_validation_range(worksheet, 2, 1, 4, 1, data_validation);
 *
 *    // Same as above with the RANGE() macro.
 *    worksheet_data_validation_range(worksheet, RANGE("B3:B5"), data_validation);
 *
 * @endcode
 *
 * Data validation and the various options of #lxw_data_validation are
 * described in more detail in @ref working_with_data_validation.
 */
lxw_error worksheet_data_validation_range(lxw_worksheet *worksheet,
                                          lxw_row_t first_row,
                                          lxw_col_t first_col,
                                          lxw_row_t last_row,
                                          lxw_col_t last_col,
                                          lxw_data_validation *validation);

 /**
  * @brief Make a worksheet the active, i.e., visible worksheet.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_activate()` function is used to specify which worksheet is
  * initially visible in a multi-sheet workbook:
  *
  * @code
  *     lxw_worksheet *worksheet1 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet2 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet3 = workbook_add_worksheet(workbook, NULL);
  *
  *     worksheet_activate(worksheet3);
  * @endcode
  *
  * @image html worksheet_activate.png
  *
  * More than one worksheet can be selected via the `worksheet_select()`
  * function, see below, however only one worksheet can be active.
  *
  * The default active worksheet is the first worksheet.
  *
  */
void worksheet_activate(lxw_worksheet *worksheet);

 /**
  * @brief Set a worksheet tab as selected.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_select()` function is used to indicate that a worksheet is
  * selected in a multi-sheet workbook:
  *
  * @code
  *     worksheet_activate(worksheet1);
  *     worksheet_select(worksheet2);
  *     worksheet_select(worksheet3);
  *
  * @endcode
  *
  * A selected worksheet has its tab highlighted. Selecting worksheets is a
  * way of grouping them together so that, for example, several worksheets
  * could be printed in one go. A worksheet that has been activated via the
  * `worksheet_activate()` function will also appear as selected.
  *
  */
void worksheet_select(lxw_worksheet *worksheet);

/**
 * @brief Hide the current worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_hide()` function is used to hide a worksheet:
 *
 * @code
 *     worksheet_hide(worksheet2);
 * @endcode
 *
 * You may wish to hide a worksheet in order to avoid confusing a user with
 * intermediate data or calculations.
 *
 * @image html hide_sheet.png
 *
 * A hidden worksheet can not be activated or selected so this function is
 * mutually exclusive with the `worksheet_activate()` and `worksheet_select()`
 * functions. In addition, since the first worksheet will default to being the
 * active worksheet, you cannot hide the first worksheet without activating
 * another sheet:
 *
 * @code
 *     worksheet_activate(worksheet2);
 *     worksheet_hide(worksheet1);
 * @endcode
 */
void worksheet_hide(lxw_worksheet *worksheet);

/**
 * @brief Set current worksheet as the first visible sheet tab.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `worksheet_activate()` function determines which worksheet is initially
 * selected.  However, if there are a large number of worksheets the selected
 * worksheet may not appear on the screen. To avoid this you can select the
 * leftmost visible worksheet tab using `%worksheet_set_first_sheet()`:
 *
 * @code
 *     worksheet_set_first_sheet(worksheet19); // First visible worksheet tab.
 *     worksheet_activate(worksheet20);        // First visible worksheet.
 * @endcode
 *
 * This function is not required very often. The default value is the first
 * worksheet.
 */
void worksheet_set_first_sheet(lxw_worksheet *worksheet);

/**
 * @brief Split and freeze a worksheet into panes.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The cell row (zero indexed).
 * @param col       The cell column (zero indexed).
 *
 * The `%worksheet_freeze_panes()` function can be used to divide a worksheet
 * into horizontal or vertical regions known as panes and to "freeze" these
 * panes so that the splitter bars are not visible.
 *
 * The parameters `row` and `col` are used to specify the location of the
 * split. It should be noted that the split is specified at the top or left of
 * a cell and that the function uses zero based indexing. Therefore to freeze
 * the first row of a worksheet it is necessary to specify the split at row 2
 * (which is 1 as the zero-based index).
 *
 * You can set one of the `row` and `col` parameters as zero if you do not
 * want either a vertical or horizontal split.
 *
 * Examples:
 *
 * @code
 *     worksheet_freeze_panes(worksheet1, 1, 0); // Freeze the first row.
 *     worksheet_freeze_panes(worksheet2, 0, 1); // Freeze the first column.
 *     worksheet_freeze_panes(worksheet3, 1, 1); // Freeze first row/column.
 *
 * @endcode
 *
 */
void worksheet_freeze_panes(lxw_worksheet *worksheet,
                            lxw_row_t row, lxw_col_t col);
/**
 * @brief Split a worksheet into panes.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param vertical   The position for the vertical split.
 * @param horizontal The position for the horizontal split.
 *
 * The `%worksheet_split_panes()` function can be used to divide a worksheet
 * into horizontal or vertical regions known as panes. This function is
 * different from the `worksheet_freeze_panes()` function in that the splits
 * between the panes will be visible to the user and each pane will have its
 * own scroll bars.
 *
 * The parameters `vertical` and `horizontal` are used to specify the vertical
 * and horizontal position of the split. The units for `vertical` and
 * `horizontal` are the same as those used by Excel to specify row height and
 * column width. However, the vertical and horizontal units are different from
 * each other. Therefore you must specify the `vertical` and `horizontal`
 * parameters in terms of the row heights and column widths that you have set
 * or the default values which are 15 for a row and 8.43 for a column.
 *
 * Examples:
 *
 * @code
 *     worksheet_split_panes(worksheet1, 15, 0);    // First row.
 *     worksheet_split_panes(worksheet2, 0,  8.43); // First column.
 *     worksheet_split_panes(worksheet3, 15, 8.43); // First row and column.
 *
 * @endcode
 *
 */
void worksheet_split_panes(lxw_worksheet *worksheet,
                           double vertical, double horizontal);

/* worksheet_freeze_panes() with infrequent options. Undocumented for now. */
void worksheet_freeze_panes_opt(lxw_worksheet *worksheet,
                                lxw_row_t first_row, lxw_col_t first_col,
                                lxw_row_t top_row, lxw_col_t left_col,
                                uint8_t type);

/* worksheet_split_panes() with infrequent options. Undocumented for now. */
void worksheet_split_panes_opt(lxw_worksheet *worksheet,
                               double vertical, double horizontal,
                               lxw_row_t top_row, lxw_col_t left_col);
/**
 * @brief Set the selected cell or cells in a worksheet:
 *
 * @param worksheet   Pointer to a lxw_worksheet instance to be updated.
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 *
 *
 * The `%worksheet_set_selection()` function can be used to specify which cell
 * or range of cells is selected in a worksheet: The most common requirement
 * is to select a single cell, in which case the `first_` and `last_`
 * parameters should be the same.
 *
 * The active cell within a selected range is determined by the order in which
 * `first_` and `last_` are specified.
 *
 * Examples:
 *
 * @code
 *     worksheet_set_selection(worksheet1, 3, 3, 3, 3);     // Cell D4.
 *     worksheet_set_selection(worksheet2, 3, 3, 6, 6);     // Cells D4 to G7.
 *     worksheet_set_selection(worksheet3, 6, 6, 3, 3);     // Cells G7 to D4.
 *     worksheet_set_selection(worksheet5, RANGE("D4:G7")); // Using the RANGE macro.
 *
 * @endcode
 *
 */
void worksheet_set_selection(lxw_worksheet *worksheet,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);

/**
 * @brief Set the page orientation as landscape.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to landscape:
 *
 * @code
 *     worksheet_set_landscape(worksheet);
 * @endcode
 */
void worksheet_set_landscape(lxw_worksheet *worksheet);

/**
 * @brief Set the page orientation as portrait.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to portrait. The default worksheet orientation is portrait, so this
 * function isn't generally required:
 *
 * @code
 *     worksheet_set_portrait(worksheet);
 * @endcode
 */
void worksheet_set_portrait(lxw_worksheet *worksheet);

/**
 * @brief Set the page layout to page view mode.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to display the worksheet in "Page View/Layout" mode:
 *
 * @code
 *     worksheet_set_page_view(worksheet);
 * @endcode
 */
void worksheet_set_page_view(lxw_worksheet *worksheet);

/**
 * @brief Set the paper type for printing.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param paper_type The Excel paper format type.
 *
 * This function is used to set the paper format for the printed output of a
 * worksheet. The following paper styles are available:
 *
 *
 *   Index    | Paper format            | Paper size
 *   :------- | :---------------------- | :-------------------
 *   0        | Printer default         | Printer default
 *   1        | Letter                  | 8 1/2 x 11 in
 *   2        | Letter Small            | 8 1/2 x 11 in
 *   3        | Tabloid                 | 11 x 17 in
 *   4        | Ledger                  | 17 x 11 in
 *   5        | Legal                   | 8 1/2 x 14 in
 *   6        | Statement               | 5 1/2 x 8 1/2 in
 *   7        | Executive               | 7 1/4 x 10 1/2 in
 *   8        | A3                      | 297 x 420 mm
 *   9        | A4                      | 210 x 297 mm
 *   10       | A4 Small                | 210 x 297 mm
 *   11       | A5                      | 148 x 210 mm
 *   12       | B4                      | 250 x 354 mm
 *   13       | B5                      | 182 x 257 mm
 *   14       | Folio                   | 8 1/2 x 13 in
 *   15       | Quarto                  | 215 x 275 mm
 *   16       | ---                     | 10x14 in
 *   17       | ---                     | 11x17 in
 *   18       | Note                    | 8 1/2 x 11 in
 *   19       | Envelope 9              | 3 7/8 x 8 7/8
 *   20       | Envelope 10             | 4 1/8 x 9 1/2
 *   21       | Envelope 11             | 4 1/2 x 10 3/8
 *   22       | Envelope 12             | 4 3/4 x 11
 *   23       | Envelope 14             | 5 x 11 1/2
 *   24       | C size sheet            | ---
 *   25       | D size sheet            | ---
 *   26       | E size sheet            | ---
 *   27       | Envelope DL             | 110 x 220 mm
 *   28       | Envelope C3             | 324 x 458 mm
 *   29       | Envelope C4             | 229 x 324 mm
 *   30       | Envelope C5             | 162 x 229 mm
 *   31       | Envelope C6             | 114 x 162 mm
 *   32       | Envelope C65            | 114 x 229 mm
 *   33       | Envelope B4             | 250 x 353 mm
 *   34       | Envelope B5             | 176 x 250 mm
 *   35       | Envelope B6             | 176 x 125 mm
 *   36       | Envelope                | 110 x 230 mm
 *   37       | Monarch                 | 3.875 x 7.5 in
 *   38       | Envelope                | 3 5/8 x 6 1/2 in
 *   39       | Fanfold                 | 14 7/8 x 11 in
 *   40       | German Std Fanfold      | 8 1/2 x 12 in
 *   41       | German Legal Fanfold    | 8 1/2 x 13 in
 *
 * Note, it is likely that not all of these paper types will be available to
 * the end user since it will depend on the paper formats that the user's
 * printer supports. Therefore, it is best to stick to standard paper types:
 *
 * @code
 *     worksheet_set_paper(worksheet1, 1);  // US Letter
 *     worksheet_set_paper(worksheet2, 9);  // A4
 * @endcode
 *
 * If you do not specify a paper type the worksheet will print using the
 * printer's default paper style.
 */
void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);

/**
 * @brief Set the worksheet margins for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param left    Left margin in inches.   Excel default is 0.7.
 * @param right   Right margin in inches.  Excel default is 0.7.
 * @param top     Top margin in inches.    Excel default is 0.75.
 * @param bottom  Bottom margin in inches. Excel default is 0.75.
 *
 * The `%worksheet_set_margins()` function is used to set the margins of the
 * worksheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    worksheet_set_margins(worksheet, 1.3, 1.2, -1, -1);
 * @endcode
 *
 */
void worksheet_set_margins(lxw_worksheet *worksheet, double left,
                           double right, double top, double bottom);

/**
 * @brief Set the printed page header caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 *
 * @return A #lxw_error code.
 *
 * Headers and footers are generated using a string which is a combination of
 * plain text and control characters.
 *
 * The available control character are:
 *
 *
 *   | Control         | Category      | Description           |
 *   | --------------- | ------------- | --------------------- |
 *   | `&L`            | Justification | Left                  |
 *   | `&C`            |               | Center                |
 *   | `&R`            |               | Right                 |
 *   | `&P`            | Information   | Page number           |
 *   | `&N`            |               | Total number of pages |
 *   | `&D`            |               | Date                  |
 *   | `&T`            |               | Time                  |
 *   | `&F`            |               | File name             |
 *   | `&A`            |               | Worksheet name        |
 *   | `&Z`            |               | Workbook path         |
 *   | `&fontsize`     | Font          | Font size             |
 *   | `&"font,style"` |               | Font name and style   |
 *   | `&U`            |               | Single underline      |
 *   | `&E`            |               | Double underline      |
 *   | `&S`            |               | Strikethrough         |
 *   | `&X`            |               | Superscript           |
 *   | `&Y`            |               | Subscript             |
 *
 *
 * Text in headers and footers can be justified (aligned) to the left, center
 * and right by prefixing the text with the control characters `&L`, `&C` and
 * `&R`.
 *
 * For example (with ASCII art representation of the results):
 *
 * @code
 *     worksheet_set_header(worksheet, "&LHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Hello                                                         |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&CHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&RHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                                                         Hello |
 *     |                                                               |
 *
 *
 * @endcode
 *
 * For simple text, if you do not specify any justification the text will be
 * centered. However, you must prefix the text with `&C` if you specify a font
 * name or any other formatting:
 *
 * @code
 *     worksheet_set_header(worksheet, "Hello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 * @endcode
 *
 * You can have text in each of the justification regions:
 *
 * @code
 *     worksheet_set_header(worksheet, "&LCiao&CBello&RCielo");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Ciao                     Bello                          Cielo |
 *     |                                                               |
 *
 * @endcode
 *
 * The information control characters act as variables that Excel will update
 * as the workbook or worksheet changes. Times and dates are in the users
 * default format:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CPage &P of &N");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                        Page 1 of 6                            |
 *     |                                                               |
 *
 *     worksheet_set_header(worksheet, "&CUpdated at &T");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                    Updated at 12:30 PM                        |
 *     |                                                               |
 *
 * @endcode
 *
 * You can specify the font size of a section of the text by prefixing it with
 * the control character `&n` where `n` is the font size:
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&30Hello Big");
 *     worksheet_set_header(worksheet2, "&C&10Hello Small");
 *
 * @endcode
 *
 * You can specify the font of a section of the text by prefixing it with the
 * control sequence `&"font,style"` where `fontname` is a font name such as
 * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
 * "Courier New" or "Times New Roman" and `style` is one of the standard
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&\"Courier New,Italic\"Hello");
 *     worksheet_set_header(worksheet2, "&C&\"Courier New,Bold Italic\"Hello");
 *     worksheet_set_header(worksheet3, "&C&\"Times New Roman,Regular\"Hello");
 *
 * @endcode
 *
 * It is possible to combine all of these features together to create
 * sophisticated headers and footers. As an aid to setting up complicated
 * headers and footers you can record a page set-up as a macro in Excel and
 * look at the format strings that VBA produces. Remember however that VBA
 * uses two double quotes `""` to indicate a single double quote. For the last
 * example above the equivalent VBA code looks like this:
 *
 * @code
 *     .LeftHeader = ""
 *     .CenterHeader = "&""Times New Roman,Regular""Hello"
 *     .RightHeader = ""
 *
 * @endcode
 *
 * Alternatively you can inspect the header and footer strings in an Excel
 * file by unzipping it and grepping the XML sub-files. The following shows
 * how to do that using libxml's xmllint to format the XML for clarity:
 *
 * @code
 *
 *    $ unzip myfile.xlsm -d myfile
 *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer"
 *
 *      <headerFooter scaleWithDoc="0">
 *        <oddHeader>&amp;L&amp;P</oddHeader>
 *      </headerFooter>
 *
 * @endcode
 *
 * Note that in this case you need to unescape the Html. In the above example
 * the header string would be `&L&P`.
 *
 * To include a single literal ampersand `&` in a header or footer you should
 * use a double ampersand `&&`:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CCuriouser && Curiouser - Attorneys at Law");
 * @endcode
 *
 * Note, the header or footer string must be less than 255 characters. Strings
 * longer than this will not be written.
 *
 */
lxw_error worksheet_set_header(lxw_worksheet *worksheet, const char *string);

/**
 * @brief Set the printed page footer caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as worksheet_set_header().
 *
 */
lxw_error worksheet_set_footer(lxw_worksheet *worksheet, const char *string);

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 * @param options   Header options.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as worksheet_set_header() with an
 * additional parameter to specify options for the header.
 *
 * Currently, the only available option is the header margin:
 *
 * @code
 *
 *    lxw_header_footer_options header_options = { 0.2 };
 *
 *    worksheet_set_header_opt(worksheet, "Some text", &header_options);
 *
 * @endcode
 *
 */
lxw_error worksheet_set_header_opt(lxw_worksheet *worksheet,
                                   const char *string,
                                   lxw_header_footer_options *options);

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 * @param options   Footer options.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as worksheet_set_header_opt().
 *
 */
lxw_error worksheet_set_footer_opt(lxw_worksheet *worksheet,
                                   const char *string,
                                   lxw_header_footer_options *options);

/**
 * @brief Set the horizontal page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_set_h_pagebreaks()` function adds horizontal page breaks to
 * a worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Horizontal page breaks act between rows.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_row_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_row_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_row_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_h_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between rows 20 and 21 you must specify the break at
 * row 21. However in zero index notation this is actually row 20:
 *
 * @code
 *    // Break between row 20 and 21.
 *    lxw_row_t breaks[] = {20, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
lxw_error worksheet_set_h_pagebreaks(lxw_worksheet *worksheet,
                                     lxw_row_t breaks[]);

/**
 * @brief Set the vertical page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #lxw_error code.
 *
 * The `%worksheet_set_v_pagebreaks()` function adds vertical page breaks to a
 * worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Vertical page breaks act between columns.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_col_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_col_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_col_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_v_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between columns 20 and 21 you must specify the break
 * at column 21. However in zero index notation this is actually column 20:
 *
 * @code
 *    // Break between column 20 and 21.
 *    lxw_col_t breaks[] = {20, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 vertical page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
lxw_error worksheet_set_v_pagebreaks(lxw_worksheet *worksheet,
                                     lxw_col_t breaks[]);

/**
 * @brief Set the order in which pages are printed.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_print_across()` function is used to change the default
 * print direction. This is referred to by Excel as the sheet "page order":
 *
 * @code
 *     worksheet_print_across(worksheet);
 * @endcode
 *
 * The default page order is shown below for a worksheet that extends over 4
 * pages. The order is called "down then across":
 *
 *     [1] [3]
 *     [2] [4]
 *
 * However, by using the `print_across` function the print order will be
 * changed to "across then down":
 *
 *     [1] [2]
 *     [3] [4]
 *
 */
void worksheet_print_across(lxw_worksheet *worksheet);

/**
 * @brief Set the worksheet zoom factor.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Worksheet zoom factor.
 *
 * Set the worksheet zoom factor in the range `10 <= zoom <= 400`:
 *
 * @code
 *     worksheet_set_zoom(worksheet1, 50);
 *     worksheet_set_zoom(worksheet2, 75);
 *     worksheet_set_zoom(worksheet3, 300);
 *     worksheet_set_zoom(worksheet4, 400);
 * @endcode
 *
 * The default zoom factor is 100. It isn't possible to set the zoom to
 * "Selection" because it is calculated by Excel at run-time.
 *
 * Note, `%worksheet_zoom()` does not affect the scale of the printed
 * page. For that you should use `worksheet_set_print_scale()`.
 */
void worksheet_set_zoom(lxw_worksheet *worksheet, uint16_t scale);

/**
 * @brief Set the option to display or hide gridlines on the screen and
 *        the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param option    Gridline option.
 *
 * Display or hide screen and print gridlines using one of the values of
 * @ref lxw_gridlines.
 *
 * @code
 *    worksheet_gridlines(worksheet1, LXW_HIDE_ALL_GRIDLINES);
 *
 *    worksheet_gridlines(worksheet2, LXW_SHOW_PRINT_GRIDLINES);
 * @endcode
 *
 * The Excel default is that the screen gridlines are on  and the printed
 * worksheet is off.
 *
 */
void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);

/**
 * @brief Center the printed page horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data horizontally between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_horizontally(worksheet);
 * @endcode
 *
 */
void worksheet_center_horizontally(lxw_worksheet *worksheet);

/**
 * @brief Center the printed page vertically.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data vertically between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_vertically(worksheet);
 * @endcode
 *
 */
void worksheet_center_vertically(lxw_worksheet *worksheet);

/**
 * @brief Set the option to print the row and column headers on the printed
 *        page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * When printing a worksheet from Excel the row and column headers (the row
 * numbers on the left and the column letters at the top) aren't printed by
 * default.
 *
 * This function sets the printer option to print these headers:
 *
 * @code
 *    worksheet_print_row_col_headers(worksheet);
 * @endcode
 *
 */
void worksheet_print_row_col_headers(lxw_worksheet *worksheet);

/**
 * @brief Set the number of rows to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row First row of repeat range.
 * @param last_row  Last row of repeat range.
 *
 * @return A #lxw_error code.
 *
 * For large Excel documents it is often desirable to have the first row or
 * rows of the worksheet print out at the top of each page.
 *
 * This can be achieved by using this function. The parameters `first_row`
 * and `last_row` are zero based:
 *
 * @code
 *     worksheet_repeat_rows(worksheet, 0, 0); // Repeat the first row.
 *     worksheet_repeat_rows(worksheet, 0, 1); // Repeat the first two rows.
 * @endcode
 */
lxw_error worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row,
                                lxw_row_t last_row);

/**
 * @brief Set the number of columns to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col First column of repeat range.
 * @param last_col  Last column of repeat range.
 *
 * @return A #lxw_error code.
 *
 * For large Excel documents it is often desirable to have the first column or
 * columns of the worksheet print out at the left of each page.
 *
 * This can be achieved by using this function. The parameters `first_col`
 * and `last_col` are zero based:
 *
 * @code
 *     worksheet_repeat_columns(worksheet, 0, 0); // Repeat the first col.
 *     worksheet_repeat_columns(worksheet, 0, 1); // Repeat the first two cols.
 * @endcode
 */
lxw_error worksheet_repeat_columns(lxw_worksheet *worksheet,
                                   lxw_col_t first_col, lxw_col_t last_col);

/**
 * @brief Set the print area for a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #lxw_error code.
 *
 * This function is used to specify the area of the worksheet that will be
 * printed. The RANGE() macro is often convenient for this.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 *
 * In order to set a row or column range you must specify the entire range:
 *
 * @code
 *     worksheet_print_area(worksheet, RANGE("A1:H1048576")); // Same as A:H.
 * @endcode
 */
lxw_error worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row,
                               lxw_col_t first_col, lxw_row_t last_row,
                               lxw_col_t last_col);
/**
 * @brief Fit the printed area to a specific number of pages both vertically
 *        and horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param width     Number of pages horizontally.
 * @param height    Number of pages vertically.
 *
 * The `%worksheet_fit_to_pages()` function is used to fit the printed area to
 * a specific number of pages both vertically and horizontally. If the printed
 * area exceeds the specified number of pages it will be scaled down to
 * fit. This ensures that the printed area will always appear on the specified
 * number of pages even if the page size or margins change:
 *
 * @code
 *     worksheet_fit_to_pages(worksheet1, 1, 1); // Fit to 1x1 pages.
 *     worksheet_fit_to_pages(worksheet2, 2, 1); // Fit to 2x1 pages.
 *     worksheet_fit_to_pages(worksheet3, 1, 2); // Fit to 1x2 pages.
 * @endcode
 *
 * The print area can be defined using the `worksheet_print_area()` function
 * as described above.
 *
 * A common requirement is to fit the printed output to `n` pages wide but
 * have the height be as long as necessary. To achieve this set the `height`
 * to zero:
 *
 * @code
 *     // 1 page wide and as long as necessary.
 *     worksheet_fit_to_pages(worksheet, 1, 0);
 * @endcode
 *
 * **Note**:
 *
 * - Although it is valid to use both `%worksheet_fit_to_pages()` and
 *   `worksheet_set_print_scale()` on the same worksheet Excel only allows one
 *   of these options to be active at a time. The last function call made will
 *   set the active option.
 *
 * - The `%worksheet_fit_to_pages()` function will override any manual page
 *   breaks that are defined in the worksheet.
 *
 * - When using `%worksheet_fit_to_pages()` it may also be required to set the
 *   printer paper size using `worksheet_set_paper()` or else Excel will
 *   default to "US Letter".
 *
 */
void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width,
                            uint16_t height);

/**
 * @brief Set the start page number when printing.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param start_page Starting page number.
 *
 * The `%worksheet_set_start_page()` function is used to set the number of
 * the starting page when the worksheet is printed out:
 *
 * @code
 *     // Start print from page 2.
 *     worksheet_set_start_page(worksheet, 2);
 * @endcode
 */
void worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);

/**
 * @brief Set the scale factor for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Print scale of worksheet to be printed.
 *
 * This function sets the scale factor of the printed page. The Scale factor
 * must be in the range `10 <= scale <= 400`:
 *
 * @code
 *     worksheet_set_print_scale(worksheet1, 75);
 *     worksheet_set_print_scale(worksheet2, 400);
 * @endcode
 *
 * The default scale factor is 100. Note, `%worksheet_set_print_scale()` does
 * not affect the scale of the visible page in Excel. For that you should use
 * `worksheet_set_zoom()`.
 *
 * Note that although it is valid to use both `worksheet_fit_to_pages()` and
 * `%worksheet_set_print_scale()` on the same worksheet Excel only allows one
 * of these options to be active at a time. The last function call made will
 * set the active option.
 *
 */
void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);

/**
 * @brief Display the worksheet cells from right to left for some versions of
 *        Excel.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
  * The `%worksheet_right_to_left()` function is used to change the default
 * direction of the worksheet from left-to-right, with the `A1` cell in the
 * top left, to right-to-left, with the `A1` cell in the top right.
 *
 * @code
 *     worksheet_right_to_left(worksheet1);
 * @endcode
 *
 * This is useful when creating Arabic, Hebrew or other near or far eastern
 * worksheets that use right-to-left as the default direction.
 */
void worksheet_right_to_left(lxw_worksheet *worksheet);

/**
 * @brief Hide zero values in worksheet cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_hide_zero()` function is used to hide any zero values that
 * appear in cells:
 *
 * @code
 *     worksheet_hide_zero(worksheet1);
 * @endcode
 */
void worksheet_hide_zero(lxw_worksheet *worksheet);

/**
 * @brief Set the color of the worksheet tab.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param color     The tab color.
 *
 * The `%worksheet_set_tab_color()` function is used to change the color of the worksheet
 * tab:
 *
 * @code
 *      worksheet_set_tab_color(worksheet1, LXW_COLOR_RED);
 *      worksheet_set_tab_color(worksheet2, LXW_COLOR_GREEN);
 *      worksheet_set_tab_color(worksheet3, 0xFF9900); // Orange.
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
void worksheet_set_tab_color(lxw_worksheet *worksheet, lxw_color_t color);

/**
 * @brief Protect elements of a worksheet from modification.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param password  A worksheet password.
 * @param options   Worksheet elements to protect.
 *
 * The `%worksheet_protect()` function protects worksheet elements from modification:
 *
 * @code
 *     worksheet_protect(worksheet, "Some Password", options);
 * @endcode
 *
 * The `password` and lxw_protection pointer are both optional:
 *
 * @code
 *     worksheet_protect(worksheet1, NULL,       NULL);
 *     worksheet_protect(worksheet2, NULL,       my_options);
 *     worksheet_protect(worksheet3, "password", NULL);
 *     worksheet_protect(worksheet4, "password", my_options);
 * @endcode
 *
 * Passing a `NULL` password is the same as turning on protection without a
 * password. Passing a `NULL` password and `NULL` options, or any other
 * combination has the effect of enabling a cell's `locked` and `hidden`
 * properties if they have been set.
 *
 * A *locked* cell cannot be edited and this property is on by default for all
 * cells. A *hidden* cell will display the results of a formula but not the
 * formula itself. These properties can be set using the format_set_unlocked()
 * and format_set_hidden() format functions.
 *
 * You can specify which worksheet elements you wish to protect by passing a
 * lxw_protection pointer in the `options` argument with any or all of the
 * following members set:
 *
 *     no_select_locked_cells
 *     no_select_unlocked_cells
 *     format_cells
 *     format_columns
 *     format_rows
 *     insert_columns
 *     insert_rows
 *     insert_hyperlinks
 *     delete_columns
 *     delete_rows
 *     sort
 *     autofilter
 *     pivot_tables
 *     scenarios
 *     objects
 *
 * All parameters are off by default. Individual elements can be protected as
 * follows:
 *
 * @code
 *     lxw_protection options = {
 *         .format_cells             = 1,
 *         .insert_hyperlinks        = 1,
 *         .insert_rows              = 1,
 *         .delete_rows              = 1,
 *         .insert_columns           = 1,
 *         .delete_columns           = 1,
 *     };
 *
 *     worksheet_protect(worksheet, NULL, &options);
 *
 * @endcode
 *
 * See also the format_set_unlocked() and format_set_hidden() format functions.
 *
 * **Note:** Worksheet level passwords in Excel offer **very** weak
 * protection. They don't encrypt your data and are very easy to
 * deactivate. Full workbook encryption is not supported by `libxlsxwriter`
 * since it requires a completely different file format and would take several
 * man months to implement.
 */
void worksheet_protect(lxw_worksheet *worksheet, const char *password,
                       lxw_protection *options);

/**
 * @brief Set the Outline and Grouping display properties.
 *
 * @param worksheet      Pointer to a lxw_worksheet instance to be updated.
 * @param visible        Outlines are visible. Optional, defaults to True.
 * @param symbols_below  Show row outline symbols below the outline bar.
 * @param symbols_right  Show column outline symbols to the right of outline.
 * @param auto_style     Use Automatic outline style.
 *
 * The `%worksheet_outline_settings()` method is used to control the
 * appearance of outlines in Excel. Outlines are described the section on
 * @ref working_with_outlines.
 *
 * The `visible` parameter is used to control whether or not outlines are
 * visible. Setting this parameter to False will cause all outlines on the
 * worksheet to be hidden. They can be un-hidden in Excel by means of the
 * "Show Outline Symbols" command button. The default Excel setting is True
 * for visible outlines.
 *
 * The `symbols_below` parameter is used to control whether the row outline
 * symbol will appear above or below the outline level bar. The default Excel
 * setting is True for symbols to appear below the outline level bar.
 *
 * The `symbols_right` parameter is used to control whether the column outline
 * symbol will appear to the left or the right of the outline level bar. The
 * default Excel setting is True for symbols to appear to the right of the
 * outline level bar.
 *
 * The `auto_style` parameter is used to control whether the automatic outline
 * generator in Excel uses automatic styles when creating an outline. This has
 * no effect on a file generated by XlsxWriter but it does have an effect on
 * how the worksheet behaves after it is created. The default Excel setting is
 * False for "Automatic Styles" to be turned off.
 *
 * The default settings for all of these parameters in libxlsxwriter
 * correspond to Excel's default parameters and are shown below:
 *
 * @code
 *     worksheet_outline_settings(worksheet1, LXW_TRUE, LXW_TRUE, LXW_TRUE, LXW_FALSE);
 * @endcode
 *
 * The worksheet parameters controlled by `worksheet_outline_settings()` are
 * rarely used.
 */
void worksheet_outline_settings(lxw_worksheet *worksheet, uint8_t visible,
                                uint8_t symbols_below, uint8_t symbols_right,
                                uint8_t auto_style);

/**
 * @brief Set the default row properties.
 *
 * @param worksheet        Pointer to a lxw_worksheet instance to be updated.
 * @param height           Default row height.
 * @param hide_unused_rows Hide unused cells.
 *
 * The `%worksheet_set_default_row()` function is used to set Excel default
 * row properties such as the default height and the option to hide unused
 * rows. These parameters are an optimization used by Excel to set row
 * properties without generating a very large file with an entry for each row.
 *
 * To set the default row height:
 *
 * @code
 *     worksheet_set_default_row(worksheet, 24, LXW_FALSE);
 *
 * @endcode
 *
 * To hide unused rows:
 *
 * @code
 *     worksheet_set_default_row(worksheet, 15, LXW_TRUE);
 * @endcode
 *
 * Note, in the previous case we use the default height #LXW_DEF_ROW_HEIGHT =
 * 15 so the the height remains unchanged.
 */
void worksheet_set_default_row(lxw_worksheet *worksheet, double height,
                               uint8_t hide_unused_rows);

lxw_worksheet *lxw_worksheet_new(lxw_worksheet_init_data *init_data);
void lxw_worksheet_free(lxw_worksheet *worksheet);
void lxw_worksheet_assemble_xml_file(lxw_worksheet *worksheet);
void lxw_worksheet_write_single_row(lxw_worksheet *worksheet);

void lxw_worksheet_prepare_image(lxw_worksheet *worksheet,
                                 uint16_t image_ref_id, uint16_t drawing_id,
                                 lxw_image_options *image_data);

void lxw_worksheet_prepare_chart(lxw_worksheet *worksheet,
                                 uint16_t chart_ref_id, uint16_t drawing_id,
                                 lxw_image_options *image_data);

lxw_row *lxw_worksheet_find_row(lxw_worksheet *worksheet, lxw_row_t row_num);
lxw_cell *lxw_worksheet_find_cell(lxw_row *row, lxw_col_t col_num);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _worksheet_xml_declaration(lxw_worksheet *worksheet);
STATIC void _worksheet_write_worksheet(lxw_worksheet *worksheet);
STATIC void _worksheet_write_dimension(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_view(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_views(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_format_pr(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_data(lxw_worksheet *worksheet);
STATIC void _worksheet_write_page_margins(lxw_worksheet *worksheet);
STATIC void _worksheet_write_page_setup(lxw_worksheet *worksheet);
STATIC void _worksheet_write_col_info(lxw_worksheet *worksheet,
                                      lxw_col_options *options);
STATIC void _write_row(lxw_worksheet *worksheet, lxw_row *row, char *spans);
STATIC lxw_row *_get_row_list(struct lxw_table_rows *table,
                              lxw_row_t row_num);

STATIC void _worksheet_write_merge_cell(lxw_worksheet *worksheet,
                                        lxw_merged_range *merged_range);
STATIC void _worksheet_write_merge_cells(lxw_worksheet *worksheet);

STATIC void _worksheet_write_odd_header(lxw_worksheet *worksheet);
STATIC void _worksheet_write_odd_footer(lxw_worksheet *worksheet);
STATIC void _worksheet_write_header_footer(lxw_worksheet *worksheet);

STATIC void _worksheet_write_print_options(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_pr(lxw_worksheet *worksheet);
STATIC void _worksheet_write_tab_color(lxw_worksheet *worksheet);
STATIC void _worksheet_write_sheet_protection(lxw_worksheet *worksheet);
STATIC void _worksheet_write_data_validations(lxw_worksheet *self);
#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_WORKSHEET_H__ */
