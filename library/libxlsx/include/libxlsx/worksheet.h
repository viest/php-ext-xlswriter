/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 */

/**
 * @page lxlsx_worksheet_page The Worksheet object
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
 * created by calling the lxlsx_workbook_add_worksheet() function from a
 * Workbook object:
 *
 * @code
 *     #include "libxlsx.h"
 *
 *     int main() {
 *
 *         lxlsx_workbook  *workbook  = lxlsx_workbook_new("filename.xlsx");
 *         lxlsx_worksheet *worksheet = lxlsx_workbook_add_worksheet(workbook, NULL);
 *
 *         lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello Excel", NULL);
 *
 *         return lxlsx_workbook_close(workbook);
 *     }
 * @endcode
 *
 */
#ifndef __LXLSX_WORKSHEET_H__
#define __LXLSX_WORKSHEET_H__

#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>

#include "shared_strings.h"
#include "chart.h"
#include "drawing.h"
#include "common.h"
#include "edit.h"
#include "format.h"
#include "cell.h"
#include "styles.h"
#include "utility.h"
#include "relationships.h"

#define LXLSX_ROW_MAX                 1048576
#define LXLSX_COL_MAX                 16384
#define LXLSX_COL_META_MAX            128
#define LXLSX_HEADER_FOOTER_MAX       255
#define LXLSX_MAX_NUMBER_URLS         65530
#define LXLSX_PANE_NAME_LENGTH        12  /* bottomRight + 1 */
#define LXLSX_IMAGE_BUFFER_SIZE       1024
#define LXLSX_HEADER_FOOTER_OBJS_MAX  6   /* Header/footer image objs. */

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
#define LXLSX_BREAKS_MAX        1023

/** Default Excel column width in character units. */
#define LXLSX_DEF_COL_WIDTH (double)8.43

/** Default Excel column height in character units. */
#define LXLSX_DEF_ROW_HEIGHT (double)15.0

/** Default Excel column width in pixels. */
#define LXLSX_DEF_COL_WIDTH_PIXELS 64

/** Default Excel column height in pixels. */
#define LXLSX_DEF_ROW_HEIGHT_PIXELS 20

/** Gridline options using in `lxlsx_worksheet_gridlines()`. */
enum lxlsx_gridlines {
    /** Hide screen and print gridlines. */
    LXLSX_HIDE_ALL_GRIDLINES = 0,

    /** Show screen gridlines. */
    LXLSX_SHOW_SCREEN_GRIDLINES,

    /** Show print gridlines. */
    LXLSX_SHOW_PRINT_GRIDLINES,

    /** Show screen and print gridlines. */
    LXLSX_SHOW_ALL_GRIDLINES
};

/** Data validation property values. */
enum lxlsx_validation_boolean {
    LXLSX_VALIDATION_DEFAULT,

    /** Turn a data validation property off. */
    LXLSX_VALIDATION_OFF,

    /** Turn a data validation property on. Data validation properties are
     * generally on by default. */
    LXLSX_VALIDATION_ON
};

/** Data validation types. */
enum lxlsx_validation_types {
    LXLSX_VALIDATION_TYPE_NONE,

    /** Restrict cell input to whole/integer numbers only. */
    LXLSX_VALIDATION_TYPE_INTEGER,

    /** Restrict cell input to whole/integer numbers only, using a cell
     *  reference. */
    LXLSX_VALIDATION_TYPE_INTEGER_FORMULA,

    /** Restrict cell input to decimal numbers only. */
    LXLSX_VALIDATION_TYPE_DECIMAL,

    /** Restrict cell input to decimal numbers only, using a cell
     * reference. */
    LXLSX_VALIDATION_TYPE_DECIMAL_FORMULA,

    /** Restrict cell input to a list of strings in a dropdown. */
    LXLSX_VALIDATION_TYPE_LIST,

    /** Restrict cell input to a list of strings in a dropdown, using a
     * cell range. */
    LXLSX_VALIDATION_TYPE_LIST_FORMULA,

    /** Restrict cell input to date values only, using a lxlsx_datetime type. */
    LXLSX_VALIDATION_TYPE_DATE,

    /** Restrict cell input to date values only, using a cell reference. */
    LXLSX_VALIDATION_TYPE_DATE_FORMULA,

    /* Restrict cell input to date values only, as a serial number.
     * Undocumented. */
    LXLSX_VALIDATION_TYPE_DATE_NUMBER,

    /** Restrict cell input to time values only, using a lxlsx_datetime type. */
    LXLSX_VALIDATION_TYPE_TIME,

    /** Restrict cell input to time values only, using a cell reference. */
    LXLSX_VALIDATION_TYPE_TIME_FORMULA,

    /* Restrict cell input to time values only, as a serial number.
     * Undocumented. */
    LXLSX_VALIDATION_TYPE_TIME_NUMBER,

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    LXLSX_VALIDATION_TYPE_LENGTH,

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    LXLSX_VALIDATION_TYPE_LENGTH_FORMULA,

    /** Restrict cell to input controlled by a custom formula that returns
     * `TRUE/FALSE`. */
    LXLSX_VALIDATION_TYPE_CUSTOM_FORMULA,

    /** Allow any type of input. Mainly only useful for pop-up messages. */
    LXLSX_VALIDATION_TYPE_ANY
};

/** Data validation criteria uses to control the selection of data. */
enum lxlsx_validation_criteria {
    LXLSX_VALIDATION_CRITERIA_NONE,

    /** Select data between two values. */
    LXLSX_VALIDATION_CRITERIA_BETWEEN,

    /** Select data that is not between two values. */
    LXLSX_VALIDATION_CRITERIA_NOT_BETWEEN,

    /** Select data equal to a value. */
    LXLSX_VALIDATION_CRITERIA_EQUAL_TO,

    /** Select data not equal to a value. */
    LXLSX_VALIDATION_CRITERIA_NOT_EQUAL_TO,

    /** Select data greater than a value. */
    LXLSX_VALIDATION_CRITERIA_GREATER_THAN,

    /** Select data less than a value. */
    LXLSX_VALIDATION_CRITERIA_LESS_THAN,

    /** Select data greater than or equal to a value. */
    LXLSX_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,

    /** Select data less than or equal to a value. */
    LXLSX_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO
};

/** Data validation error types for pop-up messages. */
enum lxlsx_validation_error_types {
    /** Show a "Stop" data validation pop-up message. This is the default. */
    LXLSX_VALIDATION_ERROR_TYPE_STOP,

    /** Show an "Error" data validation pop-up message. */
    LXLSX_VALIDATION_ERROR_TYPE_WARNING,

    /** Show an "Information" data validation pop-up message. */
    LXLSX_VALIDATION_ERROR_TYPE_INFORMATION
};

/** Set the display type for a cell comment. This is hidden by default but
 *  can be set to visible with the `lxlsx_worksheet_show_comments()` function. */
enum lxlsx_comment_display_types {
    /** Default to the worksheet default which can be hidden or visible.*/
    LXLSX_COMMENT_DISPLAY_DEFAULT,

    /** Hide the cell comment. Usually the default. */
    LXLSX_COMMENT_DISPLAY_HIDDEN,

    /** Show the cell comment. Can also be set for the worksheet with the
     *  `lxlsx_worksheet_show_comments()` function.*/
    LXLSX_COMMENT_DISPLAY_VISIBLE
};

/** @brief Type definitions for conditional formats.
 *
 * Values used to set the "type" field of conditional format.
 */
enum lxlsx_conditional_format_types {
    LXLSX_CONDITIONAL_TYPE_NONE,

    /** The Cell type is the most common conditional formatting type. It is
     *  used when a format is applied to a cell based on a simple
     *  criterion.  */
    LXLSX_CONDITIONAL_TYPE_CELL,

    /** The Text type is used to specify Excel's "Specific Text" style
     *  conditional format. */
    LXLSX_CONDITIONAL_TYPE_TEXT,

    /** The Time Period type is used to specify Excel's "Dates Occurring"
     *  style conditional format. */
    LXLSX_CONDITIONAL_TYPE_TIME_PERIOD,

    /** The Average type is used to specify Excel's "Average" style
     *  conditional format. */
    LXLSX_CONDITIONAL_TYPE_AVERAGE,

    /** The Duplicate type is used to highlight duplicate cells in a range. */
    LXLSX_CONDITIONAL_TYPE_DUPLICATE,

    /** The Unique type is used to highlight unique cells in a range. */
    LXLSX_CONDITIONAL_TYPE_UNIQUE,

    /** The Top type is used to specify the top n values by number or
     *  percentage in a range. */
    LXLSX_CONDITIONAL_TYPE_TOP,

    /** The Bottom type is used to specify the bottom n values by number or
     *  percentage in a range. */
    LXLSX_CONDITIONAL_TYPE_BOTTOM,

    /** The Blanks type is used to highlight blank cells in a range. */
    LXLSX_CONDITIONAL_TYPE_BLANKS,

    /** The No Blanks type is used to highlight non blank cells in a range. */
    LXLSX_CONDITIONAL_TYPE_NO_BLANKS,

    /** The Errors type is used to highlight error cells in a range. */
    LXLSX_CONDITIONAL_TYPE_ERRORS,

    /** The No Errors type is used to highlight non error cells in a range. */
    LXLSX_CONDITIONAL_TYPE_NO_ERRORS,

    /** The Formula type is used to specify a conditional format based on a
     *  user defined formula. */
    LXLSX_CONDITIONAL_TYPE_FORMULA,

    /** The 2 Color Scale type is used to specify Excel's "2 Color Scale"
     *  style conditional format. */
    LXLSX_CONDITIONAL_2_COLOR_SCALE,

    /** The 3 Color Scale type is used to specify Excel's "3 Color Scale"
     *  style conditional format. */
    LXLSX_CONDITIONAL_3_COLOR_SCALE,

    /** The Data Bar type is used to specify Excel's "Data Bar" style
     *  conditional format. */
    LXLSX_CONDITIONAL_DATA_BAR,

    /** The Icon Set type is used to specify a conditional format with a set
     *  of icons such as traffic lights or arrows. */
    LXLSX_CONDITIONAL_TYPE_ICON_SETS,

    LXLSX_CONDITIONAL_TYPE_LAST
};

/** @brief The criteria used in a conditional format.
 *
 * Criteria used to define how a conditional format works.
 */
enum lxlsx_conditional_criteria {
    LXLSX_CONDITIONAL_CRITERIA_NONE,

    /** Format cells equal to a value. */
    LXLSX_CONDITIONAL_CRITERIA_EQUAL_TO,

    /** Format cells not equal to a value. */
    LXLSX_CONDITIONAL_CRITERIA_NOT_EQUAL_TO,

    /** Format cells greater than a value. */
    LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN,

    /** Format cells less than a value. */
    LXLSX_CONDITIONAL_CRITERIA_LESS_THAN,

    /** Format cells greater than or equal to a value. */
    LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO,

    /** Format cells less than or equal to a value. */
    LXLSX_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO,

    /** Format cells between two values. */
    LXLSX_CONDITIONAL_CRITERIA_BETWEEN,

    /** Format cells that is not between two values. */
    LXLSX_CONDITIONAL_CRITERIA_NOT_BETWEEN,

    /** Format cells that contain the specified text. */
    LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING,

    /** Format cells that don't contain the specified text. */
    LXLSX_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING,

    /** Format cells that begin with the specified text. */
    LXLSX_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH,

    /** Format cells that end with the specified text. */
    LXLSX_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH,

    /** Format cells with a date of yesterday. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY,

    /** Format cells with a date of today. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY,

    /** Format cells with a date of tomorrow. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW,

    /** Format cells with a date in the last 7 days. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS,

    /** Format cells with a date in the last week. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK,

    /** Format cells with a date in the current week. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK,

    /** Format cells with a date in the next week. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK,

    /** Format cells with a date in the last month. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH,

    /** Format cells with a date in the current month. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH,

    /** Format cells with a date in the next month. */
    LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH,

    /** Format cells above the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE,

    /** Format cells below the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW,

    /** Format cells above or equal to the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL,

    /** Format cells below or equal to the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL,

    /** Format cells 1 standard deviation above the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE,

    /** Format cells 1 standard deviation below the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW,

    /** Format cells 2 standard deviation above the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE,

    /** Format cells 2 standard deviation below the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW,

    /** Format cells 3 standard deviation above the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE,

    /** Format cells 3 standard deviation below the average for the range. */
    LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW,

    /** Format cells in the top of bottom percentage. */
    LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT
};

/** @brief Conditional format rule types.
 *
 * Conditional format rule types that apply to Color Scale and Data Bars.
 */
enum lxlsx_conditional_format_rule_types {
    LXLSX_CONDITIONAL_RULE_TYPE_NONE,

    /** Conditional format rule type: matches the minimum values in the
     *  range. Can only be applied to min_rule_type.*/
    LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM,

    /** Conditional format rule type: use a number to set the bound.*/
    LXLSX_CONDITIONAL_RULE_TYPE_NUMBER,

    /** Conditional format rule type: use a percentage to set the bound.*/
    LXLSX_CONDITIONAL_RULE_TYPE_PERCENT,

    /** Conditional format rule type: use a percentile to set the bound.*/
    LXLSX_CONDITIONAL_RULE_TYPE_PERCENTILE,

    /** Conditional format rule type: use a formula to set the bound.*/
    LXLSX_CONDITIONAL_RULE_TYPE_FORMULA,

    /** Conditional format rule type: matches the maximum values in the
     *  range. Can only be applied to max_rule_type.*/
    LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM,

    /* Used internally for Excel2010 bars. Not documented. */
    LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MIN,

    /* Used internally for Excel2010 bars. Not documented. */
    LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MAX
};

/** @brief Conditional format data bar directions.
 *
 * Values used to set the bar direction of a conditional format data bar.
 */
enum lxlsx_conditional_format_bar_direction {

    /** Data bar direction is set by Excel based on the context of the data
     *  displayed. */
    LXLSX_CONDITIONAL_BAR_DIRECTION_CONTEXT,

    /** Data bar direction is from right to left. */
    LXLSX_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,

    /** Data bar direction is from left to right. */
    LXLSX_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT
};

/** @brief Conditional format data bar axis options.
 *
 * Values used to set the position of the axis in a conditional format data
 * bar.
 */
enum lxlsx_conditional_bar_axis_position {

    /** Data bar axis position is set by Excel based on the context of the
     *  data displayed. */
    LXLSX_CONDITIONAL_BAR_AXIS_AUTOMATIC,

    /** Data bar axis position is set at the midpoint. */
    LXLSX_CONDITIONAL_BAR_AXIS_MIDPOINT,

    /** Data bar axis is turned off. */
    LXLSX_CONDITIONAL_BAR_AXIS_NONE
};

/** @brief Icon types used in the #lxlsx_conditional_format icon_style field.
 *
 * Definitions of icon styles used with Icon Set conditional formats.
 */
enum lxlsx_conditional_icon_types {

    /** Icon style: 3 colored arrows showing up, sideways and down. */
    LXLSX_CONDITIONAL_ICONS_3_ARROWS_COLORED,

    /** Icon style: 3 gray arrows showing up, sideways and down. */
    LXLSX_CONDITIONAL_ICONS_3_ARROWS_GRAY,

    /** Icon style: 3 colored flags in red, yellow and green. */
    LXLSX_CONDITIONAL_ICONS_3_FLAGS,

    /** Icon style: 3 traffic lights - rounded. */
    LXLSX_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,

    /** Icon style: 3 traffic lights with a rim - squarish. */
    LXLSX_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,

    /** Icon style: 3 colored shapes - a circle, triangle and diamond. */
    LXLSX_CONDITIONAL_ICONS_3_SIGNS,

    /** Icon style: 3 circled symbols with tick mark, exclamation and
     *  cross. */
    LXLSX_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,

    /** Icon style: 3 symbols with tick mark, exclamation and cross. */
    LXLSX_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED,

    /** Icon style: 4 colored arrows showing up, diagonal up, diagonal down
     *  and down. */
    LXLSX_CONDITIONAL_ICONS_4_ARROWS_COLORED,

    /** Icon style: 4 gray arrows showing up, diagonal up, diagonal down and
     * down. */
    LXLSX_CONDITIONAL_ICONS_4_ARROWS_GRAY,

    /** Icon style: 4 circles in 4 colors going from red to black. */
    LXLSX_CONDITIONAL_ICONS_4_RED_TO_BLACK,

    /** Icon style: 4 histogram ratings. */
    LXLSX_CONDITIONAL_ICONS_4_RATINGS,

    /** Icon style: 4 traffic lights. */
    LXLSX_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,

    /** Icon style: 5 colored arrows showing up, diagonal up, sideways,
     * diagonal down and down. */
    LXLSX_CONDITIONAL_ICONS_5_ARROWS_COLORED,

    /** Icon style: 5 gray arrows showing up, diagonal up, sideways, diagonal
     *  down and down. */
    LXLSX_CONDITIONAL_ICONS_5_ARROWS_GRAY,

    /** Icon style: 5 histogram ratings. */
    LXLSX_CONDITIONAL_ICONS_5_RATINGS,

    /** Icon style: 5 quarters, from 0 to 4 quadrants filled. */
    LXLSX_CONDITIONAL_ICONS_5_QUARTERS
};

/** @brief The type of table style.
 *
 * The type of table style (Light, Medium or Dark).
 */
enum lxlsx_table_style_type {

    LXLSX_TABLE_STYLE_TYPE_DEFAULT,

    /** Light table style. */
    LXLSX_TABLE_STYLE_TYPE_LIGHT,

    /** Light table style. */
    LXLSX_TABLE_STYLE_TYPE_MEDIUM,

    /** Light table style. */
    LXLSX_TABLE_STYLE_TYPE_DARK
};

/**
 * @brief Standard Excel functions for totals in tables.
 *
 * Definitions for the standard Excel functions that are available via the
 * dropdown in the total row of an Excel table.
 *
 */
enum lxlsx_table_total_functions {

    LXLSX_TABLE_FUNCTION_NONE = 0,

    /** Use the average function as the table total. */
    LXLSX_TABLE_FUNCTION_AVERAGE = 101,

    /** Use the count numbers function as the table total. */
    LXLSX_TABLE_FUNCTION_COUNT_NUMS = 102,

    /** Use the count function as the table total. */
    LXLSX_TABLE_FUNCTION_COUNT = 103,

    /** Use the max function as the table total. */
    LXLSX_TABLE_FUNCTION_MAX = 104,

    /** Use the min function as the table total. */
    LXLSX_TABLE_FUNCTION_MIN = 105,

    /** Use the standard deviation function as the table total. */
    LXLSX_TABLE_FUNCTION_STD_DEV = 107,

    /** Use the sum function as the table total. */
    LXLSX_TABLE_FUNCTION_SUM = 109,

    /** Use the var function as the table total. */
    LXLSX_TABLE_FUNCTION_VAR = 110
};

/** @brief The criteria used in autofilter rules.
 *
 * Criteria used to define an autofilter rule condition.
 */
enum lxlsx_filter_criteria {
    LXLSX_FILTER_CRITERIA_NONE,

    /** Filter cells equal to a value. */
    LXLSX_FILTER_CRITERIA_EQUAL_TO,

    /** Filter cells not equal to a value. */
    LXLSX_FILTER_CRITERIA_NOT_EQUAL_TO,

    /** Filter cells greater than a value. */
    LXLSX_FILTER_CRITERIA_GREATER_THAN,

    /** Filter cells less than a value. */
    LXLSX_FILTER_CRITERIA_LESS_THAN,

    /** Filter cells greater than or equal to a value. */
    LXLSX_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,

    /** Filter cells less than or equal to a value. */
    LXLSX_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,

    /** Filter cells that are blank. */
    LXLSX_FILTER_CRITERIA_BLANKS,

    /** Filter cells that are not blank. */
    LXLSX_FILTER_CRITERIA_NON_BLANKS
};

/**
 * @brief And/or operator when using 2 filter rules.
 *
 * And/or operator conditions when using 2 filter rules with
 * lxlsx_worksheet_filter_column2(). In general LXLSX_FILTER_OR is used with
 * LXLSX_FILTER_CRITERIA_EQUAL_TO and LXLSX_FILTER_AND is used with the other
 * filter criteria.
 */
enum lxlsx_filter_operator {
    /** Logical "and" of 2 filter rules. */
    LXLSX_FILTER_AND,

    /** Logical "or" of 2 filter rules. */
    LXLSX_FILTER_OR
};

/* Internal filter types. */
enum lxlsx_filter_type {
    LXLSX_FILTER_TYPE_NONE,

    LXLSX_FILTER_TYPE_SINGLE,

    LXLSX_FILTER_TYPE_AND,

    LXLSX_FILTER_TYPE_OR,

    LXLSX_FILTER_TYPE_STRING_LIST
};

/** Options to control the positioning of worksheet objects such as images
 *  or charts. See @ref working_with_object_positioning. */
enum lxlsx_object_position {

    /** Default positioning for the object. */
    LXLSX_OBJECT_POSITION_DEFAULT,

    /** Move and size the worksheet object with the cells. */
    LXLSX_OBJECT_MOVE_AND_SIZE,

    /** Move but don't size the worksheet object with the cells. */
    LXLSX_OBJECT_MOVE_DONT_SIZE,

    /** Don't move or size the worksheet object with the cells. */
    LXLSX_OBJECT_DONT_MOVE_DONT_SIZE,

    /** Same as #LXLSX_OBJECT_MOVE_AND_SIZE except libxlsxwriter applies hidden
     *  cells after the object is inserted. */
    LXLSX_OBJECT_MOVE_AND_SIZE_AFTER
};

/** Options for ignoring worksheet errors/warnings. See lxlsx_worksheet_ignore_errors(). */
enum lxlsx_ignore_errors {

    /** Turn off errors/warnings for numbers stores as text. */
    LXLSX_IGNORE_NUMBER_STORED_AS_TEXT = 1,

    /** Turn off errors/warnings for formula errors (such as divide by
     *  zero). */
    LXLSX_IGNORE_EVAL_ERROR,

    /** Turn off errors/warnings for formulas that differ from surrounding
     *  formulas. */
    LXLSX_IGNORE_FORMULA_DIFFERS,

    /** Turn off errors/warnings for formulas that omit cells in a range. */
    LXLSX_IGNORE_FORMULA_RANGE,

    /** Turn off errors/warnings for unlocked cells that contain formulas. */
    LXLSX_IGNORE_FORMULA_UNLOCKED,

    /** Turn off errors/warnings for formulas that refer to empty cells. */
    LXLSX_IGNORE_EMPTY_CELL_REFERENCE,

    /** Turn off errors/warnings for cells in a table that do not comply with
     *  applicable data validation rules. */
    LXLSX_IGNORE_LIST_DATA_VALIDATION,

    /** Turn off errors/warnings for cell formulas that differ from the column
     *  formula. */
    LXLSX_IGNORE_CALCULATED_COLUMN,

    /** Turn off errors/warnings for formulas that contain a two digit text
     *  representation of a year. */
    LXLSX_IGNORE_TWO_DIGIT_TEXT_YEAR,

    LXLSX_IGNORE_LAST_OPTION
};

enum pane_types {
    NO_PANES = 0,
    FREEZE_PANES,
    SPLIT_PANES,
    FREEZE_SPLIT_PANES
};

enum lxlsx_image_position {
    HEADER_LEFT = 0,
    HEADER_CENTER,
    HEADER_RIGHT,
    FOOTER_LEFT,
    FOOTER_CENTER,
    FOOTER_RIGHT
};

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxlsx_table_cells, lxlsx_cell);
RB_HEAD(lxlsx_drawing_rel_ids, lxlsx_drawing_rel_id);
RB_HEAD(lxlsx_vml_drawing_rel_ids, lxlsx_drawing_rel_id);
RB_HEAD(lxlsx_cond_format_hash, lxlsx_cond_format_hash_element);

/* Define a RB_TREE struct manually to add extra members. */
struct lxlsx_table_rows {
    struct lxlsx_row *rbh_root;
    struct lxlsx_row *cached_row;
    lxlsx_row_t cached_row_num;
};

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXLSX_RB_GENERATE_ROW(name, type, field, cmp)       \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxlsx_rb_generate_row{int unused;}

#define LXLSX_RB_GENERATE_CELL(name, type, field, cmp)      \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxlsx_rb_generate_cell{int unused;}

#define LXLSX_RB_GENERATE_DRAWING_REL_IDS(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)         \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)         \
    RB_GENERATE_INSERT(name, type, field, cmp, static)          \
    RB_GENERATE_REMOVE(name, type, field, static)               \
    RB_GENERATE_FIND(name, type, field, cmp, static)            \
    RB_GENERATE_NEXT(name, type, field, static)                 \
    RB_GENERATE_MINMAX(name, type, field, static)               \
    /* Add unused struct to allow adding a semicolon */         \
    struct lxlsx_rb_generate_drawing_rel_ids{int unused;}

#define LXLSX_RB_GENERATE_VML_DRAWING_REL_IDS(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)         \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)         \
    RB_GENERATE_INSERT(name, type, field, cmp, static)          \
    RB_GENERATE_REMOVE(name, type, field, static)               \
    RB_GENERATE_FIND(name, type, field, cmp, static)            \
    RB_GENERATE_NEXT(name, type, field, static)                 \
    RB_GENERATE_MINMAX(name, type, field, static)               \
    /* Add unused struct to allow adding a semicolon */         \
    struct lxlsx_rb_generate_vml_drawing_rel_ids{int unused;}

#define LXLSX_RB_GENERATE_COND_FORMAT_HASH(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)         \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)         \
    RB_GENERATE_INSERT(name, type, field, cmp, static)          \
    RB_GENERATE_REMOVE(name, type, field, static)               \
    RB_GENERATE_FIND(name, type, field, cmp, static)            \
    RB_GENERATE_NEXT(name, type, field, static)                 \
    RB_GENERATE_MINMAX(name, type, field, static)               \
    /* Add unused struct to allow adding a semicolon */         \
    struct lxlsx_rb_generate_cond_format_hash{int unused;}

STAILQ_HEAD(lxlsx_merged_ranges, lxlsx_merged_range);
STAILQ_HEAD(lxlsx_selections, lxlsx_selection);
STAILQ_HEAD(lxlsx_data_validations, lxlsx_data_val_obj);
STAILQ_HEAD(lxlsx_cond_format_list, lxlsx_cond_format_obj);
STAILQ_HEAD(lxlsx_image_props, lxlsx_object_properties);
STAILQ_HEAD(lxlsx_embedded_image_props, lxlsx_object_properties);
STAILQ_HEAD(lxlsx_chart_props, lxlsx_object_properties);
STAILQ_HEAD(lxlsx_comment_objs, lxlsx_vml_obj);
STAILQ_HEAD(lxlsx_table_objs, lxlsx_table_obj);

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the lxlsx_worksheet_set_column() and lxlsx_worksheet_set_row()
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
typedef struct lxlsx_row_col_options {
    /** Hide the row/column. @ref ww_outlines_grouping.*/
    uint8_t hidden;

    /** Outline level. See @ref ww_outlines_grouping.*/
    uint8_t level;

    /** Set the outline row as collapsed. See @ref ww_outlines_grouping.*/
    uint8_t collapsed;
} lxlsx_row_col_options;

typedef struct lxlsx_col_options {
    lxlsx_col_t firstcol;
    lxlsx_col_t lastcol;
    double width;
    lxlsx_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
} lxlsx_col_options;

typedef struct lxlsx_merged_range {
    lxlsx_row_t first_row;
    lxlsx_row_t last_row;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;

    STAILQ_ENTRY (lxlsx_merged_range) list_pointers;
} lxlsx_merged_range;

typedef struct lxlsx_repeat_rows {
    uint8_t in_use;
    lxlsx_row_t first_row;
    lxlsx_row_t last_row;
} lxlsx_repeat_rows;

typedef struct lxlsx_repeat_cols {
    uint8_t in_use;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;
} lxlsx_repeat_cols;

typedef struct lxlsx_print_area {
    uint8_t in_use;
    lxlsx_row_t first_row;
    lxlsx_row_t last_row;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;
} lxlsx_print_area;

typedef struct lxlsx_autofilter {
    uint8_t in_use;
    uint8_t has_rules;
    lxlsx_row_t first_row;
    lxlsx_row_t last_row;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;
} lxlsx_autofilter;

typedef struct lxlsx_panes {
    uint8_t type;
    lxlsx_row_t first_row;
    lxlsx_col_t first_col;
    lxlsx_row_t top_row;
    lxlsx_col_t left_col;
    double x_split;
    double y_split;
} lxlsx_panes;

typedef struct lxlsx_selection {
    char pane[LXLSX_PANE_NAME_LENGTH];
    char active_cell[LXLSX_MAX_CELL_RANGE_LENGTH];
    char sqref[LXLSX_MAX_CELL_RANGE_LENGTH];

    STAILQ_ENTRY (lxlsx_selection) list_pointers;

} lxlsx_selection;

/**
 * @brief Worksheet data validation options.
 */
typedef struct lxlsx_data_validation {

    /**
     * Set the validation type. Should be a #lxlsx_validation_types value.
     */
    uint8_t validate;

    /**
     * Set the validation criteria type to select the data. Should be a
     * #lxlsx_validation_criteria value.
     */
    uint8_t criteria;

    /** Controls whether a data validation is not applied to blank data in the
     * cell. Should be a #lxlsx_validation_boolean value. It is on by
     * default.
     */
    uint8_t ignore_blank;

    /**
     * This parameter is used to toggle on and off the 'Show input message
     * when cell is selected' option in the Excel data validation dialog. When
     * the option is off an input message is not displayed even if it has been
     * set using input_message. Should be a #lxlsx_validation_boolean value. It
     * is on by default.
     */
    uint8_t show_input;

    /**
     * This parameter is used to toggle on and off the 'Show error alert
     * after invalid data is entered' option in the Excel data validation
     * dialog. When the option is off an error message is not displayed even
     * if it has been set using error_message. Should be a
     * #lxlsx_validation_boolean value. It is on by default.
     */
    uint8_t show_error;

    /**
     * This parameter is used to specify the type of error dialog that is
     * displayed. Should be a #lxlsx_validation_error_types value.
     */
    uint8_t error_type;

    /**
     * This parameter is used to toggle on and off the 'In-cell dropdown'
     * option in the Excel data validation dialog. When the option is on a
     * dropdown list will be shown for list validations. Should be a
     * #lxlsx_validation_boolean value. It is on by default.
     */
    uint8_t dropdown;

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
    const char *value_formula;

    /**
     * This parameter is used to set a list of strings for a dropdown list.
     * The list should be a `NULL` terminated array of char* strings:
     *
     * @code
     *    char *list[] = {"open", "high", "close", NULL};
     *
     *    data_validation->validate   = LXLSX_VALIDATION_TYPE_LIST;
     *    data_validation->value_list = list;
     * @endcode
     *
     * The `value_formula` parameter can also be used to specify a list from
     * an Excel cell range.
     *
     * Note, the string list is restricted by Excel to 255 characters,
     * including comma separators.
     */
    const char **value_list;

    /**
     * This parameter is used to set the limiting value to which the date or
     * time criteria is applied using a #lxlsx_datetime struct.
     */
    lxlsx_datetime value_datetime;

    /**
     * This parameter is the same as `value_number` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    double minimum_number;

    /**
     * This parameter is the same as `value_formula` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    const char *minimum_formula;

    /**
     * This parameter is the same as `value_datetime` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    lxlsx_datetime minimum_datetime;

    /**
     * This parameter is the same as `value_number` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    double maximum_number;

    /**
     * This parameter is the same as `value_formula` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    const char *maximum_formula;

    /**
     * This parameter is the same as `value_datetime` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    lxlsx_datetime maximum_datetime;

    /**
     * The input_title parameter is used to set the title of the input message
     * that is displayed when a cell is entered. It has no default value and
     * is only displayed if the input message is displayed. See the
     * `input_message` parameter below.
     *
     * The maximum title length is 32 characters.
     */
    const char *input_title;

    /**
     * The input_message parameter is used to set the input message that is
     * displayed when a cell is entered. It has no default value.
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    const char *input_message;

    /**
     * The error_title parameter is used to set the title of the error message
     * that is displayed when the data validation criteria is not met. The
     * default error title is 'Microsoft Excel'. The maximum title length is
     * 32 characters.
     */
    const char *error_title;

    /**
     * The error_message parameter is used to set the error message that is
     * displayed when a cell is entered. The default error message is "The
     * value you entered is not valid. A user has restricted values that can
     * be entered into the cell".
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    const char *error_message;

} lxlsx_data_validation;

/* A copy of lxlsx_data_validation which is used internally and which contains
 * some additional fields.
 */
typedef struct lxlsx_data_val_obj {
    uint8_t validate;
    uint8_t criteria;
    uint8_t ignore_blank;
    uint8_t show_input;
    uint8_t show_error;
    uint8_t error_type;
    uint8_t dropdown;
    double value_number;
    char *value_formula;
    char **value_list;
    double minimum_number;
    char *minimum_formula;
    lxlsx_datetime minimum_datetime;
    double maximum_number;
    char *maximum_formula;
    lxlsx_datetime maximum_datetime;
    char *input_title;
    char *input_message;
    char *error_title;
    char *error_message;
    char sqref[LXLSX_MAX_CELL_RANGE_LENGTH];

    STAILQ_ENTRY (lxlsx_data_val_obj) list_pointers;
} lxlsx_data_val_obj;

/**
 * @brief Worksheet conditional formatting options.
 *
 * The fields/options in the the lxlsx_conditional_format are used to define a
 * worksheet conditional format. It is used in conjunction with
 * `lxlsx_worksheet_conditional_format()`.
 *
 */
typedef struct lxlsx_conditional_format {

    /** The type of conditional format such as #LXLSX_CONDITIONAL_TYPE_CELL or
     *  #LXLSX_CONDITIONAL_DATA_BAR. Should be a #lxlsx_conditional_format_types
     *  value.*/
    uint8_t type;

    /** The criteria parameter is used to set the criteria by which the cell
     *  data will be evaluated. For example in the expression `a > 5 the
     *  criteria is `>` or, in libxlsxwriter terms,
     *  #LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN. The criteria that are
     *  applicable depend on the conditional format type.  The criteria
     *  options are defined in #lxlsx_conditional_criteria. */
    uint8_t criteria;

    /** The number value to which the condition refers. For example in the
     * expression `a > 5`, the value is 5.*/
    double value;

    /** The string value to which the condition refers, such as `"=A1"`. If a
     *  value_string exists in the struct then the number value is
     *  ignored. Note, if the condition refers to a text string then it must
     *  be double quoted like this `"foo"`. */
    const char *value_string;

    /** The format field is used to specify the #lxlsx_format format that will
     *  be applied to the cell when the conditional formatting criterion is
     *  met. The #lxlsx_format is created using the `lxlsx_workbook_add_format()`
     *  method in the same way as cell formats.
     *
     *  @note In Excel, a conditional format is superimposed over the existing
     *  cell format and not all cell format properties can be
     *  modified. Properties that @b cannot be modified, in Excel, by a
     *  conditional format are: font name, font size, superscript and
     *  subscript, diagonal borders, all alignment properties and all
     *  protection properties. */
    lxlsx_format *format;

    /** The minimum value used for Cell, Color Scale and Data Bar conditional
     *  formats. For Cell types this is usually used with a "Between" style criteria. */
    double min_value;

    /** The minimum string value used for Cell, Color Scale and Data Bar conditional
     *  formats. Usually used to set ranges like `=A1`. */
    const char *min_value_string;

    /** The rule used for the minimum condition in Color Scale and Data Bar
     *  conditional formats. The rule types are defined in
     *  #lxlsx_conditional_format_rule_types. */
    uint8_t min_rule_type;

    /** The color used for the minimum Color Scale conditional format.
     *  See @ref working_with_colors. */
    lxlsx_color_t min_color;

    /** The middle value used for Color Scale and Data Bar conditional
     *  formats.  */
    double mid_value;

    /** The middle string value used for Color Scale and Data Bar conditional
     *  formats. Usually used to set ranges like `=A1`. */
    const char *mid_value_string;

    /** The rule used for the middle condition in Color Scale and Data Bar
     *  conditional formats. The rule types are defined in
     *  #lxlsx_conditional_format_rule_types. */
    uint8_t mid_rule_type;

    /** The color used for the middle Color Scale conditional format.
     *  See @ref working_with_colors. */
    lxlsx_color_t mid_color;

    /** The maximum value used for Cell, Color Scale and Data Bar conditional
     *  formats. For Cell types this is usually used with a "Between" style
     *  criteria. */
    double max_value;

    /** The maximum string value used for Cell, Color Scale and Data Bar conditional
     *  formats. Usually used to set ranges like `=A1`. */
    const char *max_value_string;

    /** The rule used for the maximum condition in Color Scale and Data Bar
     *  conditional formats. The rule types are defined in
     *  #lxlsx_conditional_format_rule_types. */
    uint8_t max_rule_type;

    /** The color used for the maximum Color Scale conditional format.
     *  See @ref working_with_colors. */
    lxlsx_color_t max_color;

    /** The bar_color field sets the fill color for data bars. See @ref
     *  working_with_colors. */
    lxlsx_color_t bar_color;

    /** The bar_only field sets The bar_only field displays a bar data but
     *  not the data in the cells. */
    uint8_t bar_only;

    /** In Excel 2010 additional data bar properties were added such as solid
     *  (non-gradient) bars and control over how negative values are
     *  displayed. These properties can shown below.
     *
     *  The data_bar_2010 field sets Excel 2010 style data bars even when
     *  Excel 2010 specific properties aren't used. */
    uint8_t data_bar_2010;

    /** The bar_solid field turns on a solid (non-gradient) fill for data
     *  bars. Set to LXLSX_TRUE to turn on. Excel 2010 only. */
    uint8_t bar_solid;

    /** The bar_negative_color field sets the color fill for the negative
     *  portion of a data bar. See @ref working_with_colors. Excel 2010 only. */
    lxlsx_color_t bar_negative_color;

    /** The bar_border_color field sets the color for the border line of a
     *  data bar. See @ref working_with_colors. Excel 2010 only. */
    lxlsx_color_t bar_border_color;

    /** The bar_negative_border_color field sets the color for the border of
     *  the negative portion of a data bar. See @ref
     *  working_with_colors. Excel 2010 only. */
    lxlsx_color_t bar_negative_border_color;

    /** The bar_negative_color_same field sets the fill color for the negative
     *  portion of a data bar to be the same as the fill color for the
     *  positive portion of the data bar. Set to LXLSX_TRUE to turn on. Excel
     *  2010 only. */
    uint8_t bar_negative_color_same;

    /** The bar_negative_border_color_same field sets the border color for the
     *  negative portion of a data bar to be the same as the border color for
     *  the positive portion of the data bar. Set to LXLSX_TRUE to turn
     *  on. Excel 2010 only. */
    uint8_t bar_negative_border_color_same;

    /** The bar_no_border field turns off the border for data bars. Set to
     *  LXLSX_TRUE to enable. Excel 2010 only. */
    uint8_t bar_no_border;

    /** The bar_direction field sets the direction for data bars. This
     *  property can be either left for left-to-right or right for
     *  right-to-left. If the property isn't set then Excel will adjust the
     *  position automatically based on the context. Should be a
     *  #lxlsx_conditional_format_bar_direction value. Excel 2010 only. */
    uint8_t bar_direction;

    /** The bar_axis_position field sets the position within the cells for the
     *  axis that is shown in data bars when there are negative values to
     *  display. The property can be either middle or none. If the property
     *  isn't set then Excel will position the axis based on the range of
     *  positive and negative values. Should be a
     *  lxlsx_conditional_bar_axis_position value. Excel 2010 only. */
    uint8_t bar_axis_position;

    /** The bar_axis_color field sets the color for the axis that is shown
     *  in data bars when there are negative values to display. See @ref
     *  working_with_colors. Excel 2010 only. */
    lxlsx_color_t bar_axis_color;

    /** The Icons Sets style is specified by the icon_style parameter. Should
     *  be a #lxlsx_conditional_icon_types. */
    uint8_t icon_style;

    /** The order of Icon Sets icons can be reversed by setting reverse_icons
     *  to LXLSX_TRUE.  */
    uint8_t reverse_icons;

    /** The icons can be displayed without the cell value by settings the
     *  icons_only parameter to LXLSX_TRUE.  */
    uint8_t icons_only;

    /** The multi_range field is used to extend a conditional format over
     *  non-contiguous ranges.
     *
     *  It is possible to apply the conditional format to different cell
     *  ranges in a worksheet using multiple calls to
     *  `lxlsx_worksheet_conditional_format()`. However, as a minor optimization it
     *  is also possible in Excel to apply the same conditional format to
     *  different non-contiguous cell ranges.
     *
     *  This is replicated in `lxlsx_worksheet_conditional_format()` using the
     *  multi_range option. The range must contain the primary range for the
     *  conditional format and any others separated by spaces. For example
     *  `"A1 C1:C5 E2 G1:G100"`.
     */
    const char *multi_range;

    /** The stop_if_true parameter can be used to set the "stop if true"
     *  feature of a conditional formatting rule when more than one rule is
     *  applied to a cell or a range of cells. When this parameter is set then
     *  subsequent rules are not evaluated if the current rule is true. Set to
     *  LXLSX_TRUE to turn on. */
    uint8_t stop_if_true;

} lxlsx_conditional_format;

/* Internal */
typedef struct lxlsx_cond_format_obj {
    uint8_t type;
    uint8_t criteria;

    double min_value;
    char *min_value_string;
    uint8_t min_rule_type;
    lxlsx_color_t min_color;

    double mid_value;
    char *mid_value_string;
    uint8_t mid_value_type;
    uint8_t mid_rule_type;
    lxlsx_color_t mid_color;

    double max_value;
    char *max_value_string;
    uint8_t max_value_type;
    uint8_t max_rule_type;
    lxlsx_color_t max_color;

    uint8_t data_bar_2010;
    uint8_t auto_min;
    uint8_t auto_max;
    uint8_t bar_only;
    uint8_t bar_solid;
    uint8_t bar_negative_color_same;
    uint8_t bar_negative_border_color_same;
    uint8_t bar_no_border;
    uint8_t bar_direction;
    uint8_t bar_axis_position;
    lxlsx_color_t bar_color;
    lxlsx_color_t bar_negative_color;
    lxlsx_color_t bar_border_color;
    lxlsx_color_t bar_negative_border_color;
    lxlsx_color_t bar_axis_color;

    uint8_t icon_style;
    uint8_t reverse_icons;
    uint8_t icons_only;

    uint8_t stop_if_true;
    uint8_t has_max;
    char *type_string;
    char *guid;

    int32_t dxf_index;
    uint32_t dxf_priority;

    char first_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char sqref[LXLSX_MAX_ATTRIBUTE_LENGTH];

    STAILQ_ENTRY (lxlsx_cond_format_obj) list_pointers;
} lxlsx_cond_format_obj;

typedef struct lxlsx_cond_format_hash_element {
    char sqref[LXLSX_MAX_ATTRIBUTE_LENGTH];

    struct lxlsx_cond_format_list *cond_formats;

    RB_ENTRY (lxlsx_cond_format_hash_element) tree_pointers;
} lxlsx_cond_format_hash_element;

/**
 * @brief Table columns options.
 *
 * Structure to set the options of a table column added with
 * lxlsx_worksheet_add_table(). See @ref ww_tables_columns.
 */
typedef struct lxlsx_table_column {

    /** Set the header name/caption for the column. If NULL the header defaults
     *  to  Column 1, Column 2, etc. */
    const char *header;

    /** Set the formula for the column. */
    const char *formula;

    /** Set the string description for the column total.  */
    const char *total_string;

    /** Set the function for the column total.  */
    uint8_t total_function;

    /** Set the format for the column header.  */
    lxlsx_format *header_format;

    /** Set the format for the data rows in the column.  */
    lxlsx_format *format;

    /** Set the formula value for the column total (not generally required). */
    double total_value;

} lxlsx_table_column;

/**
 * @brief Worksheet table options.
 *
 * Options used to define worksheet tables. See @ref working_with_tables for
 * more information.
 *
 */
typedef struct lxlsx_table_options {

    /**
     * The `name` parameter is used to set the name of the table. This
     * parameter is optional and by default tables are named `Table1`,
     * `Table2`, etc. in the worksheet order that they are added.
     *
     * @code
     *     lxlsx_table_options options = {.name = "Sales"};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:G8"), &options);
     * @endcode
     *
     * If you override the table name you must ensure that it doesn't clash
     * with an existing table name and that it follows Excel's requirements
     * for table names, see the Microsoft Office documentation on
     * [Naming an Excel Table]
     * (https://support.microsoft.com/en-us/office/rename-an-excel-table-fbf49a4f-82a3-43eb-8ba2-44d21233b114).
     */
    const char *name;

    /**
     * The `no_header_row` parameter can be used to turn off the header row in
     * the table. It is on by default:
     *
     * @code
     *     lxlsx_table_options options = {.no_header_row = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B4:F7"), &options);
     * @endcode
     *
     * @image html tables4.png
     *
     * Without this option the header row will contain default captions such
     * as `Column 1`, ``Column 2``, etc. These captions can be overridden
     * using the `columns` parameter shown below.
     *
     */
    uint8_t no_header_row;

    /**
     * The `no_autofilter` parameter can be used to turn off the autofilter in
     * the header row. It is on by default:
     *
     * @code
     *     lxlsx_table_options options = {.no_autofilter = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:F7"), &options);
     * @endcode
     *
     * @image html tables3.png
     *
     * The autofilter is only shown if the `no_header_row` parameter is off
     * (the default). Filter conditions within the table are not supported.
     *
     */
    uint8_t no_autofilter;

    /**
     * The `no_banded_rows` parameter can be used to turn off the rows of alternating
     * color in the table. It is on by default:
     *
     * @code
     *     lxlsx_table_options options = {.no_banded_rows = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:F7"), &options);
     * @endcode
     *
     * @image html tables6.png
     *
     */
    uint8_t no_banded_rows;

    /**
     * The `banded_columns` parameter can be used to used to create columns of
     * alternating color in the table. It is off by default:
     *
     * @code
     *     lxlsx_table_options options = {.banded_columns = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:F7"), &options);
     * @endcode
     *
     * The banded columns formatting is shown in the image in the previous
     * section above.
     */
    uint8_t banded_columns;

    /**
     * The `first_column` parameter can be used to highlight the first column
     * of the table. The type of highlighting will depend on the `style_type`
     * of the table. It may be bold text or a different color. It is off by
     * default:
     *
     * @code
     *     lxlsx_table_options options = {.first_column = LXLSX_TRUE, .last_column = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:F7"), &options);
     * @endcode
     *
     * @image html tables5.png
     */
    uint8_t first_column;

    /**
     * The `last_column` parameter can be used to highlight the last column of
     * the table. The type of highlighting will depend on the `style` of the
     * table. It may be bold text or a different color. It is off by default:
     *
     * @code
     *     lxlsx_table_options options = {.first_column = LXLSX_TRUE, .last_column = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:F7"), &options);
     * @endcode
     *
     * The `last_column` formatting is shown in the image in the previous
     * section above.
     */
    uint8_t last_column;

    /**
     * The `style_type` parameter can be used to set the style of the table,
     * in conjunction with the `style_type_number` parameter:
     *
     * @code
     *     lxlsx_table_options options = {
     *         .style_type = LXLSX_TABLE_STYLE_TYPE_LIGHT,
     *         .style_type_number = 11,
     *     };
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:G8"), &options);
     * @endcode
     *
     *
     * @image html tables11.png
     *
     * There are three types of table style in Excel: Light, Medium and Dark
     * which are represented using the #lxlsx_table_style_type enum values:
     *
     * - #LXLSX_TABLE_STYLE_TYPE_LIGHT
     *
     * - #LXLSX_TABLE_STYLE_TYPE_MEDIUM
     *
     * - #LXLSX_TABLE_STYLE_TYPE_DARK
     *
     * Within those ranges there are between 11 and 28 other style types which
     * can be set with `style_type_number` (depending on the style type).
     * Check Excel to find the style that you want. The dialog with the
     * options laid out in numeric order are shown below:
     *
     * @image html tables14.png
     *
     * The default table style in Excel is 'Table Style Medium 9' (highlighted
     * with a green border in the image above), which is set by default in
     * libxlsxwriter as:
     *
     * @code
     *     lxlsx_table_options options = {
     *         .style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM,
     *         .style_type_number = 9,
     *     };
     * @endcode
     *
     * You can also turn the table style off by setting it to Light 0:
     *
     * @code
     *     lxlsx_table_options options = {
     *         .style_type = LXLSX_TABLE_STYLE_TYPE_LIGHT,
     *         .style_type_number = 0,
     *     };
     * @endcode
     *
     * @image html tables13.png
     *
     */
    uint8_t style_type;

    /**
     * The `style_type_number` parameter is used with `style_type` to set the
     * style of a worksheet table. */
    uint8_t style_type_number;

    /**
     * The `total_row` parameter can be used to turn on the total row in the
     * last row of a table. It is distinguished from the other rows by a
     * different formatting and also with dropdown `SUBTOTAL` functions:
     *
     * @code
     *     lxlsx_table_options options = {.total_row = LXLSX_TRUE};
     *
     *     lxlsx_worksheet_add_table(worksheet, RANGE("B3:G8"), &options);
     * @endcode
     *
     * @image html tables9.png
     *
     * The default total row doesn't have any captions or functions. These
     * must by specified via the `columns` parameter below.
     */
    uint8_t total_row;

    /**
     * The `columns` parameter can be used to set properties for columns
     * within the table. See @ref ww_tables_columns for a detailed
     * explanation.
     */
    lxlsx_table_column **columns;

} lxlsx_table_options;

typedef struct lxlsx_table_obj {
    char *name;
    char *total_string;
    lxlsx_table_column **columns;
    uint8_t banded_columns;
    uint8_t first_column;
    uint8_t last_column;
    uint8_t no_autofilter;
    uint8_t no_banded_rows;
    uint8_t no_header_row;
    uint8_t style_type;
    uint8_t style_type_number;
    uint8_t total_row;

    lxlsx_row_t first_row;
    lxlsx_col_t first_col;
    lxlsx_row_t last_row;
    lxlsx_col_t last_col;
    lxlsx_col_t num_cols;
    uint32_t id;

    char sqref[LXLSX_MAX_ATTRIBUTE_LENGTH];
    char filter_sqref[LXLSX_MAX_ATTRIBUTE_LENGTH];
    STAILQ_ENTRY (lxlsx_table_obj) list_pointers;

} lxlsx_table_obj;

/**
 * @brief Options for autofilter rules.
 *
 * Options to define an autofilter rule.
 *
 */
typedef struct lxlsx_filter_rule {

    /** The #lxlsx_filter_criteria to define the rule. */
    uint8_t criteria;

    /** String value to which the criteria applies. */
    const char *value_string;

    /** Numeric value to which the criteria applies (if value_string isn't used). */
    double value;

} lxlsx_filter_rule;

typedef struct lxlsx_filter_rule_obj {

    uint8_t type;
    uint8_t is_custom;
    uint8_t has_blanks;
    lxlsx_col_t col_num;

    uint8_t criteria1;
    uint8_t criteria2;
    double value1;
    double value2;
    char *value1_string;
    char *value2_string;

    uint16_t num_list_filters;
    char **list;

} lxlsx_filter_rule_obj;

/**
 * @brief Options for inserted images.
 *
 * Options for modifying images inserted via `lxlsx_worksheet_insert_image_opt()` and
 * `lxlsx_worksheet_embed_image_opt()`.
 */
typedef struct lxlsx_image_options {

    /** Offset from the left of the cell in pixels. */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. */
    int32_t y_offset;

    /** X scale of the image as a decimal. */
    double x_scale;

    /** Y scale of the image as a decimal. */
    double y_scale;

    /** Object position - use one of the values of #lxlsx_object_position.
     *  See @ref working_with_object_positioning.*/
    uint8_t object_position;

    /** Optional description or "Alt text" for the image. This field can be
     *  used to provide a text description of the image to help
     *  accessibility. Defaults to the image filename as in Excel. Set to ""
     *  to ignore the description field. */
    const char *description;

    /** Optional parameter to help accessibility. It is used to mark the image
     *  as decorative, and thus uninformative, for automated screen
     *  readers. As in Excel, if this parameter is in use the `description`
     *  field isn't written. */
    uint8_t decorative;

    /** Add an optional hyperlink to the image. Follows the same URL rules
     *  and types as `lxlsx_worksheet_write_url()`. */
    const char *url;

    /** Add an optional mouseover tip for a hyperlink to the image. */
    const char *tip;

    /** Add an optional format to the cell. Only used with
     * `lxlsx_worksheet_embed_image_opt()` */
    lxlsx_format *cell_format;

} lxlsx_image_options;

/**
 * @brief Options for inserted charts.
 *
 * Options for modifying charts inserted via `lxlsx_worksheet_insert_chart_opt()`.
 *
 */
typedef struct lxlsx_chart_options {

    /** Offset from the left of the cell in pixels. */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. */
    int32_t y_offset;

    /** X scale of the chart as a decimal. */
    double x_scale;

    /** Y scale of the chart as a decimal. */
    double y_scale;

    /** Object position - use one of the values of #lxlsx_object_position.
     *  See @ref working_with_object_positioning.*/
    uint8_t object_position;

    /** Optional description or "Alt text" for the chart. This field can be
     *  used to provide a text description of the chart to help
     *  accessibility. Defaults to the image filename as in Excel. Set to NULL
     *  to ignore the description field. */
    const char *description;

    /** Optional parameter to help accessibility. It is used to mark the chart
     *  as decorative, and thus uninformative, for automated screen
     *  readers. As in Excel, if this parameter is in use the `description`
     *  field isn't written. */
    uint8_t decorative;

} lxlsx_chart_options;

/* Internal struct to represent lxlsx_image_options and lxlsx_chart_options
 * values as well as internal metadata.
 */
typedef struct lxlsx_object_properties {
    int32_t x_offset;
    int32_t y_offset;
    double x_scale;
    double y_scale;
    lxlsx_row_t row;
    lxlsx_col_t col;
    char *filename;
    char *description;
    char *url;
    char *tip;
    uint8_t object_position;
    FILE *stream;
    uint8_t image_type;
    uint8_t is_image_buffer;
    char *image_buffer;
    size_t image_buffer_size;
    double width;
    double height;
    char *extension;
    double x_dpi;
    double y_dpi;
    lxlsx_chart *chart;
    uint8_t is_duplicate;
    uint8_t is_background;
    char *md5;
    char *image_position;
    uint8_t decorative;
    lxlsx_format *format;

    STAILQ_ENTRY (lxlsx_object_properties) list_pointers;
} lxlsx_object_properties;

/**
 * @brief Options for inserted comments.
 *
 * Options for modifying comments inserted via `lxlsx_worksheet_write_comment_opt()`.
 *
 */
typedef struct lxlsx_comment_options {

    /** This option is used to make a cell comment visible when the worksheet
     *  is opened. The default behavior in Excel is that comments are
     *  initially hidden. However, it is also possible in Excel to make
     *  individual comments or all comments visible.  You can make all
     *  comments in the worksheet visible using the
     *  `lxlsx_worksheet_show_comments()` function. Defaults to
     *  LXLSX_COMMENT_DISPLAY_DEFAULT. See also @ref ww_comments_visible. */
    uint8_t visible;

    /** This option is used to indicate the author of the cell comment. Excel
     *  displays the author in the status bar at the bottom of the
     *  worksheet. The default author for all cell comments in a worksheet can
     *  be set using the `lxlsx_worksheet_set_comments_author()` function. Set to
     *  NULL if not required.  See also @ref ww_comments_author. */
    const char *author;

    /** This option is used to set the width of the cell comment box
     *  explicitly in pixels. The default width is 128 pixels. See also @ref
     *  ww_comments_width. */
    uint16_t width;

    /** This option is used to set the height of the cell comment box
     *  explicitly in pixels. The default height is 74 pixels.  See also @ref
     *  ww_comments_height. */
    uint16_t height;

    /** X scale of the comment as a decimal. See also
     * @ref ww_comments_x_scale. */
    double x_scale;

    /** Y scale of the comment as a decimal. See also
     * @ref ww_comments_y_scale. */
    double y_scale;

    /** This option is used to set the background color of cell comment
     *  box. The color should be an RGB integer value, see @ref
     *  working_with_colors. See also @ref ww_comments_color. */
    lxlsx_color_t color;

    /** This option is used to set the font for the comment. The default font
     *  is 'Tahoma'.  See also @ref ww_comments_font_name. */
    const char *font_name;

     /** This option is used to set the font size for the comment. The default
      * is 8. See also @ref ww_comments_font_size. */
    double font_size;

    /** This option is used to set the font family number for the comment.
     *  Not required very often. Set to 0. */
    uint8_t font_family;

    /** This option is used to set the row in which the comment will
     *  appear. By default Excel displays comments one cell to the right and
     *  one cell above the cell to which the comment relates. The `start_row`
     *  and `start_col` options should both be set to 0 if not used.  See also
     *  @ref ww_comments_start_row. */
    lxlsx_row_t start_row;

    /** This option is used to set the column in which the comment will
     *   appear. See the `start_row` option for more information and see also
     *   @ref ww_comments_start_col. */
    lxlsx_col_t start_col;

    /** Offset from the left of the cell in pixels. See also
     * @ref ww_comments_x_offset. */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. See also
     * @ref ww_comments_y_offset. */
    int32_t y_offset;

} lxlsx_comment_options;

/**
 * @brief Options for inserted buttons.
 *
 * Options for modifying buttons inserted via `lxlsx_worksheet_insert_button()`.
 *
 */
typedef struct lxlsx_button_options {

    /** Sets the caption on the button. The default is "Button n" where n is
     *  the current number of buttons in the worksheet, including this
     *  button. */
    const char *caption;

    /** Name of the macro to run when the button is pressed. The macro must be
     *  included with lxlsx_workbook_add_vba_project(). */
    const char *macro;

    /** Optional description or "Alt text" for the button. This field can be
     *  used to provide a text description of the button to help
     *  accessibility. Set to NULL to ignore the description field. */
    const char *description;

    /** This option is used to set the width of the cell button box
     *  explicitly in pixels. The default width is 64 pixels. */
    uint16_t width;

    /** This option is used to set the height of the cell button box
     *  explicitly in pixels. The default height is 20 pixels. */
    uint16_t height;

    /** X scale of the button as a decimal. */
    double x_scale;

    /** Y scale of the button as a decimal. */
    double y_scale;

    /** Offset from the left of the cell in pixels.  */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. */
    int32_t y_offset;

} lxlsx_button_options;

/* Internal structure for VML object options. */
typedef struct lxlsx_vml_obj {

    lxlsx_row_t row;
    lxlsx_col_t col;
    lxlsx_row_t start_row;
    lxlsx_col_t start_col;
    int32_t x_offset;
    int32_t y_offset;
    uint64_t col_absolute;
    uint64_t row_absolute;
    uint32_t width;
    uint32_t height;
    double x_dpi;
    double y_dpi;
    lxlsx_color_t color;
    uint8_t font_family;
    uint8_t visible;
    uint32_t author_id;
    uint32_t rel_index;
    double font_size;
    struct lxlsx_drawing_coords from;
    struct lxlsx_drawing_coords to;
    char *author;
    char *font_name;
    char *text;
    char *image_position;
    char *name;
    char *macro;
    STAILQ_ENTRY (lxlsx_vml_obj) list_pointers;

} lxlsx_vml_obj;

/**
 * @brief Header and footer options.
 *
 * Optional parameters used in the `lxlsx_worksheet_set_header_opt()` and
 * lxlsx_worksheet_set_footer_opt() functions.
 *
 */
typedef struct lxlsx_header_footer_options {
    /** Header or footer margin in inches. Excel default is 0.3. Must by
     *  larger than 0.0.  See `lxlsx_worksheet_set_header_opt()`. */
    double margin;

    /** The left header image filename, with path if required. This should
     * have a corresponding `&G/&[Picture]` placeholder in the `&L` section of
     * the header/footer string. See `lxlsx_worksheet_set_header_opt()`. */
    const char *image_left;

    /** The center header image filename, with path if required. This should
     * have a corresponding `&G/&[Picture]` placeholder in the `&C` section of
     * the header/footer string. See `lxlsx_worksheet_set_header_opt()`. */
    const char *image_center;

    /** The right header image filename, with path if required. This should
     * have a corresponding `&G/&[Picture]` placeholder in the `&R` section of
     * the header/footer string. See `lxlsx_worksheet_set_header_opt()`. */
    const char *image_right;

} lxlsx_header_footer_options;

/**
 * @brief Worksheet protection options.
 */
typedef struct lxlsx_protection {
    /** Turn off selection of locked cells. This in on in Excel by default.*/
    uint8_t no_select_locked_cells;

    /** Turn off selection of unlocked cells. This in on in Excel by default.*/
    uint8_t no_select_unlocked_cells;

    /** Prevent formatting of cells. */
    uint8_t lxlsx_format_cells;

    /** Prevent formatting of columns. */
    uint8_t lxlsx_format_columns;

    /** Prevent formatting of rows. */
    uint8_t lxlsx_format_rows;

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

    /** Protect drawing objects. Worksheets only. */
    uint8_t objects;

    /** Turn off chartsheet content protection. */
    uint8_t no_content;

    /** Turn off chartsheet objects. */
    uint8_t no_objects;

} lxlsx_protection;

/* Internal struct to copy lxlsx_protection options and internal metadata. */
typedef struct lxlsx_protection_obj {
    uint8_t no_select_locked_cells;
    uint8_t no_select_unlocked_cells;
    uint8_t lxlsx_format_cells;
    uint8_t lxlsx_format_columns;
    uint8_t lxlsx_format_rows;
    uint8_t insert_columns;
    uint8_t insert_rows;
    uint8_t insert_hyperlinks;
    uint8_t delete_columns;
    uint8_t delete_rows;
    uint8_t sort;
    uint8_t autofilter;
    uint8_t pivot_tables;
    uint8_t scenarios;
    uint8_t objects;
    uint8_t no_content;
    uint8_t no_objects;
    uint8_t no_sheet;
    uint8_t is_configured;
    char hash[5];
} lxlsx_protection_obj;

/**
 * @brief Struct to represent a rich string format/string pair.
 *
 * Arrays of this struct are used to define "rich" multi-format strings that
 * are passed to `lxlsx_worksheet_write_rich_string()`. Each struct represents a
 * fragment of the rich multi-format string with a lxlsx_format to define the
 * format for the string part. If the string fragment is unformatted then
 * `NULL` can be used for the format.
 */
typedef struct lxlsx_rich_string_tuple {

    /** The format for a string fragment in a rich string. NULL if the string
     *  isn't formatted. */
    lxlsx_format *format;

    /** The string fragment. */
    const char *string;
} lxlsx_rich_string_tuple;

/**
 * @brief Struct to represent an Excel worksheet.
 *
 * The members of the lxlsx_worksheet struct aren't modified directly. Instead
 * the worksheet properties are set by calling the functions shown in
 * worksheet.h.
 */
typedef struct lxlsx_worksheet {

    FILE *file;
    FILE *optimize_tmpfile;
    char *optimize_buffer;
    size_t optimize_buffer_size;
    struct lxlsx_table_rows *table;
    struct lxlsx_table_rows *hyperlinks;
    struct lxlsx_table_rows *comments;
    struct lxlsx_cell **array;
    struct lxlsx_merged_ranges *merged_ranges;
    struct lxlsx_selections *selections;
    struct lxlsx_data_validations *data_validations;
    struct lxlsx_cond_format_hash *conditional_formats;
    struct lxlsx_image_props *image_props;
    struct lxlsx_image_props *embedded_image_props;
    struct lxlsx_chart_props *lxlsx_chart_data;
    struct lxlsx_drawing_rel_ids *lxlsx_drawing_rel_ids;
    struct lxlsx_vml_drawing_rel_ids *lxlsx_vml_drawing_rel_ids;
    struct lxlsx_comment_objs *comment_objs;
    struct lxlsx_comment_objs *header_image_objs;
    struct lxlsx_comment_objs *button_objs;
    struct lxlsx_table_objs *lxlsx_table_objs;
    uint16_t lxlsx_table_count;

    lxlsx_row_t dim_rowmin;
    lxlsx_row_t dim_rowmax;
    lxlsx_col_t dim_colmin;
    lxlsx_col_t dim_colmax;

    lxlsx_sst *sst;
    const char *name;
    const char *quoted_name;
    lxlsx_edit_session *edit_session;
    char *edit_sheet_name;
    uint8_t is_edit;
    const char *tmpdir;

    uint16_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    uint8_t is_chartsheet;

    lxlsx_col_options **col_options;
    uint16_t col_options_max;

    double *col_sizes;
    uint16_t col_sizes_max;

    lxlsx_format **col_formats;
    uint16_t col_formats_max;

    uint8_t col_size_changed;
    uint8_t row_size_changed;
    uint8_t optimize;
    struct lxlsx_row *optimize_row;

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
    uint8_t vcenter;
    uint8_t zoom_scale_normal;
    uint8_t black_white;
    uint8_t num_validations;
    uint8_t has_dynamic_functions;
    char *vba_codename;
    uint16_t num_buttons;

    lxlsx_color_t tab_color;

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
    char *header;
    char *footer;

    struct lxlsx_repeat_rows repeat_rows;
    struct lxlsx_repeat_cols repeat_cols;
    struct lxlsx_print_area print_area;
    struct lxlsx_autofilter autofilter;

    uint16_t merged_range_count;
    uint16_t max_url_length;

    lxlsx_row_t *hbreaks;
    lxlsx_col_t *vbreaks;
    uint16_t hbreaks_count;
    uint16_t vbreaks_count;

    uint32_t lxlsx_drawing_rel_id;
    uint32_t lxlsx_vml_drawing_rel_id;
    struct lxlsx_rel_tuples *external_hyperlinks;
    struct lxlsx_rel_tuples *external_drawing_links;
    struct lxlsx_rel_tuples *lxlsx_drawing_links;
    struct lxlsx_rel_tuples *lxlsx_vml_drawing_links;
    struct lxlsx_rel_tuples *external_table_links;

    struct lxlsx_panes panes;
    char top_left_cell[LXLSX_MAX_CELL_NAME_LENGTH];

    struct lxlsx_protection_obj protection;

    lxlsx_drawing *drawing;
    lxlsx_format *default_url_format;

    uint8_t has_vml;
    uint8_t has_comments;
    uint8_t has_header_vml;
    uint8_t has_background_image;
    uint8_t has_buttons;
    uint8_t storing_embedded_image;
    lxlsx_rel_tuple *external_vml_comment_link;
    lxlsx_rel_tuple *external_comment_link;
    lxlsx_rel_tuple *external_vml_header_link;
    lxlsx_rel_tuple *external_background_link;
    char *comment_author;
    char *lxlsx_vml_data_id_str;
    char *lxlsx_vml_header_id_str;
    uint32_t lxlsx_vml_shape_id;
    uint32_t lxlsx_vml_header_id;
    uint32_t dxf_priority;
    uint8_t comment_display_default;
    uint32_t data_bar_2010_index;

    uint8_t has_ignore_errors;
    char *ignore_number_stored_as_text;
    char *ignore_eval_error;
    char *ignore_formula_differs;
    char *ignore_formula_range;
    char *ignore_formula_unlocked;
    char *ignore_empty_cell_reference;
    char *ignore_list_data_validation;
    char *ignore_calculated_column;
    char *ignore_two_digit_text_year;

    uint8_t use_1904_epoch;

    uint16_t excel_version;

    lxlsx_object_properties **header_footer_objs[LXLSX_HEADER_FOOTER_OBJS_MAX];
    lxlsx_object_properties *header_left_object_props;
    lxlsx_object_properties *header_center_object_props;
    lxlsx_object_properties *header_right_object_props;
    lxlsx_object_properties *footer_left_object_props;
    lxlsx_object_properties *footer_center_object_props;
    lxlsx_object_properties *footer_right_object_props;
    lxlsx_object_properties *background_image;

    lxlsx_filter_rule_obj **filter_rules;
    lxlsx_col_t num_filter_rules;

    STAILQ_ENTRY (lxlsx_worksheet) list_pointers;

} lxlsx_worksheet;

/*
 * Worksheet initialization data.
 */
typedef struct lxlsx_worksheet_init_data {
    uint16_t index;
    uint8_t hidden;
    uint8_t optimize;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    lxlsx_sst *sst;
    const char *name;
    const char *quoted_name;
    const char *tmpdir;
    lxlsx_format *default_url_format;
    uint16_t max_url_length;
    uint8_t use_1904_epoch;

} lxlsx_worksheet_init_data;

/* Struct to represent a worksheet row. */
typedef struct lxlsx_row {
    lxlsx_row_t row_num;
    double height;
    lxlsx_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
    uint8_t row_changed;
    uint8_t data_changed;
    uint8_t height_changed;

    struct lxlsx_table_cells *cells;

    /* tree management pointers for tree.h. */
    RB_ENTRY (lxlsx_row) tree_pointers;
} lxlsx_row;

/* Struct to represent a drawing Target/ID pair. */
typedef struct lxlsx_drawing_rel_id {
    uint32_t id;
    char *target;

    RB_ENTRY (lxlsx_drawing_rel_id) tree_pointers;
} lxlsx_drawing_rel_id;



/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Write a number to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param number    The number to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `lxlsx_worksheet_write_number()` function writes numeric types to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     lxlsx_worksheet_write_number(worksheet, 0, 0, 123456, NULL);
 *     lxlsx_worksheet_write_number(worksheet, 1, 0, 2.3451, NULL);
 * @endcode
 *
 * @image html write_number01.png
 *
 * The native data type for all numbers in Excel is a IEEE-754 64-bit
 * double-precision floating point, which is also the default type used by
 * `%lxlsx_worksheet_write_number`.
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * @code
 *     lxlsx_format *format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_num_format(format, "$#,##0.00");
 *
 *     lxlsx_worksheet_write_number(worksheet, 0, 0, 1234.567, format);
 * @endcode
 *
 * @image html write_number02.png
 *
 * @note Excel doesn't support `NaN`, `Inf` or `-Inf` as a number value. If
 * you are writing data that contains these values then your application
 * should convert them to a string or handle them in some other way.
 *
 */
lxlsx_error lxlsx_worksheet_write_number(lxlsx_worksheet *worksheet,
                                 lxlsx_row_t row,
                                 lxlsx_col_t col, double number,
                                 lxlsx_format *format);
/**
 * @brief Write a string to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param string    String to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_string()` function writes a string to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "This phrase is English!", NULL);
 * @endcode
 *
 * @image html write_string01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object:
 *
 * @code
 *     lxlsx_format *format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_bold(format);
 *
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "This phrase is Bold!", format);
 * @endcode
 *
 * @image html write_string02.png
 *
 * Unicode strings are supported in UTF-8 encoding. This generally requires
 * that your source file is UTF-8 encoded or that the data has been read from
 * a UTF-8 source:
 *
 * @code
 *    lxlsx_worksheet_write_string(worksheet, 0, 0, "Это фраза на русском!", NULL);
 * @endcode
 *
 * @image html write_string03.png
 *
 */
lxlsx_error lxlsx_worksheet_write_string(lxlsx_worksheet *worksheet,
                                 lxlsx_row_t row,
                                 lxlsx_col_t col, const char *string,
                                 lxlsx_format *format);
/**
 * @brief Write a formula to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_formula()` function writes a formula or function to
 * the cell specified by `row` and `column`:
 *
 * @code
 *  lxlsx_worksheet_write_formula(worksheet, 0, 0, "=B3 + 6",                    NULL);
 *  lxlsx_worksheet_write_formula(worksheet, 1, 0, "=SIN(PI()/4)",               NULL);
 *  lxlsx_worksheet_write_formula(worksheet, 2, 0, "=SUM(A1:A2)",                NULL);
 *  lxlsx_worksheet_write_formula(worksheet, 3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL);
 *  lxlsx_worksheet_write_formula(worksheet, 4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL);
 *  lxlsx_worksheet_write_formula(worksheet, 5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL);
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
 * `lxlsx_worksheet_write_formula_num()` function and the discussion in that section.
 *
 * Formulas must be written with the US style separator/range operator which
 * is a comma (not semi-colon). Therefore a formula with multiple values
 * should be written as follows:
 *
 * @code
 *     // OK.
 *     lxlsx_worksheet_write_formula(worksheet, 0, 0, "=SUM(1, 2, 3)", NULL);
 *
 *     // NO. Error on load.
 *     lxlsx_worksheet_write_formula(worksheet, 1, 0, "=SUM(1; 2; 3)", NULL);
 * @endcode
 *
 * See also @ref working_with_formulas.
 */
lxlsx_error lxlsx_worksheet_write_formula(lxlsx_worksheet *worksheet,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col, const char *formula,
                                  lxlsx_format *format);
/**
 * @brief Write an array formula to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param formula   Array formula to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_array_formula()` function writes an array formula to
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
 *     lxlsx_worksheet_write_array_formula(worksheet, 4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL);
 *
 *     // Same as above using the RANGE() macro.
 *     lxlsx_worksheet_write_array_formula(worksheet, RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL);
 * @endcode
 *
 * If the array formula returns a single value then the `first_` and `last_`
 * parameters should be the same:
 *
 * @code
 *     lxlsx_worksheet_write_array_formula(worksheet, 1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL);
 *     lxlsx_worksheet_write_array_formula(worksheet, RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL);
 * @endcode
 *
 */
lxlsx_error lxlsx_worksheet_write_array_formula(lxlsx_worksheet *worksheet,
                                        lxlsx_row_t first_row,
                                        lxlsx_col_t first_col,
                                        lxlsx_row_t last_row,
                                        lxlsx_col_t last_col,
                                        const char *formula,
                                        lxlsx_format *format);

/**
 * @brief Write an Excel 365 dynamic array formula to a worksheet range.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param formula   Dynamic Array formula to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 *
 * The `%lxlsx_worksheet_write_dynamic_array_formula()` function writes an Excel 365
 * dynamic array formula to a cell range. Some examples of functions that
 * return dynamic arrays are:
 *
 * - `FILTER`
 * - `RANDARRAY`
 * - `SEQUENCE`
 * - `SORTBY`
 * - `SORT`
 * - `UNIQUE`
 * - `XLOOKUP`
 * - `XMATCH`
 *
 * Dynamic array formulas and their usage in libxlsxwriter is explained in
 * detail @ref ww_formulas_dynamic_arrays. The following is a example usage:
 *
 * @code
 *     lxlsx_worksheet_write_dynamic_array_formula(worksheet, 1, 5, 1, 5,
 *                                           "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)",
 *                                           NULL);
 * @endcode
 *
 * This formula gives the results shown in the image below.
 *
 * @image html dynamic_arrays02.png
 *
 * The need for the `_xlfn._xlws.` prefix in the formula is explained in @ref
 * ww_formulas_future.
 */
lxlsx_error lxlsx_worksheet_write_dynamic_array_formula(lxlsx_worksheet *worksheet,
                                                lxlsx_row_t first_row,
                                                lxlsx_col_t first_col,
                                                lxlsx_row_t last_row,
                                                lxlsx_col_t last_col,
                                                const char *formula,
                                                lxlsx_format *format);

/**
 * @brief Write an Excel 365 dynamic array formula to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_dynamic_formula()` function is similar to the
 * `lxlsx_worksheet_write_dynamic_array_formula()` function, shown above, except
 * that it writes a dynamic array formula to a single cell, rather than a
 * range. This is a syntactic shortcut since the array range isn't generally
 * known for a dynamic range and specifying the initial cell is sufficient for
 * Excel, as shown in the example below:
 *
 * @code
 *     lxlsx_worksheet_write_dynamic_formula(worksheet, 7, 1,
 *                                     "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))",
 *                                     NULL);
 * @endcode
 *
 * This formula gives the following result:
 *
 * @image html dynamic_arrays01.png
 *
 * The need for the `_xlfn.` and `_xlfn._xlws.` prefixes in the formula is
 * explained in @ref ww_formulas_future.
 */
lxlsx_error lxlsx_worksheet_write_dynamic_formula(lxlsx_worksheet *worksheet,
                                          lxlsx_row_t row,
                                          lxlsx_col_t col,
                                          const char *formula,
                                          lxlsx_format *format);

lxlsx_error lxlsx_worksheet_write_array_formula_num(lxlsx_worksheet *worksheet,
                                            lxlsx_row_t first_row,
                                            lxlsx_col_t first_col,
                                            lxlsx_row_t last_row,
                                            lxlsx_col_t last_col,
                                            const char *formula,
                                            lxlsx_format *format,
                                            double result);

lxlsx_error lxlsx_worksheet_write_dynamic_array_formula_num(lxlsx_worksheet *worksheet,
                                                    lxlsx_row_t first_row,
                                                    lxlsx_col_t first_col,
                                                    lxlsx_row_t last_row,
                                                    lxlsx_col_t last_col,
                                                    const char *formula,
                                                    lxlsx_format *format,
                                                    double result);

lxlsx_error lxlsx_worksheet_write_dynamic_formula_num(lxlsx_worksheet *worksheet,
                                              lxlsx_row_t row,
                                              lxlsx_col_t col,
                                              const char *formula,
                                              lxlsx_format *format,
                                              double result);

/**
 * @brief Write a date or time to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param datetime  The datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_datetime()` function can be used to write a date or
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
lxlsx_error lxlsx_worksheet_write_datetime(lxlsx_worksheet *worksheet,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col, lxlsx_datetime *datetime,
                                   lxlsx_format *format);

/**
 * @brief Write a Unix datetime to a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param unixtime  The Unix datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_unixtime()` function can be used to write dates and
 * times in Unix date format to the cell specified by `row` and
 * `column`. [Unix Time](https://en.wikipedia.org/wiki/Unix_time) which is a
 * common integer time format. It is defined as the number of seconds since
 * the Unix epoch (1970-01-01 00:00 UTC). Negative values can also be used for
 * dates prior to 1970:
 *
 * @dontinclude dates_and_times03.c
 * @skip 1970
 * @until 2208988800
 *
 * The `format` parameter should be used to apply formatting to the cell using
 * a @ref format.h "Format" object as shown above. Without a date format the
 * datetime will appear as a number only.
 *
 * The output from this code sample is:
 *
 * @image html date_example03.png
 *
 * Unixtime is generally represented with a 32 bit `time_t` type which has a
 * range of approximately 1900-12-14 to 2038-01-19. To access the full Excel
 * date range of 1900-01-01 to 9999-12-31 this function uses a 64 bit
 * parameter.
 *
 * See @ref working_with_dates for more information about handling dates and
 * times in libxlsxwriter.
 */
lxlsx_error lxlsx_worksheet_write_unixtime(lxlsx_worksheet *worksheet,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col, int64_t unixtime,
                                   lxlsx_format *format);

/**
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param url       The url to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 *
 * The `%lxlsx_worksheet_write_url()` function is used to write a URL/hyperlink to a
 * worksheet cell specified by `row` and `column`.
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", NULL);
 * @endcode
 *
 * @image html hyperlinks_short.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL`, in which case the default Excel blue underlined
 * hyperlink style will be used. If required a user defined @ref format.h
 * "Format" object can be used:
 * underline:
 *
 * @code
 *    lxlsx_format *url_format   = lxlsx_workbook_add_format(workbook);
 *
 *    lxlsx_format_set_underline (url_format, LXLSX_UNDERLINE_SINGLE);
 *    lxlsx_format_set_font_color(url_format, LXLSX_COLOR_RED);
 *
 * @endcode
 *
 * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
 * and `mailto:` :
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "ftp://www.python.org/",     NULL);
 *     lxlsx_worksheet_write_url(worksheet, 1, 0, "http://www.python.org/",    NULL);
 *     lxlsx_worksheet_write_url(worksheet, 2, 0, "https://www.python.org/",   NULL);
 *     lxlsx_worksheet_write_url(worksheet, 3, 0, "mailto:jmcnamara@cpan.org", NULL);
 *
 * @endcode
 *
 * An Excel hyperlink is comprised of two elements: the displayed string and
 * the non-displayed link. By default the displayed string is the same as the
 * link. However, it is possible to overwrite it with any other
 * `libxlsxwriter` type using the appropriate `lxlsx_worksheet_write_*()`
 * function. The most common case is to overwrite the displayed link text with
 * another string. To do this we must also match the default URL format using
 * `lxlsx_workbook_get_default_url_format()`:
 *
 * @code
 *     // Write a hyperlink with the default blue underline format.
 *     lxlsx_worksheet_write_url(worksheet, 2, 0, "http://libxlsxwriter.github.io", NULL);
 *
 *     // Get the default url format.
 *     lxlsx_format *url_format = lxlsx_workbook_get_default_url_format(workbook);
 *
 *     // Overwrite the hyperlink with a user defined string and default format.
 *     lxlsx_worksheet_write_string(worksheet, 2, 0, "Read the documentation.", url_format);
 * @endcode
 *
 * @image html hyperlinks_short2.png
 *
 * Two local URIs are supported: `internal:` and `external:`. These are used
 * for hyperlinks to internal worksheet references or external workbook and
 * worksheet references:
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "internal:Sheet2!A1",                NULL);
 *     lxlsx_worksheet_write_url(worksheet, 1, 0, "internal:Sheet2!B2",                NULL);
 *     lxlsx_worksheet_write_url(worksheet, 2, 0, "internal:Sheet2!A1:B2",             NULL);
 *     lxlsx_worksheet_write_url(worksheet, 3, 0, "internal:'Sales Data'!A1",          NULL);
 *     lxlsx_worksheet_write_url(worksheet, 4, 0, "external:c:\\temp\\foo.xlsx",       NULL);
 *     lxlsx_worksheet_write_url(worksheet, 5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   NULL);
 *     lxlsx_worksheet_write_url(worksheet, 6, 0, "external:..\\foo.xlsx",             NULL);
 *     lxlsx_worksheet_write_url(worksheet, 7, 0, "external:..\\foo.xlsx#Sheet2!A1",   NULL);
 *     lxlsx_worksheet_write_url(worksheet, 8, 0, "external:\\\\NET\\share\\foo.xlsx", NULL);
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
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   NULL);
 * @endcode
 *
 * You can also link to a named range in the target worksheet: For example say
 * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
 * you could link to it as follows:
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "external:c:\\temp\\foo.xlsx#my_name", NULL);
 *
 * @endcode
 *
 * Excel requires that worksheet names containing spaces or non alphanumeric
 * characters are single quoted as follows:
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "internal:'Sales Data'!A1", NULL);
 * @endcode
 *
 * Links to network files are also supported. Network files normally begin
 * with two back slashes as follows `\\NETWORK\etc`. In order to represent
 * this in a C string literal the backslashes should be escaped:
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "external:\\\\NET\\share\\foo.xlsx", NULL);
 * @endcode
 *
 *
 * Alternatively, you can use Unix style forward slashes. These are
 * translated internally to backslashes:
 *
 * @code
 *     lxlsx_worksheet_write_url(worksheet, 0, 0, "external:c:/temp/foo.xlsx",     NULL);
 *     lxlsx_worksheet_write_url(worksheet, 1, 0, "external://NET/share/foo.xlsx", NULL);
 *
 * @endcode
 *
 *
 * **Note:**
 *
 *    libxlsxwriter will escape the following characters in URLs as required
 *    by Excel: `\s " < > \ [ ]  ^ { }`. Existing URL `%%xx` style escapes in
 *    the string are ignored to allow for user-escaped strings.
 *
 * **Note:**
 *
 *    The maximum allowable URL length in recent versions of Excel is 2079
 *    characters. In older versions of Excel (and libxlsxwriter <= 0.8.8) the
 *    limit was 255 characters.
 */
lxlsx_error lxlsx_worksheet_write_url(lxlsx_worksheet *worksheet,
                              lxlsx_row_t row,
                              lxlsx_col_t col, const char *url,
                              lxlsx_format *format);

 /* Don't document for now since the string option can be achieved by a
  * subsequent cell lxlsx_worksheet_write() as shown in the docs, and the
  * tooltip option isn't very useful. */
lxlsx_error lxlsx_worksheet_write_url_opt(lxlsx_worksheet *worksheet,
                                  lxlsx_row_t row_num,
                                  lxlsx_col_t col_num, const char *url,
                                  lxlsx_format *format, const char *string,
                                  const char *tooltip);

/**
 * @brief Write a formatted boolean worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param value     The boolean value to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * Write an Excel boolean to the cell specified by `row` and `column`:
 *
 * @code
 *     lxlsx_worksheet_write_boolean(worksheet, 2, 2, 0, my_format);
 * @endcode
 *
 */
lxlsx_error lxlsx_worksheet_write_boolean(lxlsx_worksheet *worksheet,
                                  lxlsx_row_t row, lxlsx_col_t col,
                                  int value, lxlsx_format *format);

/**
 * @brief Write a formatted blank worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * Write a blank cell specified by `row` and `column`:
 *
 * @code
 *     lxlsx_worksheet_write_blank(worksheet, 1, 1, border_format);
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
lxlsx_error lxlsx_worksheet_write_blank(lxlsx_worksheet *worksheet,
                                lxlsx_row_t row, lxlsx_col_t col,
                                lxlsx_format *format);

/**
 * @brief Write a formula to a worksheet cell with a user defined numeric
 *        result.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined numeric result for the formula.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_formula_num()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * numeric result:
 *
 * @code
 *     // Required as a workaround only.
 *     lxlsx_worksheet_write_formula_num(worksheet, 0, 0, "=1 + 2", NULL, 3);
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
 * If required, the `%lxlsx_worksheet_write_formula_num()` function can be used to
 * specify a formula and its result.
 *
 * This function is rarely required and is only provided for compatibility
 * with some third party applications. For most applications the
 * lxlsx_worksheet_write_formula() function is the recommended way of writing
 * formulas.
 *
 * See also @ref working_with_formulas.
 */
lxlsx_error lxlsx_worksheet_write_formula_num(lxlsx_worksheet *worksheet,
                                      lxlsx_row_t row,
                                      lxlsx_col_t col,
                                      const char *formula,
                                      lxlsx_format *format, double result);

/**
 * @brief Write a formula to a worksheet cell with a user defined string
 *        result.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined string result for the formula.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_formula_str()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * string result:
 *
 * @code
 *     // The example formula is A & B -> AB.
 *     lxlsx_worksheet_write_formula_str(worksheet, 0, 0, "=\"A\" & \"B\"", NULL, "AB");
 * @endcode
 *
 * The `%lxlsx_worksheet_write_formula_str()` function is similar to the
 * `%lxlsx_worksheet_write_formula_num()` function except it writes a string result
 * instead or a numeric result. See `lxlsx_worksheet_write_formula_num()`  for more
 * details on why/when these functions are required.
 *
 * One place where the `%lxlsx_worksheet_write_formula_str()` function may be required
 * is to specify an empty result which will force a recalculation of the formula
 * when loaded in LibreOffice.
 *
 * @code
 *     lxlsx_worksheet_write_formula_str(worksheet, 0, 0, "=Sheet1!$A$1", NULL, "");
 * @endcode
 *
 * See the FAQ @ref faq_formula_zero.
 *
 * See also @ref working_with_formulas.
 */
lxlsx_error lxlsx_worksheet_write_formula_str(lxlsx_worksheet *worksheet,
                                      lxlsx_row_t row,
                                      lxlsx_col_t col,
                                      const char *formula,
                                      lxlsx_format *format, const char *result);

/**
 * @brief Write a "Rich" multi-format string to a worksheet cell.
 *
 * @param worksheet   Pointer to a lxlsx_worksheet instance to be updated.
 * @param row         The zero indexed row number.
 * @param col         The zero indexed column number.
 * @param rich_string An array of format/string lxlsx_rich_string_tuple fragments.
 * @param format      A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_rich_string()` function is used to write strings with
 * multiple formats. For example to write the string 'This is **bold**
 * and this is *italic*' you would use the following:
 *
 * @code
 *     lxlsx_format *bold = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_bold(bold);
 *
 *     lxlsx_format *italic = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_italic(italic);
 *
 *     lxlsx_rich_string_tuple fragment11 = {.format = NULL,   .string = "This is "     };
 *     lxlsx_rich_string_tuple fragment12 = {.format = bold,   .string = "bold"         };
 *     lxlsx_rich_string_tuple fragment13 = {.format = NULL,   .string = " and this is "};
 *     lxlsx_rich_string_tuple fragment14 = {.format = italic, .string = "italic"       };
 *
 *     lxlsx_rich_string_tuple *rich_string1[] = {&fragment11, &fragment12,
 *                                              &fragment13, &fragment14, NULL};
 *
 *     lxlsx_worksheet_write_rich_string(worksheet, CELL("A1"), rich_string1, NULL);
 *
 * @endcode
 *
 * @image html rich_strings_small.png
 *
 * The basic rule is to break the string into fragments and put a lxlsx_format
 * object before the fragment that you want to format. So if we look at the
 * above example again:
 *
 * This is **bold** and this is *italic*
 *
 * The would be broken down into 4 fragments:
 *
 *      default: |This is |
 *      bold:    |bold|
 *      default: | and this is |
 *      italic:  |italic|
 *
 * This in then converted to the lxlsx_rich_string_tuple fragments shown in the
 * example above. For the default format we use `NULL`.
 *
 * The fragments are passed to `%lxlsx_worksheet_write_rich_string()` as a `NULL`
 * terminated array:
 *
 * @code
 *     lxlsx_rich_string_tuple *rich_string1[] = {&fragment11, &fragment12,
 *                                              &fragment13, &fragment14, NULL};
 *
 *     lxlsx_worksheet_write_rich_string(worksheet, CELL("A1"), rich_string1, NULL);
 *
 * @endcode
 *
 * **Note**:
 * Excel doesn't allow the use of two consecutive formats in a rich string or
 * an empty string fragment. For either of these conditions a warning is
 * raised and the input to `%lxlsx_worksheet_write_rich_string()` is ignored.
 *
 */
lxlsx_error lxlsx_worksheet_write_rich_string(lxlsx_worksheet *worksheet,
                                      lxlsx_row_t row,
                                      lxlsx_col_t col,
                                      lxlsx_rich_string_tuple *rich_string[],
                                      lxlsx_format *format);

/**
 * @brief Write a comment to a worksheet cell.
 *
 * @param worksheet   Pointer to a lxlsx_worksheet instance to be updated.
 * @param row         The zero indexed row number.
 * @param col         The zero indexed column number.
 * @param string      The comment string to be written.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_comment()` function is used to add a comment to a
 * cell. A comment is indicated in Excel by a small red triangle in the upper
 * right-hand corner of the cell. Moving the cursor over the red triangle will
 * reveal the comment.
 *
 * The following example shows how to add a comment to a cell:
 *
 * @code
 *     lxlsx_worksheet_write_comment(worksheet, 0, 0, "This is a comment");
 * @endcode
 *
 * @image html comments1.png
 *
 * See also @ref working_with_comments
 *
 */
lxlsx_error lxlsx_worksheet_write_comment(lxlsx_worksheet *worksheet,
                                  lxlsx_row_t row, lxlsx_col_t col,
                                  const char *string);

/**
 * @brief Write a comment to a worksheet cell with options.
 *
 * @param worksheet   Pointer to a lxlsx_worksheet instance to be updated.
 * @param row         The zero indexed row number.
 * @param col         The zero indexed column number.
 * @param string      The comment string to be written.
 * @param options     #lxlsx_comment_options to control position and format
 *                    of the comment.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_write_comment_opt()` function is used to add a comment to a
 * cell with option that control the position, format and metadata of the
 * comment. A comment is indicated in Excel by a small red triangle in the
 * upper right-hand corner of the cell. Moving the cursor over the red
 * triangle will reveal the comment.
 *
 * The following example shows how to add a comment to a cell with options:
 *
 * @code
 *     lxlsx_comment_options options = {.visible = LXLSX_COMMENT_DISPLAY_VISIBLE};
 *
 *     lxlsx_worksheet_write_comment_opt(worksheet, CELL("C6"), "Hello.", &options);
 * @endcode
 *
 * The following options are available in #lxlsx_comment_options:
 *
 * - `author`
 * - `visible`
 * - `width`
 * - `height`
 * - `x_scale`
 * - `y_scale`
 * - `color`
 * - `font_name`
 * - `font_size`
 * - `start_row`
 * - `start_col`
 * - `x_offset`
 * - `y_offset`
 *
 * @image html comments2.png
 *
 * Comment options are explained in detail in the @ref ww_comments_properties
 * section of the docs.
 */
lxlsx_error lxlsx_worksheet_write_comment_opt(lxlsx_worksheet *worksheet,
                                      lxlsx_row_t row, lxlsx_col_t col,
                                      const char *string,
                                      lxlsx_comment_options *options);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height, in character units.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_row()` function is used to change the default
 * properties of a row. The most common use for this function is to change the
 * height of a row:
 *
 * @code
 *     // Set the height of Row 1 to 20.
 *     lxlsx_worksheet_set_row(worksheet, 0, 20, NULL);
 * @endcode
 *
 * The height is specified in character units. To specify the height in pixels
 * use the `lxlsx_worksheet_set_row_pixels()` function.
 *
 * The other common use for `%lxlsx_worksheet_set_row()` is to set the a @ref
 * format.h "Format" for all cells in the row:
 *
 * @code
 *     lxlsx_format *bold = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_bold(bold);
 *
 *     // Set the header row to bold.
 *     lxlsx_worksheet_set_row(worksheet, 0, 15, bold);
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height of #LXLSX_DEF_ROW_HEIGHT = 15:
 *
 * @code
 *     lxlsx_worksheet_set_row(worksheet, 0, LXLSX_DEF_ROW_HEIGHT, format);
 *     lxlsx_worksheet_set_row(worksheet, 0, 15, format); // Same as above.
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the row that don't
 * have a format. As with Excel the row format is overridden by an explicit
 * cell format. For example:
 *
 * @code
 *     // Row 1 has format1.
 *     lxlsx_worksheet_set_row(worksheet, 0, 15, format1);
 *
 *     // Cell A1 in Row 1 defaults to format1.
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell B1 in Row 1 keeps format2.
 *     lxlsx_worksheet_write_string(worksheet, 0, 1, "Hello", format2);
 * @endcode
 *
 */
lxlsx_error lxlsx_worksheet_set_row(lxlsx_worksheet *worksheet,
                            lxlsx_row_t row, double height, lxlsx_format *format);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_row_opt()` function  is the same as
 *  `lxlsx_worksheet_set_row()` with an additional `options` parameter.
 *
 * The `options` parameter is a #lxlsx_row_col_options struct. It has the
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
 *     lxlsx_row_col_options options1 = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     // Hide the fourth and fifth (zero indexed) rows.
 *     lxlsx_worksheet_set_row_opt(worksheet, 3,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 4,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
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
 *     lxlsx_row_col_options options1 = {.hidden = 0, .level = 2, .collapsed = 0};
 *     lxlsx_row_col_options options2 = {.hidden = 0, .level = 1, .collapsed = 0};
 *
 *
 *     // Set the row options with the outline level.
 *     lxlsx_worksheet_set_row_opt(worksheet, 1,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 2,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 3,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 4,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 5,  LXLSX_DEF_ROW_HEIGHT, NULL, &options2);
 *
 *     lxlsx_worksheet_set_row_opt(worksheet, 6,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 7,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 8,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 9,  LXLSX_DEF_ROW_HEIGHT, NULL, &options1);
 *     lxlsx_worksheet_set_row_opt(worksheet, 10, LXLSX_DEF_ROW_HEIGHT, NULL, &options2);
 * @endcode
 *
 * @image html outline1.png
 *
 */
lxlsx_error lxlsx_worksheet_set_row_opt(lxlsx_worksheet *worksheet,
                                lxlsx_row_t row,
                                double height,
                                lxlsx_format *format,
                                lxlsx_row_col_options *options);

/**
 * @brief Set the properties for a row of cells, with the height in pixels.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param pixels    The row height in pixels.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_row_pixels()` function is the same as the
 * `lxlsx_worksheet_set_row()` function except that the height can be set in pixels
 *
 * @code
 *     // Set the height of Row 1 to 20 pixels.
 *     lxlsx_worksheet_set_row_pixels(worksheet, 0, 20, NULL);
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height in pixels: #LXLSX_DEF_ROW_HEIGHT_PIXELS.
 */
lxlsx_error lxlsx_worksheet_set_row_pixels(lxlsx_worksheet *worksheet,
                                   lxlsx_row_t row, uint32_t pixels,
                                   lxlsx_format *format);
/**
 * @brief Set the properties for a row of cells, with the height in pixels.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param pixels    The row height in pixels.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_row_pixels_opt()` function is the same as the
 * `lxlsx_worksheet_set_row_opt()` function except that the height can be set in
 * pixels.
 *
 */
lxlsx_error lxlsx_worksheet_set_row_pixels_opt(lxlsx_worksheet *worksheet,
                                       lxlsx_row_t row,
                                       uint32_t pixels,
                                       lxlsx_format *format,
                                       lxlsx_row_col_options *options);

/**
 * @brief Set the properties for one or more columns of cells.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_column()` function can be used to change the default
 * properties of a single column or a range of columns:
 *
 * @code
 *     // Width of columns B:D set to 30.
 *     lxlsx_worksheet_set_column(worksheet, 1, 3, 30, NULL);
 *
 * @endcode
 *
 * If `%lxlsx_worksheet_set_column()` is applied to a single column the value of
 * `first_col` and `last_col` should be the same:
 *
 * @code
 *     // Width of column B set to 30.
 *     lxlsx_worksheet_set_column(worksheet, 1, 1, 30, NULL);
 *
 * @endcode
 *
 * It is also possible, and generally clearer, to specify a column range using
 * the form of `COLS()` macro:
 *
 * @code
 *     lxlsx_worksheet_set_column(worksheet, 4, 4, 20, NULL);
 *     lxlsx_worksheet_set_column(worksheet, 5, 8, 30, NULL);
 *
 *     // Same as the examples above but clearer.
 *     lxlsx_worksheet_set_column(worksheet, COLS("E:E"), 20, NULL);
 *     lxlsx_worksheet_set_column(worksheet, COLS("F:H"), 30, NULL);
 *
 * @endcode
 *
 * The `width` parameter sets the column width in the same units used by Excel
 * which is: the number of characters in the default font. The default width
 * is 8.43 in the default font of Calibri 11. The actual relationship between
 * a string width and a column width in Excel is complex. See the
 * [following explanation of column widths](https://support.microsoft.com/en-us/kb/214123)
 * from the Microsoft support documentation for more details. To set the width
 * in pixels use the `lxlsx_worksheet_set_column_pixels()` function.
 *
 * There is no way to specify "AutoFit" for a column in the Excel file
 * format. This feature is only available at runtime from within Excel. It is
 * possible to simulate "AutoFit" in your application by tracking the maximum
 * width of the data in the column as your write it and then adjusting the
 * column width at the end.
 *
 * As usual the @ref format.h `format` parameter is optional. If you wish to
 * set the format without changing the width you can pass a default column
 * width of #LXLSX_DEF_COL_WIDTH = 8.43:
 *
 * @code
 *     lxlsx_format *bold = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_bold(bold);
 *
 *     // Set the first column to bold.
 *     lxlsx_worksheet_set_column(worksheet, 0, 0, LXLSX_DEF_COL_WIDTH, bold);
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the column that
 * don't have a format. For example:
 *
 * @code
 *     // Column 1 has format1.
 *     lxlsx_worksheet_set_column(worksheet, COLS("A:A"), 8.43, format1);
 *
 *     // Cell A1 in column 1 defaults to format1.
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell A2 in column 1 keeps format2.
 *     lxlsx_worksheet_write_string(worksheet, 1, 0, "Hello", format2);
 * @endcode
 *
 * As in Excel a row format takes precedence over a default column format:
 *
 * @code
 *     // Row 1 has format1.
 *     lxlsx_worksheet_set_row(worksheet, 0, 15, format1);
 *
 *     // Col 1 has format2.
 *     lxlsx_worksheet_set_column(worksheet, COLS("A:A"), 8.43, format2);
 *
 *     // Cell A1 defaults to format1, the row format.
 *     lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *    // Cell A2 keeps format2, the column format.
 *     lxlsx_worksheet_write_string(worksheet, 1, 0, "Hello", NULL);
 * @endcode
 */
lxlsx_error lxlsx_worksheet_set_column(lxlsx_worksheet *worksheet,
                               lxlsx_col_t first_col,
                               lxlsx_col_t last_col,
                               double width, lxlsx_format *format);

/**
 * @brief Set the properties for one or more columns of cells with options.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_column_opt()` function  is the same as
 * `lxlsx_worksheet_set_column()` with an additional `options` parameter.
 *
 * The `options` parameter is a #lxlsx_row_col_options struct. It has the
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
 *     lxlsx_row_col_options options1 = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     lxlsx_worksheet_set_column_opt(worksheet, COLS("D:E"),  LXLSX_DEF_COL_WIDTH, NULL, &options1);
 * @endcode
 *
 * @image html hide_row_col3.png
 *
 * The `"hidden"`, `"level"`,  and `"collapsed"`, options can also be used to
 * create Outlines and Grouping. See @ref working_with_outlines.
 *
 * @code
 *     lxlsx_row_col_options options1 = {.hidden = 0, .level = 1, .collapsed = 0};
 *
 *     lxlsx_worksheet_set_column_opt(worksheet, COLS("B:G"),  5, NULL, &options1);
 * @endcode
 *
 * @image html outline8.png
 */
lxlsx_error lxlsx_worksheet_set_column_opt(lxlsx_worksheet *worksheet,
                                   lxlsx_col_t first_col,
                                   lxlsx_col_t last_col,
                                   double width,
                                   lxlsx_format *format,
                                   lxlsx_row_col_options *options);

/**
 * @brief Set the properties for one or more columns of cells, with the width
 *        in pixels.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param pixels    The width of the column(s) in pixels.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_column_pixels()` function is the same as
 * `lxlsx_worksheet_set_column()` function except that the width can be set in
 * pixels:
 *
 * @code
 *     // Column width set to 75 pixels, the same as 10 character units.
 *     lxlsx_worksheet_set_column(worksheet, 5, 5, 75, NULL);
 * @endcode
 *
 * @image html set_column_pixels.png
 *
 * If you wish to set the format of a column without changing the width you can
 * pass the default column width in pixels: #LXLSX_DEF_COL_WIDTH_PIXELS.
 */
lxlsx_error lxlsx_worksheet_set_column_pixels(lxlsx_worksheet *worksheet,
                                      lxlsx_col_t first_col,
                                      lxlsx_col_t last_col,
                                      uint32_t pixels, lxlsx_format *format);

/**
 * @brief Set the properties for one or more columns of cells with options,
 *        with the width in pixels.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param pixels    The width of the column(s) in pixels.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_column_pixels_opt()` function is the same as the
 * `lxlsx_worksheet_set_column_opt()` function except that the width can be set in
 * pixels.
 *
 */
lxlsx_error lxlsx_worksheet_set_column_pixels_opt(lxlsx_worksheet *worksheet,
                                          lxlsx_col_t first_col,
                                          lxlsx_col_t last_col,
                                          uint32_t pixels,
                                          lxlsx_format *format,
                                          lxlsx_row_col_options *options);

/**
 * @brief Insert an image in a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 *
 * @return A #lxlsx_error code.
 *
 * This function can be used to insert a image into a worksheet. The image can
 * be in PNG, JPEG, GIF or BMP format:
 *
 * @code
 *     lxlsx_worksheet_insert_image(worksheet, 2, 1, "logo.png");
 * @endcode
 *
 * @image html insert_image.png
 *
 * The `lxlsx_worksheet_insert_image_opt()` function takes additional optional
 * parameters to position and scale the image, see below.
 *
 * **Note**:
 * The scaling of a image may be affected if is crosses a row that has its
 * default height changed due to a font that is larger than the default font
 * size or that has text wrapping turned on. To avoid this you should explicitly
 * set the height of the row using `lxlsx_worksheet_set_row()` if it crosses an
 * inserted image. See @ref working_with_object_positioning.
 *
 * **NOTE on SVG files**:
 * Excel doesn't directly support SVG files in the same way as other image file
 * formats. It allows SVG to be inserted into a worksheet but converts them to,
 * and displays them as, PNG files. It stores the original SVG image in the file
 * so the original format can be retrieved. This removes the file size and
 * resolution advantage of using SVG files. As such SVG files are not supported
 * by `libxlsxwriter` since a conversion to the PNG format would be required
 * and that format is already supported.
 *
 * BMP images are only supported for backward compatibility. In general it is
 * best to avoid BMP images since they aren't compressed. If used, BMP images
 * must be 24 bit, true color, bitmaps.
 */
lxlsx_error lxlsx_worksheet_insert_image(lxlsx_worksheet *worksheet,
                                 lxlsx_row_t row, lxlsx_col_t col,
                                 const char *filename);

/**
 * @brief Insert an image in a worksheet cell, with options.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 * @param options   Optional image parameters.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_insert_image_opt()` function is like
 * `lxlsx_worksheet_insert_image()` function except that it takes an optional
 * #lxlsx_image_options struct with the following members/options:
 *
 * - `x_offset`: Offset from the left of the cell in pixels.
 * - `y_offset`: Offset from the top of the cell in pixels.
 * - `x_scale`: X scale of the image as a decimal.
 * - `y_scale`: Y scale of the image as a decimal.
 * - `object_position`: See @ref working_with_object_positioning.
 * - `description`: Optional description or "Alt text" for the image.
 * - `decorative`: Optional parameter to mark image as decorative.
 * - `url`: Add an optional hyperlink to the image.
 * - `tip`: Add an optional mouseover tip for a hyperlink to the image.
 *
 * For example, to scale and position the image:
 *
 * @code
 *     lxlsx_image_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.5};
 *
 *     lxlsx_worksheet_insert_image_opt(worksheet, 2, 1, "logo.png", &options);
 *
 * @endcode
 *
 * @image html insert_image_opt.png
 *
 * The `url` field of lxlsx_image_options can be use to used to add a hyperlink
 * to an image:
 *
 * @code
 *     lxlsx_image_options options = {.url = "https://github.com/jmcnamara"};
 *
 *     lxlsx_worksheet_insert_image_opt(worksheet, 3, 1, "logo.png", &options);
 * @endcode
 *
 * The supported URL formats are the same as those supported by the
 * `lxlsx_worksheet_write_url()` method and the same rules/limits apply.
 *
 * The `tip` field of lxlsx_image_options can be use to used to add a mouseover
 * tip to the hyperlink:
 *
 * @code
 *      lxlsx_image_options options = {.url = "https://github.com/jmcnamara",
                                     .tip = "GitHub"};
 *
 *     lxlsx_worksheet_insert_image_opt(worksheet, 4, 1, "logo.png", &options);
 * @endcode
 *
 * @note See the notes about row scaling and BMP images in
 * `lxlsx_worksheet_insert_image()` above.
 */
lxlsx_error lxlsx_worksheet_insert_image_opt(lxlsx_worksheet *worksheet,
                                     lxlsx_row_t row, lxlsx_col_t col,
                                     const char *filename,
                                     lxlsx_image_options *options);

/**
 * @brief Insert an image in a worksheet cell, from a memory buffer.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param image_buffer Pointer to an array of bytes that holds the image data.
 * @param image_size   The size of the array of bytes.
 *
 * @return A #lxlsx_error code.
 *
 * This function can be used to insert a image into a worksheet from a memory
 * buffer:
 *
 * @code
 *     lxlsx_worksheet_insert_image_buffer(worksheet, CELL("B3"), image_buffer, image_size);
 * @endcode
 *
 * @image html image_buffer.png
 *
 * The buffer should be a pointer to an array of unsigned char data with a
 * specified size.
 *
 * See `lxlsx_worksheet_insert_image()` for details about the supported image
 * formats, and other image features.
 */
lxlsx_error lxlsx_worksheet_insert_image_buffer(lxlsx_worksheet *worksheet,
                                        lxlsx_row_t row,
                                        lxlsx_col_t col,
                                        const unsigned char *image_buffer,
                                        size_t image_size);

/**
 * @brief Insert an image in a worksheet cell, from a memory buffer.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param image_buffer Pointer to an array of bytes that holds the image data.
 * @param image_size   The size of the array of bytes.
 * @param options      Optional image parameters.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_insert_image_buffer_opt()` function is like
 * `lxlsx_worksheet_insert_image_buffer()` function except that it takes an optional
 * #lxlsx_image_options struct with the following members/options:
 *
 * - `x_offset`: Offset from the left of the cell in pixels.
 * - `y_offset`: Offset from the top of the cell in pixels.
 * - `x_scale`: X scale of the image as a decimal.
 * - `y_scale`: Y scale of the image as a decimal.
 * - `object_position`: See @ref working_with_object_positioning.
 * - `description`: Optional description or "Alt text" for the image.
 * - `decorative`: Optional parameter to mark image as decorative.
 * - `url`: Add an optional hyperlink to the image.
 * - `tip`: Add an optional mouseover tip for a hyperlink to the image.
 *
 * For example, to scale and position the image:
 *
 * @code
 *     lxlsx_image_options options = {.x_offset = 32, .y_offset = 4,
 *                                  .x_scale  = 2,  .y_scale  = 1};
 *
 *     lxlsx_worksheet_insert_image_buffer_opt(worksheet, CELL("B3"), image_buffer, image_size, &options);
 * @endcode
 *
 * @image html image_buffer_opt.png
 *
 * The buffer should be a pointer to an array of unsigned char data with a
 * specified size.
 *
 * See `lxlsx_worksheet_insert_image_buffer_opt()` for details about the supported
 * image formats, and other image options.
 */
lxlsx_error lxlsx_worksheet_insert_image_buffer_opt(lxlsx_worksheet *worksheet,
                                            lxlsx_row_t row,
                                            lxlsx_col_t col,
                                            const unsigned char *image_buffer,
                                            size_t image_size,
                                            lxlsx_image_options *options);

/**
 * @brief Embed an image in a worksheet cell.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 *
 * @return A #lxlsx_error code.
 *
 * This function can be used to embed a image into a worksheet cell and have the
 * image automatically scale to the width and height of the cell. The X/Y
 * scaling of the image is preserved but the size of the image is adjusted to
 * fit the largest possible width or height depending on the cell dimensions.
 *
 * This is the equivalent of Excel's menu option to insert an image using the
 * option to "Place in Cell" which is only available in Excel 365 versions from
 * 2023 onwards. For older versions of Excel a `#VALUE!` error is displayed.
 *
 * @dontinclude embed_images.c
 * @skip Change
 * @until B6
 *
 * @image html embed_image.png
 *
 * The `lxlsx_worksheet_embed_image_opt()` function takes additional optional
 * parameters to add urls or format the cell background, see below.
 *
 */
lxlsx_error lxlsx_worksheet_embed_image(lxlsx_worksheet *worksheet,
                                lxlsx_row_t row, lxlsx_col_t col,
                                const char *filename);

/**
 * @brief Embed an image in a worksheet cell, with options.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param filename  The image filename, with path if required.
 * @param options   Optional image parameters.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_embed_image_opt()` function is like
 * `lxlsx_worksheet_embed_image()` function except that it takes an optional
 * #lxlsx_image_options struct with the following members/options:
 *
 * - `description`: Optional description or "Alt text" for the image.
 * - `decorative`: Optional parameter to mark image as decorative.
 * - `url`: Add an optional hyperlink to the image.
 * - `cell_format`: Add a format for the cell behind the embedded image.
 *
 */
lxlsx_error lxlsx_worksheet_embed_image_opt(lxlsx_worksheet *worksheet,
                                    lxlsx_row_t row, lxlsx_col_t col,
                                    const char *filename,
                                    lxlsx_image_options *options);

/**
 * @brief Embed an image in a worksheet cell, from a memory buffer.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param image_buffer Pointer to an array of bytes that holds the image data.
 * @param image_size   The size of the array of bytes.
 *
 * @return A #lxlsx_error code.
 *
 * This function can be used to embed a image into a worksheet from a memory
 * buffer:
 *
 * @dontinclude embed_image_buffer.c
 * @skip Embed
 * @until B3
 *
 * @image html embed_image_buffer.png
 *
 * The buffer should be a pointer to an array of unsigned char data with a
 * specified size.
 *
 * See `lxlsx_worksheet_embed_image()` for details about the supported image
 * formats, and other image features.
 */
lxlsx_error lxlsx_worksheet_embed_image_buffer(lxlsx_worksheet *worksheet,
                                       lxlsx_row_t row,
                                       lxlsx_col_t col,
                                       const unsigned char *image_buffer,
                                       size_t image_size);

/**
 * @brief Embed an image in a worksheet cell, from a memory buffer.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param image_buffer Pointer to an array of bytes that holds the image data.
 * @param image_size   The size of the array of bytes.
 * @param options      Optional image parameters.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_embed_image_buffer_opt()` function is like
 * `lxlsx_worksheet_embed_image_buffer()` function except that it takes an optional
 * #lxlsx_image_options struct with the following members/options:
 *
 * - `description`: Optional description or "Alt text" for the image.
 * - `decorative`: Optional parameter to mark image as decorative.
 * - `url`: Add an optional hyperlink to the image.
 * - `cell_format`: Add a format for the cell behind the embedded image.
 *
 */
lxlsx_error lxlsx_worksheet_embed_image_buffer_opt(lxlsx_worksheet *worksheet,
                                           lxlsx_row_t row,
                                           lxlsx_col_t col,
                                           const unsigned char *image_buffer,
                                           size_t image_size,
                                           lxlsx_image_options *options);

/**
 * @brief Set the background image for a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param filename  The image filename, with path if required.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_background()` function can be used to set the
 * background image for a worksheet:
 *
 * @code
 *      lxlsx_worksheet_set_background(worksheet, "logo.png");
 * @endcode
 *
 * @image html background.png
 *
 * The ``set_background()`` method supports all the image formats supported by
 * `lxlsx_worksheet_insert_image()`.
 *
 * Some people use this method to add a watermark background to their
 * document. However, Microsoft recommends using a header image [to set a
 * watermark][watermark]. The choice of method depends on whether you want the
 * watermark to be visible in normal viewing mode or just when the file is
 * printed. In libxlsxwriter you can get the header watermark effect using
 * `lxlsx_worksheet_set_header()`:
 *
 * @code
 *     lxlsx_header_footer_options header_options = {.image_center = "watermark.png"};
 *     lxlsx_worksheet_set_header_opt(worksheet, "&C&G", &header_options);
 * @endcode
 *
 * [watermark]:https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
 *
 */
lxlsx_error lxlsx_worksheet_set_background(lxlsx_worksheet *worksheet,
                                   const char *filename);

/**
 * @brief Set the background image for a worksheet, from a buffer.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param image_buffer Pointer to an array of bytes that holds the image data.
 * @param image_size   The size of the array of bytes.
 *
 * @return A #lxlsx_error code.
 *
 * This function can be used to insert a background image into a worksheet
 * from a memory buffer:
 *
 * @code
 *     lxlsx_worksheet_set_background_buffer(worksheet, image_buffer, image_size);
 * @endcode
 *
 * The buffer should be a pointer to an array of unsigned char data with a
 * specified size.
 *
 * See `lxlsx_worksheet_set_background()` for more details.
 *
 */
lxlsx_error lxlsx_worksheet_set_background_buffer(lxlsx_worksheet *worksheet,
                                          const unsigned char *image_buffer,
                                          size_t image_size);

/**
 * @brief Insert a chart object into a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param chart     A #lxlsx_chart object created via lxlsx_workbook_add_chart().
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_insert_chart()` function can be used to insert a chart into
 * a worksheet. The chart object must be created first using the
 * `lxlsx_workbook_add_chart()` function and configured using the @ref chart.h
 * functions.
 *
 * @code
 *     // Create a chart object.
 *     lxlsx_chart *chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_LINE);
 *
 *     // Add a data series to the chart.
 *     lxlsx_chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$6");
 *
 *     // Insert the chart into the worksheet.
 *     lxlsx_worksheet_insert_chart(worksheet, 0, 2, chart);
 * @endcode
 *
 * @image html lxlsx_chart_working.png
 *
 * **Note:**
 *
 * A chart may only be inserted into a worksheet once. If several similar
 * charts are required then each one must be created separately with
 * `%lxlsx_worksheet_insert_chart()`.
 *
 */
lxlsx_error lxlsx_worksheet_insert_chart(lxlsx_worksheet *worksheet,
                                 lxlsx_row_t row, lxlsx_col_t col,
                                 lxlsx_chart *chart);

/**
 * @brief Insert a chart object into a worksheet, with options.
 *
 * @param worksheet    Pointer to a lxlsx_worksheet instance to be updated.
 * @param row          The zero indexed row number.
 * @param col          The zero indexed column number.
 * @param chart        A #lxlsx_chart object created via lxlsx_workbook_add_chart().
 * @param user_options Optional chart parameters.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_insert_chart_opt()` function is like
 * `lxlsx_worksheet_insert_chart()` function except that it takes an optional
 * #lxlsx_chart_options struct to scale and position the chart:
 *
 * @code
 *    lxlsx_chart_options options = {.x_offset = 30,  .y_offset = 10,
 *                                 .x_scale  = 0.5, .y_scale  = 0.75};
 *
 *    lxlsx_worksheet_insert_chart_opt(worksheet, 0, 2, chart, &options);
 *
 * @endcode
 *
 * @image html lxlsx_chart_line_opt.png
 *
 */
lxlsx_error lxlsx_worksheet_insert_chart_opt(lxlsx_worksheet *worksheet,
                                     lxlsx_row_t row, lxlsx_col_t col,
                                     lxlsx_chart *chart,
                                     lxlsx_chart_options *user_options);

/**
 * @brief Merge a range of cells.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param string    String to write to the merged range.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_merge_range()` function allows cells to be merged together
 * so that they act as a single area.
 *
 * Excel generally merges and centers cells at same time. To get similar
 * behavior with libxlsxwriter you need to apply a @ref format.h "Format"
 * object with the appropriate alignment:
 *
 * @code
 *     lxlsx_format *merge_format = lxlsx_workbook_add_format(workbook);
 *     lxlsx_format_set_align(merge_format, LXLSX_ALIGN_CENTER);
 *
 *     lxlsx_worksheet_merge_range(worksheet, 1, 1, 1, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * It is possible to apply other formatting to the merged cells as well:
 *
 * @code
 *    lxlsx_format_set_align   (merge_format, LXLSX_ALIGN_CENTER);
 *    lxlsx_format_set_align   (merge_format, LXLSX_ALIGN_VERTICAL_CENTER);
 *    lxlsx_format_set_border  (merge_format, LXLSX_BORDER_DOUBLE);
 *    lxlsx_format_set_bold    (merge_format);
 *    lxlsx_format_set_bg_color(merge_format, 0xD7E4BC);
 *
 *    lxlsx_worksheet_merge_range(worksheet, 2, 1, 3, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * @image html merge.png
 *
 * The `%lxlsx_worksheet_merge_range()` function writes a `char*` string using
 * `lxlsx_worksheet_write_string()`. In order to write other data types, such as a
 * number or a formula, you can overwrite the first cell with a call to one of
 * the other write functions. The same Format should be used as was used in
 * the merged range.
 *
 * @code
 *    // First write a range with a blank string.
 *    lxlsx_worksheet_merge_range (worksheet, 1, 1, 1, 3, "", format);
 *
 *    // Then overwrite the first cell with a number.
 *    lxlsx_worksheet_write_number(worksheet, 1, 1, 123, format);
 * @endcode
 *
 * @note Merged ranges generally don't work in libxlsxwriter when the Workbook
 * #lxlsx_workbook_options `constant_memory` mode is enabled.
 */
lxlsx_error lxlsx_worksheet_merge_range(lxlsx_worksheet *worksheet, lxlsx_row_t first_row,
                                lxlsx_col_t first_col, lxlsx_row_t last_row,
                                lxlsx_col_t last_col, const char *string,
                                lxlsx_format *format);

/**
 * @brief Set the autofilter area in the worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_autofilter()` function allows an autofilter to be added to
 * a worksheet.
 *
 * An autofilter is a way of adding dropdown lists to the headers of a 2D
 * range of worksheet data. This allows users to filter the data based on
 * simple criteria so that some data is shown and some is hidden.
 *
 * @image html autofilter3.png
 *
 * To add an autofilter to a worksheet:
 *
 * @code
 *     lxlsx_worksheet_autofilter(worksheet, 0, 0, 50, 3);
 *
 *     // Same as above using the RANGE() macro.
 *     lxlsx_worksheet_autofilter(worksheet, RANGE("A1:D51"));
 * @endcode
 *
 * In order to apply a filter condition it is necessary to add filter rules to
 * the columns using either the `%lxlsx_worksheet_filter_column()`,
 * `%lxlsx_worksheet_filter_column2()` or `%lxlsx_worksheet_filter_list()` functions:
 *
 * - `lxlsx_worksheet_filter_column()`: filter on a single criterion such as "Column ==
 *   East". More complex conditions such as "<=" or ">=" can also be use.
 *
 * - `lxlsx_worksheet_filter_column2()`: filter on two criteria such as "Column == East
 *   or Column == West". Complex conditions can also be used.
 *
 * - `lxlsx_worksheet_filter_list()`: filter on a list of values such as "Column in (East, West,
 *   North)".
 *
 * These functions are explained below. It isn't sufficient to just specify
 * the filter condition. You must also hide any rows that don't match the
 * filter condition. See @ref ww_autofilters_data for more details.
 *
 */
lxlsx_error lxlsx_worksheet_autofilter(lxlsx_worksheet *worksheet, lxlsx_row_t first_row,
                               lxlsx_col_t first_col, lxlsx_row_t last_row,
                               lxlsx_col_t last_col);

/**
 * @brief Write a filter rule to an autofilter column.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param col       The column in the autofilter that the rule applies to.
 * @param rule      The lxlsx_filter_rule autofilter rule.
 *
 * @return A #lxlsx_error code.
 *
 * The `lxlsx_worksheet_filter_column` function can be used to filter columns in a
 * autofilter range based on single rule conditions:
 *
 * @code
 *     lxlsx_filter_rule filter_rule = {.criteria     = LXLSX_FILTER_CRITERIA_EQUAL_TO,
 *                                    .value_string = "East"};
 *
 *    lxlsx_worksheet_filter_column(worksheet, 0, &filter_rule);
 *@endcode
 *
 * @image html autofilter4.png
 *
 * The rules and criteria are explained in more detail in @ref
 * ww_autofilters_criteria in @ref working_with_autofilters.
 *
 * The `col` parameter is a zero indexed column number and must refer to a
 * column in an existing autofilter created with `lxlsx_worksheet_autofilter()`.
 *
 * It isn't sufficient to just specify the filter condition. You must also
 * hide any rows that don't match the filter condition. See @ref
 * ww_autofilters_data for more details.
 */
lxlsx_error lxlsx_worksheet_filter_column(lxlsx_worksheet *worksheet, lxlsx_col_t col,
                                  lxlsx_filter_rule *rule);

/**
 * @brief Write two filter rules to an autofilter column.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param col       The column in the autofilter that the rules applies to.
 * @param rule1     First lxlsx_filter_rule autofilter rule.
 * @param rule2     Second lxlsx_filter_rule autofilter rule.
 * @param and_or    A #lxlsx_filter_operator and/or operator.
 *
 * @return A #lxlsx_error code.
 *
 * The `lxlsx_worksheet_filter_column2` function can be used to filter columns in a autofilter
 * range based on two rule conditions:
 *
 * @code
 *     lxlsx_filter_rule filter_rule1 = {.criteria     = LXLSX_FILTER_CRITERIA_EQUAL_TO,
 *                                     .value_string = "East"};
 *
 *     lxlsx_filter_rule filter_rule2 = {.criteria     = LXLSX_FILTER_CRITERIA_EQUAL_TO,
 *                                     .value_string = "South"};
 *
 *     lxlsx_worksheet_filter_column2(worksheet, 0, &filter_rule1, &filter_rule2, LXLSX_FILTER_OR);
 * @endcode
 *
 * @image html autofilter5.png
 *
 * The rules and criteria are explained in more detail in @ref
 * ww_autofilters_criteria in @ref working_with_autofilters.
 *
 * The `col` parameter is a zero indexed column number and must refer to a
 * column in an existing autofilter created with `lxlsx_worksheet_autofilter()`.
 *
 * The `and_or` parameter is either "and (LXLSX_FILTER_AND)" or "or  (LXLSX_FILTER_OR)".
 *
 * It isn't sufficient to just specify the filter condition. You must also
 * hide any rows that don't match the filter condition. See @ref
 * ww_autofilters_data for more details.
 */
lxlsx_error lxlsx_worksheet_filter_column2(lxlsx_worksheet *worksheet, lxlsx_col_t col,
                                   lxlsx_filter_rule *rule1,
                                   lxlsx_filter_rule *rule2, uint8_t and_or);
/**
 * @brief Write multiple string filters to an autofilter column.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param col       The column in the autofilter that the rules applies to.
 * @param list      A NULL terminated array of strings to filter on.
 *
 * @return A #lxlsx_error code.
 *
 * The `lxlsx_worksheet_filter_column_list()` function can be used specify multiple
 * string matching criteria. This is a newer type of filter introduced in
 * Excel 2007. Prior to that it was only possible to have either 1 or 2 filter
 * conditions, such as the ones used by `lxlsx_worksheet_filter_column()` and
 * `lxlsx_worksheet_filter_column2()`.
 *
 * As an example, consider a column that contains data for the months of the
 * year. The `%lxlsx_worksheet_filter_list()` function can be used to filter out
 * data rows for different months:
 *
 * @code
 *     char* list[] = {"March", "April", "May", NULL};
 *
 *     lxlsx_worksheet_filter_list(worksheet, 0, list);
 * @endcode
 *
 * @image html autofilter2.png
 *
 *
 * Note, the array must be NULL terminated to indicate the end of the array of
 * strings. To filter blanks as part of the list use `Blanks` as a list item:
 *
 * @code
 *     char* list[] = {"March", "April", "May", "Blanks", NULL};
 *
 *     lxlsx_worksheet_filter_list(worksheet, 0, list);
 * @endcode
 *
 * It isn't sufficient to just specify the filter condition. You must also
 * hide any rows that don't match the filter condition. See @ref
 * ww_autofilters_data for more details.
 */
lxlsx_error lxlsx_worksheet_filter_list(lxlsx_worksheet *worksheet, lxlsx_col_t col,
                                const char **list);

/**
 * @brief Add a data validation to a cell.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param row        The zero indexed row number.
 * @param col        The zero indexed column number.
 * @param validation A #lxlsx_data_validation object to control the validation.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_data_validation_cell()` function is used to construct an
 * Excel data validation or to limit the user input to a dropdown list of
 * values:
 *
 * @code
 *
 *    lxlsx_data_validation *data_validation = calloc(1, sizeof(lxlsx_data_validation));
 *
 *    data_validation->validate       = LXLSX_VALIDATION_TYPE_INTEGER;
 *    data_validation->criteria       = LXLSX_VALIDATION_CRITERIA_BETWEEN;
 *    data_validation->minimum_number = 1;
 *    data_validation->maximum_number = 10;
 *
 *    lxlsx_worksheet_data_validation_cell(worksheet, 2, 1, data_validation);
 *
 *    // Same as above with the CELL() macro.
 *    lxlsx_worksheet_data_validation_cell(worksheet, CELL("B3"), data_validation);
 *
 * @endcode
 *
 * @image html data_validate4.png
 *
 * Data validation and the various options of #lxlsx_data_validation are
 * described in more detail in @ref working_with_data_validation.
 */
lxlsx_error lxlsx_worksheet_data_validation_cell(lxlsx_worksheet *worksheet,
                                         lxlsx_row_t row, lxlsx_col_t col,
                                         lxlsx_data_validation *validation);

/**
 * @brief Add a data validation to a range.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row  The first row of the range. (All zero indexed.)
 * @param first_col  The first column of the range.
 * @param last_row   The last row of the range.
 * @param last_col   The last col of the range.
 * @param validation A #lxlsx_data_validation object to control the validation.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_data_validation_range()` function is the same as the
 * `%lxlsx_worksheet_data_validation_cell()`, see above,  except the data validation
 * is applied to a range of cells:
 *
 * @code
 *
 *    lxlsx_data_validation *data_validation = calloc(1, sizeof(lxlsx_data_validation));
 *
 *    data_validation->validate       = LXLSX_VALIDATION_TYPE_INTEGER;
 *    data_validation->criteria       = LXLSX_VALIDATION_CRITERIA_BETWEEN;
 *    data_validation->minimum_number = 1;
 *    data_validation->maximum_number = 10;
 *
 *    lxlsx_worksheet_data_validation_range(worksheet, 2, 1, 4, 1, data_validation);
 *
 *    // Same as above with the RANGE() macro.
 *    lxlsx_worksheet_data_validation_range(worksheet, RANGE("B3:B5"), data_validation);
 *
 * @endcode
 *
 * Data validation and the various options of #lxlsx_data_validation are
 * described in more detail in @ref working_with_data_validation.
 */
lxlsx_error lxlsx_worksheet_data_validation_range(lxlsx_worksheet *worksheet,
                                          lxlsx_row_t first_row,
                                          lxlsx_col_t first_col,
                                          lxlsx_row_t last_row,
                                          lxlsx_col_t last_col,
                                          lxlsx_data_validation *validation);

/**
 * @brief Add a conditional format to a worksheet cell.
 *
 * @param worksheet           Pointer to a lxlsx_worksheet instance to be updated.
 * @param row                 The zero indexed row number.
 * @param col                 The zero indexed column number.
 * @param conditional_format  A #lxlsx_conditional_format object to control the
 *                            conditional format.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_conditional_format_cell()` function is used to set a
 * conditional format for a cell in a worksheet:
 *
 * @code
 *     conditional_format->type     = LXLSX_CONDITIONAL_TYPE_CELL;
 *     conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
 *     conditional_format->value    = 50;
 *     conditional_format->format   = format1;
 *     lxlsx_worksheet_conditional_format_cell(worksheet, CELL("A1"), conditional_format);
 * @endcode
 *
 * The conditional format parameters is specified in #lxlsx_conditional_format.
 *
 * See @ref working_with_conditional_formatting for full details.
 */
lxlsx_error lxlsx_worksheet_conditional_format_cell(lxlsx_worksheet *worksheet,
                                            lxlsx_row_t row,
                                            lxlsx_col_t col,
                                            lxlsx_conditional_format
                                            *conditional_format);

/**
 * @brief Add a conditional format to a worksheet range.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row  The first row of the range. (All zero indexed.)
 * @param first_col  The first column of the range.
 * @param last_row   The last row of the range.
 * @param last_col   The last col of the range.
 * @param conditional_format  A #lxlsx_conditional_format object to control the
 *                            conditional format.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_conditional_format_cell()` function is used to set a
 * conditional format for a range of cells in a worksheet:
 *
 * @code
 *     conditional_format->type     = LXLSX_CONDITIONAL_TYPE_CELL;
 *     conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
 *     conditional_format->value    = 50;
 *     conditional_format->format   = format1;
 *     lxlsx_worksheet_conditional_format_range(worksheet1, RANGE("B3:K12"), conditional_format);
 *
 *     conditional_format->type     = LXLSX_CONDITIONAL_TYPE_CELL;
 *     conditional_format->criteria = LXLSX_CONDITIONAL_CRITERIA_LESS_THAN;
 *     conditional_format->value    = 50;
 *     conditional_format->format   = format2;
 *     lxlsx_worksheet_conditional_format_range(worksheet1, RANGE("B3:K12"), conditional_format);
 * @endcode
 *
 * Output:
 *
 * @image html conditional_format1.png
 *
 *
 * The conditional format parameters is specified in #lxlsx_conditional_format.
 *
 * See @ref working_with_conditional_formatting for full details.
 */
lxlsx_error lxlsx_worksheet_conditional_format_range(lxlsx_worksheet *worksheet,
                                             lxlsx_row_t first_row,
                                             lxlsx_col_t first_col,
                                             lxlsx_row_t last_row,
                                             lxlsx_col_t last_col,
                                             lxlsx_conditional_format
                                             *conditional_format);
/**
 * @brief Insert a button object into a worksheet.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param row        The zero indexed row number.
 * @param col        The zero indexed column number.
 * @param options    A #lxlsx_button_options object to set the button properties.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_insert_button()` function can be used to insert an Excel
 * form button into a worksheet. This function is generally only useful when
 * used in conjunction with the `lxlsx_workbook_add_vba_project()` function to tie
 * the button to a macro from an embedded VBA project:
 *
 * @code
 *     lxlsx_button_options options = {.caption = "Press Me",
 *                                   .macro   = "say_hello"};
 *
 *     lxlsx_worksheet_insert_button(worksheet, 2, 1, &options);
 * @endcode
 *
 * @image html macros.png
 *
 * The button properties are set using the lxlsx_button_options struct.
 *
 * See also @ref working_with_macros
 */
lxlsx_error lxlsx_worksheet_insert_button(lxlsx_worksheet *worksheet, lxlsx_row_t row,
                                  lxlsx_col_t col, lxlsx_button_options *options);

/**
 * @brief Add an Excel table to a worksheet.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row  The first row of the range. (All zero indexed.)
 * @param first_col  The first column of the range.
 * @param last_row   The last row of the range.
 * @param last_col   The last col of the range.
 * @param options    A #lxlsx_table_options struct to define the table options.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_add_table()` function is used to add a table to a
 * worksheet. Tables in Excel are a way of grouping a range of cells into a
 * single entity that has common formatting or that can be referenced from
 * formulas. Tables can have column headers, autofilters, total rows, column
 * formulas and default formatting.
 *
 * @code
 *     lxlsx_worksheet_add_table(worksheet, 2, 1, 6, 5, NULL);
 * @endcode
 *
 * Output:
 *
 * @image html tables1.png
 *
 * See @ref working_with_tables for more detailed usage information and also
 * @ref tables.c.
 *
 */
lxlsx_error lxlsx_worksheet_add_table(lxlsx_worksheet *worksheet, lxlsx_row_t first_row,
                              lxlsx_col_t first_col, lxlsx_row_t last_row,
                              lxlsx_col_t last_col, lxlsx_table_options *options);

 /**
  * @brief Make a worksheet the active, i.e., visible worksheet.
  *
  * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
  *
  * The `%lxlsx_worksheet_activate()` function is used to specify which worksheet is
  * initially visible in a multi-sheet workbook:
  *
  * @code
  *     lxlsx_worksheet *worksheet1 = lxlsx_workbook_add_worksheet(workbook, NULL);
  *     lxlsx_worksheet *worksheet2 = lxlsx_workbook_add_worksheet(workbook, NULL);
  *     lxlsx_worksheet *worksheet3 = lxlsx_workbook_add_worksheet(workbook, NULL);
  *
  *     lxlsx_worksheet_activate(worksheet3);
  * @endcode
  *
  * @image html lxlsx_worksheet_activate.png
  *
  * More than one worksheet can be selected via the `lxlsx_worksheet_select()`
  * function, see below, however only one worksheet can be active.
  *
  * The default active worksheet is the first worksheet.
  *
  */
void lxlsx_worksheet_activate(lxlsx_worksheet *worksheet);

 /**
  * @brief Set a worksheet tab as selected.
  *
  * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
  *
  * The `%lxlsx_worksheet_select()` function is used to indicate that a worksheet is
  * selected in a multi-sheet workbook:
  *
  * @code
  *     lxlsx_worksheet_activate(worksheet1);
  *     lxlsx_worksheet_select(worksheet2);
  *     lxlsx_worksheet_select(worksheet3);
  *
  * @endcode
  *
  * A selected worksheet has its tab highlighted. Selecting worksheets is a
  * way of grouping them together so that, for example, several worksheets
  * could be printed in one go. A worksheet that has been activated via the
  * `lxlsx_worksheet_activate()` function will also appear as selected.
  *
  */
void lxlsx_worksheet_select(lxlsx_worksheet *worksheet);

/**
 * @brief Hide the current worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * The `%lxlsx_worksheet_hide()` function is used to hide a worksheet:
 *
 * @code
 *     lxlsx_worksheet_hide(worksheet2);
 * @endcode
 *
 * You may wish to hide a worksheet in order to avoid confusing a user with
 * intermediate data or calculations.
 *
 * @image html hide_sheet.png
 *
 * A hidden worksheet can not be activated or selected so this function is
 * mutually exclusive with the `lxlsx_worksheet_activate()` and `lxlsx_worksheet_select()`
 * functions. In addition, since the first worksheet will default to being the
 * active worksheet, you cannot hide the first worksheet without activating
 * another sheet:
 *
 * @code
 *     lxlsx_worksheet_activate(worksheet2);
 *     lxlsx_worksheet_hide(worksheet1);
 * @endcode
 */
void lxlsx_worksheet_hide(lxlsx_worksheet *worksheet);

/**
 * @brief Set current worksheet as the first visible sheet tab.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * The `lxlsx_worksheet_activate()` function determines which worksheet is initially
 * selected.  However, if there are a large number of worksheets the selected
 * worksheet may not appear on the screen. To avoid this you can select the
 * leftmost visible worksheet tab using `%lxlsx_worksheet_set_first_sheet()`:
 *
 * @code
 *     lxlsx_worksheet_set_first_sheet(worksheet19); // First visible worksheet tab.
 *     lxlsx_worksheet_activate(worksheet20);        // First visible worksheet.
 * @endcode
 *
 * This function is not required very often. The default value is the first
 * worksheet.
 */
void lxlsx_worksheet_set_first_sheet(lxlsx_worksheet *worksheet);

/**
 * @brief Split and freeze a worksheet into panes.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The cell row (zero indexed).
 * @param col       The cell column (zero indexed).
 *
 * The `%lxlsx_worksheet_freeze_panes()` function can be used to divide a worksheet
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
 *     lxlsx_worksheet_freeze_panes(worksheet1, 1, 0); // Freeze the first row.
 *     lxlsx_worksheet_freeze_panes(worksheet2, 0, 1); // Freeze the first column.
 *     lxlsx_worksheet_freeze_panes(worksheet3, 1, 1); // Freeze first row/column.
 *
 * @endcode
 *
 */
lxlsx_error lxlsx_worksheet_freeze_panes(lxlsx_worksheet *worksheet,
                            lxlsx_row_t row, lxlsx_col_t col);
/**
 * @brief Split a worksheet into panes.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param vertical   The position for the vertical split.
 * @param horizontal The position for the horizontal split.
 *
 * The `%lxlsx_worksheet_split_panes()` function can be used to divide a worksheet
 * into horizontal or vertical regions known as panes. This function is
 * different from the `lxlsx_worksheet_freeze_panes()` function in that the splits
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
 *     lxlsx_worksheet_split_panes(worksheet1, 15, 0);    // First row.
 *     lxlsx_worksheet_split_panes(worksheet2, 0,  8.43); // First column.
 *     lxlsx_worksheet_split_panes(worksheet3, 15, 8.43); // First row and column.
 *
 * @endcode
 *
 */
void lxlsx_worksheet_split_panes(lxlsx_worksheet *worksheet,
                           double vertical, double horizontal);

/* lxlsx_worksheet_freeze_panes() with infrequent options. Undocumented for now. */
void lxlsx_worksheet_freeze_panes_opt(lxlsx_worksheet *worksheet,
                                lxlsx_row_t first_row, lxlsx_col_t first_col,
                                lxlsx_row_t top_row, lxlsx_col_t left_col,
                                uint8_t type);

/* lxlsx_worksheet_split_panes() with infrequent options. Undocumented for now. */
void lxlsx_worksheet_split_panes_opt(lxlsx_worksheet *worksheet,
                               double vertical, double horizontal,
                               lxlsx_row_t top_row, lxlsx_col_t left_col);
/**
 * @brief Set the selected cell or cells in a worksheet:
 *
 * @param worksheet   A pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 *
 *
 * The `%lxlsx_worksheet_set_selection()` function can be used to specify which cell
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
 *     lxlsx_worksheet_set_selection(worksheet1, 3, 3, 3, 3);     // Cell D4.
 *     lxlsx_worksheet_set_selection(worksheet2, 3, 3, 6, 6);     // Cells D4 to G7.
 *     lxlsx_worksheet_set_selection(worksheet3, 6, 6, 3, 3);     // Cells G7 to D4.
 *     lxlsx_worksheet_set_selection(worksheet5, RANGE("D4:G7")); // Using the RANGE macro.
 *
 * @endcode
 *
 */
lxlsx_error lxlsx_worksheet_set_selection(lxlsx_worksheet *worksheet,
                                  lxlsx_row_t first_row, lxlsx_col_t first_col,
                                  lxlsx_row_t last_row, lxlsx_col_t last_col);

/**
 * @brief Set the first visible cell at the top left of a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param row       The cell row (zero indexed).
 * @param col       The cell column (zero indexed).
 *
 * The `%lxlsx_worksheet_set_top_left_cell()` function can be used to set the
 * top leftmost visible cell in the worksheet:
 *
 * @code
 *     lxlsx_worksheet_set_top_left_cell(worksheet, 31, 26);
 *     lxlsx_worksheet_set_top_left_cell(worksheet, CELL("AA32")); // Same as above.
 * @endcode
 *
 * @image html top_left_cell.png
 *
 */
void lxlsx_worksheet_set_top_left_cell(lxlsx_worksheet *worksheet, lxlsx_row_t row,
                                 lxlsx_col_t col);

/**
 * @brief Set the page orientation as landscape.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to landscape:
 *
 * @code
 *     lxlsx_worksheet_set_landscape(worksheet);
 * @endcode
 */
void lxlsx_worksheet_set_landscape(lxlsx_worksheet *worksheet);

/**
 * @brief Set the page orientation as portrait.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to portrait. The default worksheet orientation is portrait, so this
 * function isn't generally required:
 *
 * @code
 *     lxlsx_worksheet_set_portrait(worksheet);
 * @endcode
 */
void lxlsx_worksheet_set_portrait(lxlsx_worksheet *worksheet);

/**
 * @brief Set the page layout to page view mode.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * This function is used to display the worksheet in "Page View/Layout" mode:
 *
 * @code
 *     lxlsx_worksheet_set_page_view(worksheet);
 * @endcode
 */
void lxlsx_worksheet_set_page_view(lxlsx_worksheet *worksheet);

/**
 * @brief Set the paper type for printing.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
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
 *     lxlsx_worksheet_set_paper(worksheet1, 1);  // US Letter
 *     lxlsx_worksheet_set_paper(worksheet2, 9);  // A4
 * @endcode
 *
 * If you do not specify a paper type the worksheet will print using the
 * printer's default paper style.
 */
void lxlsx_worksheet_set_paper(lxlsx_worksheet *worksheet, uint8_t paper_type);

/**
 * @brief Set the worksheet margins for the printed page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param left    Left margin in inches.   Excel default is 0.7.
 * @param right   Right margin in inches.  Excel default is 0.7.
 * @param top     Top margin in inches.    Excel default is 0.75.
 * @param bottom  Bottom margin in inches. Excel default is 0.75.
 *
 * The `%lxlsx_worksheet_set_margins()` function is used to set the margins of the
 * worksheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    lxlsx_worksheet_set_margins(worksheet, 1.3, 1.2, -1, -1);
 * @endcode
 *
 */
void lxlsx_worksheet_set_margins(lxlsx_worksheet *worksheet, double left,
                           double right, double top, double bottom);

/**
 * @brief Set the printed page header caption.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param string    The header string.
 *
 * @return A #lxlsx_error code.
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
 *   | `&[Picture]`    | Images        | Image placeholder     |
 *   | `&G`            |               | Same as `&[Picture]`  |
 *   | `&&`            | Miscellaneous | Literal ampersand &   |
 *
 * Note: inserting images requires the `lxlsx_worksheet_set_header_opt()` function.
 *
 * Text in headers and footers can be justified (aligned) to the left, center
 * and right by prefixing the text with the control characters `&L`, `&C` and
 * `&R`.
 *
 * For example (with ASCII art representation of the results):
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet, "&LHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    | Hello                                                         |
 *     //    |                                                               |
 *
 *
 *     lxlsx_worksheet_set_header(worksheet, "&CHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                          Hello                                |
 *     //    |                                                               |
 *
 *
 *     lxlsx_worksheet_set_header(worksheet, "&RHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                                                         Hello |
 *     //    |                                                               |
 *
 *
 * @endcode
 *
 * For simple text, if you do not specify any justification the text will be
 * centered. However, you must prefix the text with `&C` if you specify a font
 * name or any other formatting:
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet, "Hello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                          Hello                                |
 *     //    |                                                               |
 *
 * @endcode
 *
 * You can have text in each of the justification regions:
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet, "&LCiao&CBello&RCielo");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    | Ciao                     Bello                          Cielo |
 *     //    |                                                               |
 *
 * @endcode
 *
 * The information control characters act as variables that Excel will update
 * as the workbook or worksheet changes. Times and dates are in the users
 * default format:
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet, "&CPage &P of &N");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                        Page 1 of 6                            |
 *     //    |                                                               |
 *
 *     lxlsx_worksheet_set_header(worksheet, "&CUpdated at &T");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                    Updated at 12:30 PM                        |
 *     //    |                                                               |
 *
 * @endcode
 *
 * You can specify the font size of a section of the text by prefixing it with
 * the control character `&n` where `n` is the font size:
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet1, "&C&30Hello Big");
 *     lxlsx_worksheet_set_header(worksheet2, "&C&10Hello Small");
 *
 * @endcode
 *
 * You can specify the font of a section of the text by prefixing it with the
 * control sequence `&"font,style"` where `fontname` is a font name such as
 * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
 * "Courier New" or "Times New Roman" and `style` is one of the standard
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet1, "&C&\"Courier New,Italic\"Hello");
 *     lxlsx_worksheet_set_header(worksheet2, "&C&\"Courier New,Bold Italic\"Hello");
 *     lxlsx_worksheet_set_header(worksheet3, "&C&\"Times New Roman,Regular\"Hello");
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
 *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer" | sed 's/&amp;/\&/g'
 *
 *      <headerFooter scaleWithDoc="0">
 *        <oddHeader>&L&P</oddHeader>
 *      </headerFooter>
 *
 * @endcode
 *
 * To include a single literal ampersand `&` in a header or footer you should
 * use a double ampersand `&&`:
 *
 * @code
 *     lxlsx_worksheet_set_header(worksheet, "&CCuriouser && Curiouser - Attorneys at Law");
 * @endcode
 *
 * @note
 * Excel requires that the header or footer string cannot be longer than 255
 * characters, including the control characters. Strings longer than this will
 * not be written.
 *
 */
lxlsx_error lxlsx_worksheet_set_header(lxlsx_worksheet *worksheet, const char *string);

/**
 * @brief Set the printed page footer caption.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param string    The footer string.
 *
 * @return A #lxlsx_error code.
 *
 * The syntax of this function is the same as lxlsx_worksheet_set_header().
 *
 */
lxlsx_error lxlsx_worksheet_set_footer(lxlsx_worksheet *worksheet, const char *string);

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param string    The header string.
 * @param options   Header options.
 *
 * @return A #lxlsx_error code.
 *
 * The syntax of this function is the same as `lxlsx_worksheet_set_header()` with an
 * additional parameter to specify options for the header.
 *
 * The #lxlsx_header_footer_options options are:
 *
 * - `margin`: Header or footer margin in inches. The value must by larger
 *   than 0.0. The Excel default is 0.3.
 *
 * - `image_left`: The left header image filename, with path if required. This
 *   should have a corresponding `&G/&[Picture]` placeholder in the `&L`
 *   section of the header/footer string.
 *
 * - `image_center`: The center header image filename, with path if
 *   required. This should have a corresponding `&G/&[Picture]` placeholder in
 *   the `&C` section of the header/footer string.
 *
 * - `image_right`: The right header image filename, with path if
 *   required. This should have a corresponding `&G/&[Picture]` placeholder in
 *   the `&R` section of the header/footer string.
 *
 * @code
 *     lxlsx_header_footer_options header_options = { .margin = 0.2 };
 *
 *     lxlsx_worksheet_set_header_opt(worksheet, "Some text", &header_options);
 * @endcode
 *
 * Images can be inserted in the header by specifying the `&[Picture]`
 * placeholder and a filename/path to the image:
 *
 * @code
 *     lxlsx_header_footer_options header_options = {.image_left = "logo.png"};
 *
 *    lxlsx_worksheet_set_header_opt(worksheet, "&L&[Picture]", &header_options);
 * @endcode
 *
 * @image html headers_footers.png
 *
 */
lxlsx_error lxlsx_worksheet_set_header_opt(lxlsx_worksheet *worksheet,
                                   const char *string,
                                   lxlsx_header_footer_options *options);

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param string    The footer string.
 * @param options   Footer options.
 *
 * @return A #lxlsx_error code.
 *
 * The syntax of this function is the same as `lxlsx_worksheet_set_header_opt()`.
 *
 */
lxlsx_error lxlsx_worksheet_set_footer_opt(lxlsx_worksheet *worksheet,
                                   const char *string,
                                   lxlsx_header_footer_options *options);

/**
 * @brief Set the horizontal page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_h_pagebreaks()` function adds horizontal page breaks to
 * a worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Horizontal page breaks act between rows.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxlsx_row_t and the last element of the array must be 0:
 *
 * @code
 *    lxlsx_row_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxlsx_row_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    lxlsx_worksheet_set_h_pagebreaks(worksheet1, breaks1);
 *    lxlsx_worksheet_set_h_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between rows 20 and 21 you must specify the break at
 * row 21. However in zero index notation this is actually row 20:
 *
 * @code
 *    // Break between row 20 and 21.
 *    lxlsx_row_t breaks[] = {20, 0};
 *
 *    lxlsx_worksheet_set_h_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `lxlsx_worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
lxlsx_error lxlsx_worksheet_set_h_pagebreaks(lxlsx_worksheet *worksheet,
                                     lxlsx_row_t breaks[]);

/**
 * @brief Set the vertical page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * @return A #lxlsx_error code.
 *
 * The `%lxlsx_worksheet_set_v_pagebreaks()` function adds vertical page breaks to a
 * worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Vertical page breaks act between columns.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxlsx_col_t and the last element of the array must be 0:
 *
 * @code
 *    lxlsx_col_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxlsx_col_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    lxlsx_worksheet_set_v_pagebreaks(worksheet1, breaks1);
 *    lxlsx_worksheet_set_v_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between columns 20 and 21 you must specify the break
 * at column 21. However in zero index notation this is actually column 20:
 *
 * @code
 *    // Break between column 20 and 21.
 *    lxlsx_col_t breaks[] = {20, 0};
 *
 *    lxlsx_worksheet_set_v_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 vertical page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `lxlsx_worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
lxlsx_error lxlsx_worksheet_set_v_pagebreaks(lxlsx_worksheet *worksheet,
                                     lxlsx_col_t breaks[]);

/**
 * @brief Set the order in which pages are printed.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * The `%lxlsx_worksheet_print_across()` function is used to change the default
 * print direction. This is referred to by Excel as the sheet "page order":
 *
 * @code
 *     lxlsx_worksheet_print_across(worksheet);
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
void lxlsx_worksheet_print_across(lxlsx_worksheet *worksheet);

/**
 * @brief Set the worksheet zoom factor.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param scale     Worksheet zoom factor.
 *
 * Set the worksheet zoom factor in the range `10 <= zoom <= 400`:
 *
 * @code
 *     lxlsx_worksheet_set_zoom(worksheet1, 50);
 *     lxlsx_worksheet_set_zoom(worksheet2, 75);
 *     lxlsx_worksheet_set_zoom(worksheet3, 300);
 *     lxlsx_worksheet_set_zoom(worksheet4, 400);
 * @endcode
 *
 * The default zoom factor is 100. It isn't possible to set the zoom to
 * "Selection" because it is calculated by Excel at run-time.
 *
 * Note, `%lxlsx_worksheet_zoom()` does not affect the scale of the printed
 * page. For that you should use `lxlsx_worksheet_set_print_scale()`.
 */
void lxlsx_worksheet_set_zoom(lxlsx_worksheet *worksheet, uint16_t scale);

/**
 * @brief Set the option to display or hide gridlines on the screen and
 *        the printed page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param option    Gridline option.
 *
 * Display or hide screen and print gridlines using one of the values of
 * @ref lxlsx_gridlines.
 *
 * @code
 *    lxlsx_worksheet_gridlines(worksheet1, LXLSX_HIDE_ALL_GRIDLINES);
 *
 *    lxlsx_worksheet_gridlines(worksheet2, LXLSX_SHOW_PRINT_GRIDLINES);
 * @endcode
 *
 * The Excel default is that the screen gridlines are on  and the printed
 * worksheet is off.
 *
 */
void lxlsx_worksheet_gridlines(lxlsx_worksheet *worksheet, uint8_t option);

/**
 * @brief Center the printed page horizontally.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * Center the worksheet data horizontally between the margins on the printed
 * page:
 *
 * @code
 *     lxlsx_worksheet_center_horizontally(worksheet);
 * @endcode
 *
 */
void lxlsx_worksheet_center_horizontally(lxlsx_worksheet *worksheet);

/**
 * @brief Center the printed page vertically.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * Center the worksheet data vertically between the margins on the printed
 * page:
 *
 * @code
 *     lxlsx_worksheet_center_vertically(worksheet);
 * @endcode
 *
 */
void lxlsx_worksheet_center_vertically(lxlsx_worksheet *worksheet);

/**
 * @brief Set the option to print the row and column headers on the printed
 *        page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * When printing a worksheet from Excel the row and column headers (the row
 * numbers on the left and the column letters at the top) aren't printed by
 * default.
 *
 * This function sets the printer option to print these headers:
 *
 * @code
 *    lxlsx_worksheet_print_row_col_headers(worksheet);
 * @endcode
 *
 */
void lxlsx_worksheet_print_row_col_headers(lxlsx_worksheet *worksheet);

/**
 * @brief Set the number of rows to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row First row of repeat range.
 * @param last_row  Last row of repeat range.
 *
 * @return A #lxlsx_error code.
 *
 * For large Excel documents it is often desirable to have the first row or
 * rows of the worksheet print out at the top of each page.
 *
 * This can be achieved by using this function. The parameters `first_row`
 * and `last_row` are zero based:
 *
 * @code
 *     lxlsx_worksheet_repeat_rows(worksheet, 0, 0); // Repeat the first row.
 *     lxlsx_worksheet_repeat_rows(worksheet, 0, 1); // Repeat the first two rows.
 * @endcode
 */
lxlsx_error lxlsx_worksheet_repeat_rows(lxlsx_worksheet *worksheet, lxlsx_row_t first_row,
                                lxlsx_row_t last_row);

/**
 * @brief Set the number of columns to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_col First column of repeat range.
 * @param last_col  Last column of repeat range.
 *
 * @return A #lxlsx_error code.
 *
 * For large Excel documents it is often desirable to have the first column or
 * columns of the worksheet print out at the left of each page.
 *
 * This can be achieved by using this function. The parameters `first_col`
 * and `last_col` are zero based:
 *
 * @code
 *     lxlsx_worksheet_repeat_columns(worksheet, 0, 0); // Repeat the first col.
 *     lxlsx_worksheet_repeat_columns(worksheet, 0, 1); // Repeat the first two cols.
 * @endcode
 */
lxlsx_error lxlsx_worksheet_repeat_columns(lxlsx_worksheet *worksheet,
                                   lxlsx_col_t first_col, lxlsx_col_t last_col);

/**
 * @brief Set the print area for a worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return A #lxlsx_error code.
 *
 * This function is used to specify the area of the worksheet that will be
 * printed. The RANGE() macro is often convenient for this.
 *
 * @code
 *     lxlsx_worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     lxlsx_worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 *
 * In order to set a row or column range you must specify the entire range:
 *
 * @code
 *     lxlsx_worksheet_print_area(worksheet, RANGE("A1:H1048576")); // Same as A:H.
 * @endcode
 */
lxlsx_error lxlsx_worksheet_print_area(lxlsx_worksheet *worksheet, lxlsx_row_t first_row,
                               lxlsx_col_t first_col, lxlsx_row_t last_row,
                               lxlsx_col_t last_col);
/**
 * @brief Fit the printed area to a specific number of pages both vertically
 *        and horizontally.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param width     Number of pages horizontally.
 * @param height    Number of pages vertically.
 *
 * The `%lxlsx_worksheet_fit_to_pages()` function is used to fit the printed area to
 * a specific number of pages both vertically and horizontally. If the printed
 * area exceeds the specified number of pages it will be scaled down to
 * fit. This ensures that the printed area will always appear on the specified
 * number of pages even if the page size or margins change:
 *
 * @code
 *     lxlsx_worksheet_fit_to_pages(worksheet1, 1, 1); // Fit to 1x1 pages.
 *     lxlsx_worksheet_fit_to_pages(worksheet2, 2, 1); // Fit to 2x1 pages.
 *     lxlsx_worksheet_fit_to_pages(worksheet3, 1, 2); // Fit to 1x2 pages.
 * @endcode
 *
 * The print area can be defined using the `lxlsx_worksheet_print_area()` function
 * as described above.
 *
 * A common requirement is to fit the printed output to `n` pages wide but
 * have the height be as long as necessary. To achieve this set the `height`
 * to zero:
 *
 * @code
 *     // 1 page wide and as long as necessary.
 *     lxlsx_worksheet_fit_to_pages(worksheet, 1, 0);
 * @endcode
 *
 * **Note**:
 *
 * - Although it is valid to use both `%lxlsx_worksheet_fit_to_pages()` and
 *   `lxlsx_worksheet_set_print_scale()` on the same worksheet Excel only allows one
 *   of these options to be active at a time. The last function call made will
 *   set the active option.
 *
 * - The `%lxlsx_worksheet_fit_to_pages()` function will override any manual page
 *   breaks that are defined in the worksheet.
 *
 * - When using `%lxlsx_worksheet_fit_to_pages()` it may also be required to set the
 *   printer paper size using `lxlsx_worksheet_set_paper()` or else Excel will
 *   default to "US Letter".
 *
 */
void lxlsx_worksheet_fit_to_pages(lxlsx_worksheet *worksheet, uint16_t width,
                            uint16_t height);

/**
 * @brief Set the start/first page number when printing.
 *
 * @param worksheet  Pointer to a lxlsx_worksheet instance to be updated.
 * @param start_page Page number of the starting page when printing.
 *
 * The `%lxlsx_worksheet_set_start_page()` function is used to set the number number
 * of the first page when the worksheet is printed out. It is the same as the
 * "First Page Number" option in Excel:
 *
 * @code
 *     // Start print from page 2.
 *     lxlsx_worksheet_set_start_page(worksheet, 2);
 * @endcode
 */
void lxlsx_worksheet_set_start_page(lxlsx_worksheet *worksheet, uint16_t start_page);

/**
 * @brief Set the scale factor for the printed page.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param scale     Print scale of worksheet to be printed.
 *
 * This function sets the scale factor of the printed page. The Scale factor
 * must be in the range `10 <= scale <= 400`:
 *
 * @code
 *     lxlsx_worksheet_set_print_scale(worksheet1, 75);
 *     lxlsx_worksheet_set_print_scale(worksheet2, 400);
 * @endcode
 *
 * The default scale factor is 100. Note, `%lxlsx_worksheet_set_print_scale()` does
 * not affect the scale of the visible page in Excel. For that you should use
 * `lxlsx_worksheet_set_zoom()`.
 *
 * Note that although it is valid to use both `lxlsx_worksheet_fit_to_pages()` and
 * `%lxlsx_worksheet_set_print_scale()` on the same worksheet Excel only allows one
 * of these options to be active at a time. The last function call made will
 * set the active option.
 *
 */
void lxlsx_worksheet_set_print_scale(lxlsx_worksheet *worksheet, uint16_t scale);

/**
 * @brief Set the worksheet to print in black and white
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * Set the option to print the worksheet in black and white:
 * @code
 *     lxlsx_worksheet_print_black_and_white(worksheet);
 * @endcode
 */
void lxlsx_worksheet_print_black_and_white(lxlsx_worksheet *worksheet);

/**
 * @brief Display the worksheet cells from right to left for some versions of
 *        Excel.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
  * The `%lxlsx_worksheet_right_to_left()` function is used to change the default
 * direction of the worksheet from left-to-right, with the `A1` cell in the
 * top left, to right-to-left, with the `A1` cell in the top right.
 *
 * @code
 *     lxlsx_worksheet_right_to_left(worksheet1);
 * @endcode
 *
 * This is useful when creating Arabic, Hebrew or other near or far eastern
 * worksheets that use right-to-left as the default direction.
 */
void lxlsx_worksheet_right_to_left(lxlsx_worksheet *worksheet);

/**
 * @brief Hide zero values in worksheet cells.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 *
 * The `%lxlsx_worksheet_hide_zero()` function is used to hide any zero values that
 * appear in cells:
 *
 * @code
 *     lxlsx_worksheet_hide_zero(worksheet1);
 * @endcode
 */
void lxlsx_worksheet_hide_zero(lxlsx_worksheet *worksheet);

/**
 * @brief Set the color of the worksheet tab.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param color     The tab color.
 *
 * The `%lxlsx_worksheet_set_tab_color()` function is used to change the color of
 * the worksheet tab:
 *
 * @code
 *      lxlsx_worksheet_set_tab_color(worksheet1, LXLSX_COLOR_RED);
 *      lxlsx_worksheet_set_tab_color(worksheet2, LXLSX_COLOR_GREEN);
 *      lxlsx_worksheet_set_tab_color(worksheet3, 0xFF9900); // Orange.
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
void lxlsx_worksheet_set_tab_color(lxlsx_worksheet *worksheet, lxlsx_color_t color);

/**
 * @brief Protect elements of a worksheet from modification.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance to be updated.
 * @param password  A worksheet password.
 * @param options   Worksheet elements to protect.
 *
 * The `%lxlsx_worksheet_protect()` function protects worksheet elements from modification:
 *
 * @code
 *     lxlsx_worksheet_protect(worksheet, "Some Password", options);
 * @endcode
 *
 * The `password` and lxlsx_protection pointer are both optional:
 *
 * @code
 *     lxlsx_worksheet_protect(worksheet1, NULL,       NULL);
 *     lxlsx_worksheet_protect(worksheet2, NULL,       my_options);
 *     lxlsx_worksheet_protect(worksheet3, "password", NULL);
 *     lxlsx_worksheet_protect(worksheet4, "password", my_options);
 * @endcode
 *
 * Passing a `NULL` password is the same as turning on protection without a
 * password. Passing a `NULL` password and `NULL` options, or any other
 * combination has the effect of enabling a cell's `locked` and `hidden`
 * properties if they have been set.
 *
 * A *locked* cell cannot be edited and this property is on by default for all
 * cells. A *hidden* cell will display the results of a formula but not the
 * formula itself. These properties can be set using the lxlsx_format_set_unlocked()
 * and lxlsx_format_set_hidden() format functions.
 *
 * You can specify which worksheet elements you wish to protect by passing a
 * lxlsx_protection pointer in the `options` argument with any or all of the
 * following members set:
 *
 *     no_select_locked_cells
 *     no_select_unlocked_cells
 *     lxlsx_format_cells
 *     lxlsx_format_columns
 *     lxlsx_format_rows
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
 *     lxlsx_protection options = {
 *         .lxlsx_format_cells             = 1,
 *         .insert_hyperlinks        = 1,
 *         .insert_rows              = 1,
 *         .delete_rows              = 1,
 *         .insert_columns           = 1,
 *         .delete_columns           = 1,
 *     };
 *
 *     lxlsx_worksheet_protect(worksheet, NULL, &options);
 *
 * @endcode
 *
 * See also the lxlsx_format_set_unlocked() and lxlsx_format_set_hidden() format functions.
 *
 * **Note:** Sheet level passwords in Excel offer **very** weak
 * protection. They don't encrypt your data and are very easy to
 * deactivate. Full workbook encryption is not supported by `libxlsxwriter`
 * since it requires a completely different file format.
 */
void lxlsx_worksheet_protect(lxlsx_worksheet *worksheet, const char *password,
                       lxlsx_protection *options);

/**
 * @brief Set the Outline and Grouping display properties.
 *
 * @param worksheet      Pointer to a lxlsx_worksheet instance to be updated.
 * @param visible        Outlines are visible. Optional, defaults to True.
 * @param symbols_below  Show row outline symbols below the outline bar.
 * @param symbols_right  Show column outline symbols to the right of outline.
 * @param auto_style     Use Automatic outline style.
 *
 * The `%lxlsx_worksheet_outline_settings()` method is used to control the
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
 *     lxlsx_worksheet_outline_settings(worksheet1, LXLSX_TRUE, LXLSX_TRUE, LXLSX_TRUE, LXLSX_FALSE);
 * @endcode
 *
 * The worksheet parameters controlled by `lxlsx_worksheet_outline_settings()` are
 * rarely used.
 */
void lxlsx_worksheet_outline_settings(lxlsx_worksheet *worksheet, uint8_t visible,
                                uint8_t symbols_below, uint8_t symbols_right,
                                uint8_t auto_style);

/**
 * @brief Set the default row properties.
 *
 * @param worksheet        Pointer to a lxlsx_worksheet instance to be updated.
 * @param height           Default row height.
 * @param hide_unused_rows Hide unused cells.
 *
 * The `%lxlsx_worksheet_set_default_row()` function is used to set Excel default
 * row properties such as the default height and the option to hide unused
 * rows. These parameters are an optimization used by Excel to set row
 * properties without generating a very large file with an entry for each row.
 *
 * To set the default row height:
 *
 * @code
 *     lxlsx_worksheet_set_default_row(worksheet, 24, LXLSX_FALSE);
 *
 * @endcode
 *
 * To hide unused rows:
 *
 * @code
 *     lxlsx_worksheet_set_default_row(worksheet, 15, LXLSX_TRUE);
 * @endcode
 *
 * Note, in the previous case we use the default height #LXLSX_DEF_ROW_HEIGHT =
 * 15 so the the height remains unchanged.
 */
void lxlsx_worksheet_set_default_row(lxlsx_worksheet *worksheet, double height,
                               uint8_t hide_unused_rows);

/**
 * @brief Set the VBA name for the worksheet.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance.
 * @param name      Name of the worksheet used by VBA.
 *
 * @return A #lxlsx_error.
 *
 * The `lxlsx_worksheet_set_vba_name()` function can be used to set the VBA name for
 * the worksheet. This is sometimes required when a vbaProject macro included
 * via `lxlsx_workbook_add_vba_project()` refers to the worksheet by a name other
 * than the worksheet name:
 *
 * @code
 *     lxlsx_workbook_set_vba_name (workbook,  "MyWorkbook");
 *     lxlsx_worksheet_set_vba_name(worksheet, "MySheet1");
 * @endcode
 *
 * In general Excel uses the worksheet name such as "Sheet1" as the VBA name.
 * However, this can be changed in the VBA environment or if the the macro was
 * extracted from a foreign language version of Excel.
 *
 * See also @ref working_with_macros
 */
lxlsx_error lxlsx_worksheet_set_vba_name(lxlsx_worksheet *worksheet, const char *name);

/**
 * @brief Make all comments in the worksheet visible.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance.
 *
 * This `%lxlsx_worksheet_show_comments()` function is used to make all cell
 * comments visible when a worksheet is opened:
 *
 * @code
 *     lxlsx_worksheet_show_comments(worksheet);
 * @endcode
 *
 * Individual comments can be made visible or hidden using the `visible`
 * option of the #lxlsx_comment_options struct and the `lxlsx_worksheet_write_comment_opt()`
 * function (see above and @ref ww_comments_visible).
 */
lxlsx_error lxlsx_worksheet_show_comments(lxlsx_worksheet *worksheet);

/**
 * @brief Set the default author of the cell comments.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance.
 * @param author    The name of the comment author.
 *
 * This `%lxlsx_worksheet_set_comments_author()` function is used to set the
 * default author of all cell comments:
 *
 * @code
 *     lxlsx_worksheet_set_comments_author(worksheet, "Jane Gloriana Villanueva")
 * @endcode
 *
 * Individual authors can be set using the `author` option of the
 * #lxlsx_comment_options struct and the `lxlsx_worksheet_write_comment_opt()`
 * function (see above and @ref ww_comments_author).
 */
void lxlsx_worksheet_set_comments_author(lxlsx_worksheet *worksheet,
                                   const char *author);

/**
 * @brief Ignore various Excel errors/warnings in a worksheet for user
 *        defined ranges.
 *
 * @param worksheet Pointer to a lxlsx_worksheet instance.
 * @param type      The type of error/warning to ignore. See #lxlsx_ignore_errors.
 * @param range     The range(s) for which the error/warning should be ignored.
 *
 * @return A #lxlsx_error.
 *
 *
 * The `%lxlsx_worksheet_ignore_errors()` function can be used to ignore various
 * worksheet cell errors/warnings. For example the following code writes a string
 * that looks like a number:
 *
 * @code
 *     lxlsx_worksheet_write_string(worksheet, CELL("D2"), "123", NULL);
 * @endcode
 *
 * This causes Excel to display a small green triangle in the top left hand
 * corner of the cell to indicate an error/warning:
 *
 * @image html ignore_errors1.png
 *
 * Sometimes these warnings are useful indicators that there is an issue in
 * the spreadsheet but sometimes it is preferable to turn them off. Warnings
 * can be turned off at the Excel level for all workbooks and worksheets by
 * using the using "Excel options -> Formulas -> Error checking
 * rules". Alternatively you can turn them off for individual cells in a
 * worksheet, or ranges of cells, using the `%lxlsx_worksheet_ignore_errors()`
 * function with different #lxlsx_ignore_errors options and ranges like this:
 *
 * @code
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "C3");
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_EVAL_ERROR,            "C6");
 * @endcode
 *
 * The range can be a single cell, a range of cells, or multiple cells and ranges
 * separated by spaces:
 *
 * @code
 *     // Single cell.
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "C6");
 *
 *     // Or a single range:
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "C6:G8");
 *
 *     // Or multiple cells and ranges:
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "C6 E6 G1:G20 J2:J6");
 * @endcode
 *
 * @note Calling `%lxlsx_worksheet_ignore_errors()` more than once for the same
 * #lxlsx_ignore_errors type will overwrite the previous range.
 *
 * You can turn off warnings for an entire column by specifying the range from
 * the first cell in the column to the last cell in the column:
 *
 * @code
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "A1:A1048576");
 * @endcode
 *
 * Or for the entire worksheet by specifying the range from the first cell in
 * the worksheet to the last cell in the worksheet:
 *
 * @code
 *     lxlsx_worksheet_ignore_errors(worksheet, LXLSX_IGNORE_NUMBER_STORED_AS_TEXT, "A1:XFD1048576");
 * @endcode
 *
 * The worksheet errors/warnings that can be ignored are:
 *
 * - #LXLSX_IGNORE_NUMBER_STORED_AS_TEXT: Turn off errors/warnings for numbers
 *    stores as text.
 *
 * - #LXLSX_IGNORE_EVAL_ERROR: Turn off errors/warnings for formula errors (such
 *    as divide by zero).
 *
 * - #LXLSX_IGNORE_FORMULA_DIFFERS: Turn off errors/warnings for formulas that
 *    differ from surrounding formulas.
 *
 * - #LXLSX_IGNORE_FORMULA_RANGE: Turn off errors/warnings for formulas that
 *    omit cells in a range.
 *
 * - #LXLSX_IGNORE_FORMULA_UNLOCKED: Turn off errors/warnings for unlocked cells
 *    that contain formulas.
 *
 * - #LXLSX_IGNORE_EMPTY_CELL_REFERENCE: Turn off errors/warnings for formulas
 *    that refer to empty cells.
 *
 * - #LXLSX_IGNORE_LIST_DATA_VALIDATION: Turn off errors/warnings for cells in a
 *    table that do not comply with applicable data validation rules.
 *
 * - #LXLSX_IGNORE_CALCULATED_COLUMN: Turn off errors/warnings for cell formulas
 *    that differ from the column formula.
 *
 * - #LXLSX_IGNORE_TWO_DIGIT_TEXT_YEAR: Turn off errors/warnings for formulas
 *    that contain a two digit text representation of a year.
 *
 */
lxlsx_error lxlsx_worksheet_ignore_errors(lxlsx_worksheet *worksheet, uint8_t type,
                                  const char *range);

lxlsx_worksheet *lxlsx_worksheet_new(lxlsx_worksheet_init_data *init_data);
void lxlsx_worksheet_free(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_assemble_xml_file(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_single_row(lxlsx_worksheet *worksheet);

void lxlsx_worksheet_prepare_image(lxlsx_worksheet *worksheet,
                                 uint32_t image_ref_id, uint32_t lxlsx_drawing_id,
                                 lxlsx_object_properties *object_props);

void lxlsx_worksheet_prepare_header_image(lxlsx_worksheet *worksheet,
                                        uint32_t image_ref_id,
                                        lxlsx_object_properties *object_props);

void lxlsx_worksheet_prepare_background(lxlsx_worksheet *worksheet,
                                      uint32_t image_ref_id,
                                      lxlsx_object_properties *object_props);

void lxlsx_worksheet_prepare_chart(lxlsx_worksheet *worksheet,
                                 uint32_t lxlsx_chart_ref_id, uint32_t lxlsx_drawing_id,
                                 lxlsx_object_properties *object_props,
                                 uint8_t is_chartsheet);

uint32_t lxlsx_worksheet_prepare_vml_objects(lxlsx_worksheet *worksheet,
                                           uint32_t lxlsx_vml_data_id,
                                           uint32_t lxlsx_vml_shape_id,
                                           uint32_t lxlsx_vml_drawing_id,
                                           uint32_t comment_id);

void lxlsx_worksheet_prepare_header_vml_objects(lxlsx_worksheet *worksheet,
                                              uint32_t lxlsx_vml_header_id,
                                              uint32_t lxlsx_vml_drawing_id);

void lxlsx_worksheet_prepare_tables(lxlsx_worksheet *worksheet,
                                  uint32_t lxlsx_table_id);

lxlsx_row *lxlsx_worksheet_find_row(lxlsx_worksheet *worksheet, lxlsx_row_t row_num);
lxlsx_cell *lxlsx_worksheet_find_cell_in_row(lxlsx_row *row, lxlsx_col_t col_num);
/*
 * External functions to call intern XML functions shared with chartsheet.
 */
void lxlsx_worksheet_write_sheet_views(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_page_margins(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_drawings(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_sheet_protection(lxlsx_worksheet *worksheet,
                                          lxlsx_protection_obj *protect);
void lxlsx_worksheet_write_sheet_pr(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_page_setup(lxlsx_worksheet *worksheet);
void lxlsx_worksheet_write_header_footer(lxlsx_worksheet *worksheet);

void lxlsx_worksheet_set_error_cell(lxlsx_worksheet *worksheet,
                              lxlsx_object_properties *object_props,
                              uint32_t ref_id);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _worksheet_xml_declaration(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_worksheet(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_dimension(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_view(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_views(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_format_pr(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_data(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_page_margins(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_page_setup(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_col_info(lxlsx_worksheet *worksheet,
                                      lxlsx_col_options *options);
STATIC void _write_row(lxlsx_worksheet *worksheet, lxlsx_row *row, char *spans);
STATIC lxlsx_row *_get_row_list(struct lxlsx_table_rows *table,
                              lxlsx_row_t row_num);

STATIC void _worksheet_write_merge_cell(lxlsx_worksheet *worksheet,
                                        lxlsx_merged_range *merged_range);
STATIC void _worksheet_write_merge_cells(lxlsx_worksheet *worksheet);

STATIC void _worksheet_write_odd_header(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_odd_footer(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_header_footer(lxlsx_worksheet *worksheet);

STATIC void _worksheet_write_print_options(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_pr(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_tab_color(lxlsx_worksheet *worksheet);
STATIC void _worksheet_write_sheet_protection(lxlsx_worksheet *worksheet,
                                              lxlsx_protection_obj *protect);
STATIC void _worksheet_write_data_validations(lxlsx_worksheet *self);

STATIC double _pixels_to_height(double pixels);
STATIC double _pixels_to_width(double pixels);

STATIC void _worksheet_write_auto_filter(lxlsx_worksheet *worksheet);
#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_WORKSHEET_H__ */


/* XLSX worksheet read API. */

#ifndef LXLSX_READER_WORKSHEET_H
#define LXLSX_READER_WORKSHEET_H

#include "libxlsx/common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_worksheet lxlsx_reader_worksheet;

void lxlsx_reader_worksheet_close(lxlsx_reader_worksheet *ws);

lxlsx_reader_error lxlsx_reader_worksheet_next_row (lxlsx_reader_worksheet *ws);
/* The returned lxlsx_cell borrows string/formula buffers owned by ws. Treat
 * the cell contents, including data.reader.value.formula, as valid only until
 * the next reader call that advances or closes the worksheet. */
lxlsx_reader_error lxlsx_reader_worksheet_next_cell(lxlsx_reader_worksheet *ws, lxlsx_cell *out);

size_t   lxlsx_reader_worksheet_current_row    (const lxlsx_reader_worksheet *ws);
size_t   lxlsx_reader_worksheet_max_column_seen(const lxlsx_reader_worksheet *ws);
uint32_t lxlsx_reader_worksheet_flags          (const lxlsx_reader_worksheet *ws);

typedef int (*lxlsx_reader_cell_cb)   (const lxlsx_cell *cell, void *userdata);
typedef int (*lxlsx_reader_row_end_cb)(size_t row, size_t max_col, void *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_process(lxlsx_reader_worksheet *ws,
                                lxlsx_reader_cell_cb    cell_cb,
                                lxlsx_reader_row_end_cb row_cb,
                                void          *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_skip_rows(lxlsx_reader_worksheet *ws, size_t n);

/* ----- Phase 1 metadata accessors ---------------------------------------- */

/* Merged cells. Index is 0-based. */
size_t lxlsx_reader_worksheet_merged_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_merged_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                  lxlsx_reader_range *out);

/* Returns 1 when (row, col) (1-based) lies inside a merged range and is NOT
 * the merge's master cell; 0 otherwise. Useful for honoring
 * LXLSX_READER_SKIP_MERGED_FOLLOW in row-shaped readers. */
int lxlsx_reader_worksheet_in_merge_follow(lxlsx_reader_worksheet *ws, size_t row, size_t col);

/* Hyperlinks. The url/tooltip/display/location pointers (if non-NULL) are
 * owned by the worksheet and remain valid until lxlsx_reader_worksheet_close().
 * Internal links (e.g. "Sheet2!A1") have location set and url == NULL. */
typedef struct {
    lxlsx_reader_range range;
    const char *url;       /* external URL, or NULL */
    const char *location;  /* internal anchor, or NULL */
    const char *display;   /* display text, or NULL */
    const char *tooltip;   /* tooltip text, or NULL */
} lxlsx_reader_hyperlink;

size_t      lxlsx_reader_worksheet_hyperlink_count (const lxlsx_reader_worksheet *ws);
int         lxlsx_reader_worksheet_hyperlink_get   (const lxlsx_reader_worksheet *ws, size_t idx,
                                           lxlsx_reader_hyperlink *out);
const char *lxlsx_reader_worksheet_hyperlink_url   (const lxlsx_reader_worksheet *ws,
                                           size_t row, size_t col);

/* Sheet protection. Booleans use 1/0. password_hash is empty when absent.
 * Per ECMA-376 the attribute defaults differ per field; we report only what
 * was *present* in the XML, plus a boolean is_present for the element. */
typedef struct {
    int  is_present;
    char password_hash[20];
    int  sheet;
    int  content;
    int  objects;
    int  scenarios;
    int  format_cells;
    int  format_columns;
    int  format_rows;
    int  insert_columns;
    int  insert_rows;
    int  insert_hyperlinks;
    int  delete_columns;
    int  delete_rows;
    int  select_locked_cells;
    int  sort;
    int  auto_filter;
    int  pivot_tables;
    int  select_unlocked_cells;
} lxlsx_reader_protection;

int lxlsx_reader_worksheet_protection(const lxlsx_reader_worksheet *ws, lxlsx_reader_protection *out);

/* Row / column metadata. row/col indexes are 1-based. */
typedef struct {
    int    has_height;
    double height;
    int    hidden;
    int    outline_level;
    int    collapsed;
    int    custom_height;
} lxlsx_reader_row_options;

typedef struct {
    int    has_width;
    double width;
    int    hidden;
    int    outline_level;
    int    collapsed;
} lxlsx_reader_col_options;

int lxlsx_reader_worksheet_row_options(const lxlsx_reader_worksheet *ws, size_t row,
                              lxlsx_reader_row_options *out);
int lxlsx_reader_worksheet_col_options(const lxlsx_reader_worksheet *ws, size_t col,
                              lxlsx_reader_col_options *out);

int    lxlsx_reader_worksheet_default_row_height(const lxlsx_reader_worksheet *ws, double *out);
int    lxlsx_reader_worksheet_default_col_width (const lxlsx_reader_worksheet *ws, double *out);

/* ---- Data validation ---------------------------------------------------- */

/* All string pointers are owned by the worksheet; valid until close. */
typedef struct {
    const char *type;          /* "whole","decimal","list","date","time",
                                  "textLength","custom" — or NULL */
    const char *operator_;     /* "between","notBetween","equal", ... — or NULL */
    const char *error_style;   /* "stop","warning","information" — or NULL */
    const char *formula1;
    const char *formula2;
    const char *prompt;
    const char *prompt_title;
    const char *error;
    const char *error_title;
    const char *sqref;         /* original sqref (may be space-separated list) */
    int allow_blank;
    int show_drop_down;        /* per OOXML this is INVERTED — 1 means "hide" */
    int show_input_message;
    int show_error_message;
} lxlsx_reader_data_validation;

size_t lxlsx_reader_worksheet_data_validation_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_data_validation_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                           lxlsx_reader_data_validation *out);

/* ---- AutoFilter --------------------------------------------------------- */

typedef enum {
    LXLSX_READER_FILTER_NONE = 0,
    LXLSX_READER_FILTER_LIST,           /* discrete value match (<filters>) */
    LXLSX_READER_FILTER_CUSTOM,         /* operator-based (<customFilters>) */
    LXLSX_READER_FILTER_TOP10,          /* <top10>: top/bottom N (or %) */
    LXLSX_READER_FILTER_DYNAMIC         /* <dynamicFilter>: aboveAverage etc. */
} lxlsx_reader_filter_kind;

typedef struct {
    int             col_id;          /* 0-based column offset within the autoFilter range */
    lxlsx_reader_filter_kind kind;
    /* For LXLSX_READER_FILTER_LIST: list of value strings. */
    const char    **values;          /* NULL-terminated; NULL when not applicable */
    /* For LXLSX_READER_FILTER_CUSTOM: at most two predicates joined by AND/OR. */
    int             custom_and;      /* 1 = AND, 0 = OR */
    const char     *custom_op_1;     /* "equal","greaterThan", ... */
    const char     *custom_val_1;
    const char     *custom_op_2;
    const char     *custom_val_2;
    /* For LXLSX_READER_FILTER_TOP10. */
    int             top;             /* 1=top, 0=bottom */
    int             percent;         /* 1=percent */
    double          top_value;       /* N */
} lxlsx_reader_filter_column;

typedef struct {
    const char              *range;       /* original ref attr (e.g. "A1:Z100"), or NULL */
    const lxlsx_reader_filter_column *columns;
    size_t                   columns_count;
} lxlsx_reader_autofilter;

int lxlsx_reader_worksheet_autofilter(const lxlsx_reader_worksheet *ws, lxlsx_reader_autofilter *out);

/* ---- Page setup --------------------------------------------------------- */

typedef struct {
    /* Margins (inches). has_* indicate explicit element presence. */
    int    has_margins;
    double margin_left, margin_right, margin_top, margin_bottom;
    double margin_header, margin_footer;

    /* pageSetup attrs. */
    int    has_setup;
    int    paper_size;          /* 0 if absent */
    int    fit_to_width;
    int    fit_to_height;
    int    scale;               /* 0 if absent (interpreted as default 100) */
    int    orientation_landscape;  /* 1 = landscape, 0 = portrait/default */
    int    horizontal_dpi;
    int    vertical_dpi;
    int    first_page_number;
    int    use_first_page_number;

    /* printOptions. */
    int    print_horizontal_centered;
    int    print_vertical_centered;
    int    print_grid_lines;
    int    print_headings;

    /* Header / footer text (strings; pointers owned by worksheet). */
    const char *odd_header;
    const char *odd_footer;
    const char *even_header;
    const char *even_footer;
    const char *first_header;
    const char *first_footer;
    int   different_odd_even;
    int   different_first;
    int   scale_with_doc;
    int   align_with_margins;
} lxlsx_reader_page_setup;

int lxlsx_reader_worksheet_page_setup(const lxlsx_reader_worksheet *ws, lxlsx_reader_page_setup *out);

/* ---- Conditional formats (§8.2.3) -------------------------------------- */

typedef struct {
    const char *type;       /* "cellIs","expression","colorScale","dataBar",
                                "iconSet","top10","aboveAverage","duplicateValues",
                                "uniqueValues","containsText","containsBlanks",
                                "notContainsBlanks","containsErrors","notContainsErrors",
                                "timePeriod" — or NULL */
    const char *operator_;  /* "greaterThan","between", ... — or NULL */
    int         priority;
    int         stop_if_true;
    int         dxf_id;     /* -1 if absent */
    int         percent;    /* top10's percent attr */
    int         bottom;     /* top10's bottom attr */
    double      rank;       /* top10's rank attr */
    const char *text;       /* containsText/notContainsText/etc. */
    const char *time_period; /* timePeriod's attr */
    /* up to 2 inline formulas (formula1, formula2) */
    const char *formula1;
    const char *formula2;
} lxlsx_reader_cf_rule;

typedef struct {
    const char        *sqref;     /* the conditionalFormatting sqref */
    const lxlsx_reader_cf_rule *rules;
    size_t             rules_count;
} lxlsx_reader_cf_block;

size_t lxlsx_reader_worksheet_cf_block_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_cf_block_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                    lxlsx_reader_cf_block *out);

/* ---- Comments (§8.2.1) -------------------------------------------------- */

typedef struct {
    size_t      row;          /* 1-based */
    size_t      col;          /* 1-based */
    const char *text;
    const char *author;
    int         visible;      /* legacy VML's "visible" attr */
    int         threaded;
    const char *parent_id;    /* threaded reply parent GUID, or NULL */
} lxlsx_reader_comment_info;

typedef int (*lxlsx_reader_comment_cb)(const lxlsx_reader_comment_info *info, void *userdata);

/* Iterates legacy + threaded comments. cb returning non-zero stops. Returns
 * LXLSX_READER_NO_ERROR on success; ZIP_ENTRY_NOT_FOUND when the sheet has none. */
lxlsx_reader_error lxlsx_reader_worksheet_iterate_comments(lxlsx_reader_worksheet *ws,
                                         lxlsx_reader_comment_cb cb, void *userdata);

/* ---- Chart metadata (§8.2.4) ------------------------------------------- */

typedef struct {
    const char *name;        /* series name, may be NULL */
    const char *categories;  /* category range / formula */
    const char *values;      /* values range */
} lxlsx_reader_chart_series_info;

typedef struct {
    const char *type;        /* "bar"/"line"/"pie"/"area"/"scatter"/"radar"/"doughnut" or NULL */
    const char *title;
    /* 1-based inclusive anchor; 0 if not from a twoCellAnchor. */
    size_t from_row, from_col, to_row, to_col;
    const lxlsx_reader_chart_series_info *series;
    size_t                       series_count;
} lxlsx_reader_chart_meta;

typedef int (*lxlsx_reader_chart_cb)(const lxlsx_reader_chart_meta *info, void *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_iterate_charts(lxlsx_reader_worksheet *ws,
                                       lxlsx_reader_chart_cb cb, void *userdata);

/* ---- Rich-text runs (§8.2.2) ------------------------------------------- */

/* For STRING_CELL (SST) and INLINE_STRING_CELL cells, fills `out`
 * with up to `cap` runs and returns the actual count. When `out` is NULL
 * just returns the run count. Callers can sniff the run count first, then
 * allocate appropriately. The pointers inside each run remain valid until
 * the next streaming step that overwrites the inline buffer (for inline
 * strings) or until the workbook is closed (for SST strings). */
size_t lxlsx_reader_cell_string_runs(const lxlsx_reader_worksheet *ws, const lxlsx_cell *c,
                            lxlsx_reader_string_run *out, size_t cap);

#ifdef __cplusplus
}
#endif

#endif
