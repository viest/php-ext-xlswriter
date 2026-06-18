/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * chart - A libxlsxwriter library for creating Excel XLSX chart files.
 *
 */

/**
 * @page lxlsx_chart_page The Chart object
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * See @ref chart.h for full details of the functionality.
 *
 * @file chart.h
 *
 * @brief Functions related to adding data to and configuring  a chart.
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * A Chart object isn't created directly. Instead a chart is created by
 * calling the `lxlsx_workbook_add_chart()` function from a Workbook object. For
 * example:
 *
 * @code
 *
 * #include "libxlsx.h"
 *
 * int main() {
 *
 *     lxlsx_workbook  *workbook  = new_workbook("chart.xlsx");
 *     lxlsx_worksheet *worksheet = lxlsx_workbook_add_worksheet(workbook, NULL);
 *
 *     // User function to add data to worksheet, not shown here.
 *     write_worksheet_data(worksheet);
 *
 *     // Create a chart object.
 *     lxlsx_chart *chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_COLUMN);
 *
 *     // In the simplest case we just add some value data series.
 *     // The NULL categories will default to 1 to 5 like in Excel.
 *     lxlsx_chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$5");
 *     lxlsx_chart_add_series(chart, NULL, "=Sheet1!$B$1:$B$5");
 *     lxlsx_chart_add_series(chart, NULL, "=Sheet1!$C$1:$C$5");
 *
 *     // Insert the chart into the worksheet
 *     lxlsx_worksheet_insert_chart(worksheet, CELL("B7"), chart);
 *
 *     return lxlsx_workbook_close(workbook);
 * }
 *
 * @endcode
 *
 * The chart in the worksheet will look like this:
 * @image html lxlsx_chart_simple.png
 *
 * The basic procedure for adding a chart to a worksheet is:
 *
 * 1. Create the chart with `lxlsx_workbook_add_chart()`.
 * 2. Add one or more data series to the chart which refers to data in the
 *    workbook using `lxlsx_chart_add_series()`.
 * 3. Configure the chart with the other available functions shown below.
 * 4. Insert the chart into a worksheet using `lxlsx_worksheet_insert_chart()`.
 *
 */

#ifndef __LXLSX_CHART_H__
#define __LXLSX_CHART_H__

#include <stdint.h>
#include <string.h>

#include "common.h"
#include "format.h"

STAILQ_HEAD(lxlsx_chart_series_list, lxlsx_chart_series);
STAILQ_HEAD(lxlsx_series_data_points, lxlsx_series_data_point);

#define LXLSX_CHART_NUM_FORMAT_LEN 128
#define LXLSX_CHART_DEFAULT_GAP 501

/**
 * @brief Available chart types.
 */
typedef enum lxlsx_chart_type {

    /** None. */
    LXLSX_CHART_NONE = 0,

    /** Area chart. */
    LXLSX_CHART_AREA,

    /** Area chart - stacked. */
    LXLSX_CHART_AREA_STACKED,

    /** Area chart - percentage stacked. */
    LXLSX_CHART_AREA_STACKED_PERCENT,

    /** Bar chart. */
    LXLSX_CHART_BAR,

    /** Bar chart - stacked. */
    LXLSX_CHART_BAR_STACKED,

    /** Bar chart - percentage stacked. */
    LXLSX_CHART_BAR_STACKED_PERCENT,

    /** Column chart. */
    LXLSX_CHART_COLUMN,

    /** Column chart - stacked. */
    LXLSX_CHART_COLUMN_STACKED,

    /** Column chart - percentage stacked. */
    LXLSX_CHART_COLUMN_STACKED_PERCENT,

    /** Doughnut chart. */
    LXLSX_CHART_DOUGHNUT,

    /** Line chart. */
    LXLSX_CHART_LINE,

    /** Line chart - stacked. */
    LXLSX_CHART_LINE_STACKED,

    /** Line chart - percentage stacked. */
    LXLSX_CHART_LINE_STACKED_PERCENT,

    /** Pie chart. */
    LXLSX_CHART_PIE,

    /** Scatter chart. */
    LXLSX_CHART_SCATTER,

    /** Scatter chart - straight. */
    LXLSX_CHART_SCATTER_STRAIGHT,

    /** Scatter chart - straight with markers. */
    LXLSX_CHART_SCATTER_STRAIGHT_WITH_MARKERS,

    /** Scatter chart - smooth. */
    LXLSX_CHART_SCATTER_SMOOTH,

    /** Scatter chart - smooth with markers. */
    LXLSX_CHART_SCATTER_SMOOTH_WITH_MARKERS,

    /** Radar chart. */
    LXLSX_CHART_RADAR,

    /** Radar chart - with markers. */
    LXLSX_CHART_RADAR_WITH_MARKERS,

    /** Radar chart - filled. */
    LXLSX_CHART_RADAR_FILLED
} lxlsx_chart_type;

/**
 * @brief Chart legend positions.
 */
typedef enum lxlsx_chart_legend_position {

    /** No chart legend. */
    LXLSX_CHART_LEGEND_NONE = 0,

    /** Chart legend positioned at right side. */
    LXLSX_CHART_LEGEND_RIGHT,

    /** Chart legend positioned at left side. */
    LXLSX_CHART_LEGEND_LEFT,

    /** Chart legend positioned at top. */
    LXLSX_CHART_LEGEND_TOP,

    /** Chart legend positioned at bottom. */
    LXLSX_CHART_LEGEND_BOTTOM,

    /** Chart legend positioned at top right. */
    LXLSX_CHART_LEGEND_TOP_RIGHT,

    /** Chart legend overlaid at right side. */
    LXLSX_CHART_LEGEND_OVERLAY_RIGHT,

    /** Chart legend overlaid at left side. */
    LXLSX_CHART_LEGEND_OVERLAY_LEFT,

    /** Chart legend overlaid at top right. */
    LXLSX_CHART_LEGEND_OVERLAY_TOP_RIGHT
} lxlsx_chart_legend_position;

/**
 * @brief Chart line dash types.
 *
 * The dash types are shown in the order that they appear in the Excel dialog.
 * See @ref lxlsx_chart_lines.
 */
typedef enum lxlsx_chart_line_dash_type {

    /** Solid. */
    LXLSX_CHART_LINE_DASH_SOLID = 0,

    /** Round Dot. */
    LXLSX_CHART_LINE_DASH_ROUND_DOT,

    /** Square Dot. */
    LXLSX_CHART_LINE_DASH_SQUARE_DOT,

    /** Dash. */
    LXLSX_CHART_LINE_DASH_DASH,

    /** Dash Dot. */
    LXLSX_CHART_LINE_DASH_DASH_DOT,

    /** Long Dash. */
    LXLSX_CHART_LINE_DASH_LONG_DASH,

    /** Long Dash Dot. */
    LXLSX_CHART_LINE_DASH_LONG_DASH_DOT,

    /** Long Dash Dot Dot. */
    LXLSX_CHART_LINE_DASH_LONG_DASH_DOT_DOT,

    /* These aren't available in the dialog but are used by Excel. */
    LXLSX_CHART_LINE_DASH_DOT,
    LXLSX_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    LXLSX_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
} lxlsx_chart_line_dash_type;

/**
 * @brief Chart marker types.
 */
typedef enum lxlsx_chart_marker_type {

    /** Automatic, series default, marker type. */
    LXLSX_CHART_MARKER_AUTOMATIC,

    /** No marker type. */
    LXLSX_CHART_MARKER_NONE,

    /** Square marker type. */
    LXLSX_CHART_MARKER_SQUARE,

    /** Diamond marker type. */
    LXLSX_CHART_MARKER_DIAMOND,

    /** Triangle marker type. */
    LXLSX_CHART_MARKER_TRIANGLE,

    /** X shape marker type. */
    LXLSX_CHART_MARKER_X,

    /** Star marker type. */
    LXLSX_CHART_MARKER_STAR,

    /** Short dash marker type. */
    LXLSX_CHART_MARKER_SHORT_DASH,

    /** Long dash marker type. */
    LXLSX_CHART_MARKER_LONG_DASH,

    /** Circle marker type. */
    LXLSX_CHART_MARKER_CIRCLE,

    /** Plus (+) marker type. */
    LXLSX_CHART_MARKER_PLUS
} lxlsx_chart_marker_type;

/**
 * @brief Chart pattern types.
 */
typedef enum lxlsx_chart_pattern_type {

    /** None pattern. */
    LXLSX_CHART_PATTERN_NONE,

    /** 5 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_5,

    /** 10 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_10,

    /** 20 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_20,

    /** 25 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_25,

    /** 30 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_30,

    /** 40 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_40,

    /** 50 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_50,

    /** 60 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_60,

    /** 70 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_70,

    /** 75 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_75,

    /** 80 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_80,

    /** 90 Percent pattern. */
    LXLSX_CHART_PATTERN_PERCENT_90,

    /** Light downward diagonal pattern. */
    LXLSX_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,

    /** Light upward diagonal pattern. */
    LXLSX_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,

    /** Dark downward diagonal pattern. */
    LXLSX_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,

    /** Dark upward diagonal pattern. */
    LXLSX_CHART_PATTERN_DARK_UPWARD_DIAGONAL,

    /** Wide downward diagonal pattern. */
    LXLSX_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,

    /** Wide upward diagonal pattern. */
    LXLSX_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,

    /** Light vertical pattern. */
    LXLSX_CHART_PATTERN_LIGHT_VERTICAL,

    /** Light horizontal pattern. */
    LXLSX_CHART_PATTERN_LIGHT_HORIZONTAL,

    /** Narrow vertical pattern. */
    LXLSX_CHART_PATTERN_NARROW_VERTICAL,

    /** Narrow horizontal pattern. */
    LXLSX_CHART_PATTERN_NARROW_HORIZONTAL,

    /** Dark vertical pattern. */
    LXLSX_CHART_PATTERN_DARK_VERTICAL,

    /** Dark horizontal pattern. */
    LXLSX_CHART_PATTERN_DARK_HORIZONTAL,

    /** Dashed downward diagonal pattern. */
    LXLSX_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,

    /** Dashed upward diagonal pattern. */
    LXLSX_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,

    /** Dashed horizontal pattern. */
    LXLSX_CHART_PATTERN_DASHED_HORIZONTAL,

    /** Dashed vertical pattern. */
    LXLSX_CHART_PATTERN_DASHED_VERTICAL,

    /** Small confetti pattern. */
    LXLSX_CHART_PATTERN_SMALL_CONFETTI,

    /** Large confetti pattern. */
    LXLSX_CHART_PATTERN_LARGE_CONFETTI,

    /** Zigzag pattern. */
    LXLSX_CHART_PATTERN_ZIGZAG,

    /** Wave pattern. */
    LXLSX_CHART_PATTERN_WAVE,

    /** Diagonal brick pattern. */
    LXLSX_CHART_PATTERN_DIAGONAL_BRICK,

    /** Horizontal brick pattern. */
    LXLSX_CHART_PATTERN_HORIZONTAL_BRICK,

    /** Weave pattern. */
    LXLSX_CHART_PATTERN_WEAVE,

    /** Plaid pattern. */
    LXLSX_CHART_PATTERN_PLAID,

    /** Divot pattern. */
    LXLSX_CHART_PATTERN_DIVOT,

    /** Dotted grid pattern. */
    LXLSX_CHART_PATTERN_DOTTED_GRID,

    /** Dotted diamond pattern. */
    LXLSX_CHART_PATTERN_DOTTED_DIAMOND,

    /** Shingle pattern. */
    LXLSX_CHART_PATTERN_SHINGLE,

    /** Trellis pattern. */
    LXLSX_CHART_PATTERN_TRELLIS,

    /** Sphere pattern. */
    LXLSX_CHART_PATTERN_SPHERE,

    /** Small grid pattern. */
    LXLSX_CHART_PATTERN_SMALL_GRID,

    /** Large grid pattern. */
    LXLSX_CHART_PATTERN_LARGE_GRID,

    /** Small check pattern. */
    LXLSX_CHART_PATTERN_SMALL_CHECK,

    /** Large check pattern. */
    LXLSX_CHART_PATTERN_LARGE_CHECK,

    /** Outlined diamond pattern. */
    LXLSX_CHART_PATTERN_OUTLINED_DIAMOND,

    /** Solid diamond pattern. */
    LXLSX_CHART_PATTERN_SOLID_DIAMOND
} lxlsx_chart_pattern_type;

/**
 * @brief Chart data label positions.
 */
typedef enum lxlsx_chart_label_position {
    /** Series data label position: default position. */
    LXLSX_CHART_LABEL_POSITION_DEFAULT,

    /** Series data label position: center. */
    LXLSX_CHART_LABEL_POSITION_CENTER,

    /** Series data label position: right. */
    LXLSX_CHART_LABEL_POSITION_RIGHT,

    /** Series data label position: left. */
    LXLSX_CHART_LABEL_POSITION_LEFT,

    /** Series data label position: above. */
    LXLSX_CHART_LABEL_POSITION_ABOVE,

    /** Series data label position: below. */
    LXLSX_CHART_LABEL_POSITION_BELOW,

    /** Series data label position: inside base.  */
    LXLSX_CHART_LABEL_POSITION_INSIDE_BASE,

    /** Series data label position: inside end. */
    LXLSX_CHART_LABEL_POSITION_INSIDE_END,

    /** Series data label position: outside end. */
    LXLSX_CHART_LABEL_POSITION_OUTSIDE_END,

    /** Series data label position: best fit. */
    LXLSX_CHART_LABEL_POSITION_BEST_FIT
} lxlsx_chart_label_position;

/**
 * @brief Chart data label separator.
 */
typedef enum lxlsx_chart_label_separator {
    /** Series data label separator: comma (the default). */
    LXLSX_CHART_LABEL_SEPARATOR_COMMA,

    /** Series data label separator: semicolon. */
    LXLSX_CHART_LABEL_SEPARATOR_SEMICOLON,

    /** Series data label separator: period. */
    LXLSX_CHART_LABEL_SEPARATOR_PERIOD,

    /** Series data label separator: newline. */
    LXLSX_CHART_LABEL_SEPARATOR_NEWLINE,

    /** Series data label separator: space. */
    LXLSX_CHART_LABEL_SEPARATOR_SPACE
} lxlsx_chart_label_separator;

/**
 * @brief Chart axis types.
 */
typedef enum lxlsx_chart_axis_type {
    /** Chart X axis. */
    LXLSX_CHART_AXIS_TYPE_X,

    /** Chart Y axis. */
    LXLSX_CHART_AXIS_TYPE_Y
} lxlsx_chart_axis_type;

enum lxlsx_chart_subtype {

    LXLSX_CHART_SUBTYPE_NONE = 0,
    LXLSX_CHART_SUBTYPE_STACKED,
    LXLSX_CHART_SUBTYPE_STACKED_PERCENT
};

enum lxlsx_chart_grouping {
    LXLSX_GROUPING_CLUSTERED,
    LXLSX_GROUPING_STANDARD,
    LXLSX_GROUPING_PERCENTSTACKED,
    LXLSX_GROUPING_STACKED
};

/**
 * @brief Axis positions for category axes.
 */
typedef enum lxlsx_chart_axis_tick_position {

    LXLSX_CHART_AXIS_POSITION_DEFAULT,

    /** Position category axis on tick marks. */
    LXLSX_CHART_AXIS_POSITION_ON_TICK,

    /** Position category axis between tick marks. */
    LXLSX_CHART_AXIS_POSITION_BETWEEN
} lxlsx_chart_axis_tick_position;

/**
 * @brief Axis label positions.
 */
typedef enum lxlsx_chart_axis_label_position {

    /** Position the axis labels next to the axis. The default. */
    LXLSX_CHART_AXIS_LABEL_POSITION_NEXT_TO,

    /** Position the axis labels at the top of the chart, for horizontal
     * axes, or to the right for vertical axes.*/
    LXLSX_CHART_AXIS_LABEL_POSITION_HIGH,

    /** Position the axis labels at the bottom of the chart, for horizontal
     * axes, or to the left for vertical axes.*/
    LXLSX_CHART_AXIS_LABEL_POSITION_LOW,

    /** Turn off the the axis labels. */
    LXLSX_CHART_AXIS_LABEL_POSITION_NONE
} lxlsx_chart_axis_label_position;

/**
 * @brief Axis label alignments.
 */
typedef enum lxlsx_chart_axis_label_alignment {
    /** Chart axis label alignment: center. */
    LXLSX_CHART_AXIS_LABEL_ALIGN_CENTER,

    /** Chart axis label alignment: left. */
    LXLSX_CHART_AXIS_LABEL_ALIGN_LEFT,

    /** Chart axis label alignment: right. */
    LXLSX_CHART_AXIS_LABEL_ALIGN_RIGHT
} lxlsx_chart_axis_label_alignment;

/**
 * @brief Display units for chart value axis.
 */
typedef enum lxlsx_chart_axis_display_unit {

    /** Axis display units: None. The default. */
    LXLSX_CHART_AXIS_UNITS_NONE,

    /** Axis display units: Hundreds. */
    LXLSX_CHART_AXIS_UNITS_HUNDREDS,

    /** Axis display units: Thousands. */
    LXLSX_CHART_AXIS_UNITS_THOUSANDS,

    /** Axis display units: Ten thousands. */
    LXLSX_CHART_AXIS_UNITS_TEN_THOUSANDS,

    /** Axis display units: Hundred thousands. */
    LXLSX_CHART_AXIS_UNITS_HUNDRED_THOUSANDS,

    /** Axis display units: Millions. */
    LXLSX_CHART_AXIS_UNITS_MILLIONS,

    /** Axis display units: Ten millions. */
    LXLSX_CHART_AXIS_UNITS_TEN_MILLIONS,

    /** Axis display units: Hundred millions. */
    LXLSX_CHART_AXIS_UNITS_HUNDRED_MILLIONS,

    /** Axis display units: Billions. */
    LXLSX_CHART_AXIS_UNITS_BILLIONS,

    /** Axis display units: Trillions. */
    LXLSX_CHART_AXIS_UNITS_TRILLIONS
} lxlsx_chart_axis_display_unit;

/**
 * @brief Tick mark types for an axis.
 */
typedef enum lxlsx_chart_axis_tick_mark {

    /** Default tick mark for the chart axis. Usually outside. */
    LXLSX_CHART_AXIS_TICK_MARK_DEFAULT,

    /** No tick mark for the axis. */
    LXLSX_CHART_AXIS_TICK_MARK_NONE,

    /** Tick mark inside the axis only. */
    LXLSX_CHART_AXIS_TICK_MARK_INSIDE,

    /** Tick mark outside the axis only. */
    LXLSX_CHART_AXIS_TICK_MARK_OUTSIDE,

    /** Tick mark inside and outside the axis. */
    LXLSX_CHART_AXIS_TICK_MARK_CROSSING
} lxlsx_chart_tick_mark;

typedef struct lxlsx_series_range {
    char *formula;
    char *sheetname;
    lxlsx_row_t first_row;
    lxlsx_row_t last_row;
    lxlsx_col_t first_col;
    lxlsx_col_t last_col;
    uint8_t ignore_cache;

    uint8_t has_string_cache;
    uint16_t num_data_points;
    struct lxlsx_series_data_points *data_cache;

} lxlsx_series_range;

typedef struct lxlsx_series_data_point {
    uint8_t is_string;
    double number;
    char *string;
    uint8_t no_data;

    STAILQ_ENTRY (lxlsx_series_data_point) list_pointers;

} lxlsx_series_data_point;

/**
 * @brief Struct to represent a chart line.
 *
 * See @ref lxlsx_chart_lines.
 */
typedef struct lxlsx_chart_line {

    /** The chart font color. See @ref working_with_colors. */
    lxlsx_color_t color;

    /** Turn off/hide line. Set to 0 or 1.*/
    uint8_t none;

    /** Width of the line in increments of 0.25. Default is 2.25. */
    float width;

    /** The line dash type. See #lxlsx_chart_line_dash_type. */
    uint8_t dash_type;

    /** Set the transparency of the line. 0 - 100. Default 0. */
    uint8_t transparency;

} lxlsx_chart_line;

/**
 * @brief Struct to represent a chart fill.
 *
 * See @ref lxlsx_chart_fills.
 */
typedef struct lxlsx_chart_fill {

    /** The chart font color. See @ref working_with_colors. */
    lxlsx_color_t color;

    /** Turn off/hide line. Set to 0 or 1.*/
    uint8_t none;

    /** Set the transparency of the fill. 0 - 100. Default 0. */
    uint8_t transparency;

} lxlsx_chart_fill;

/**
 * @brief Struct to represent a chart pattern.
 *
 * See @ref lxlsx_chart_patterns.
 */
typedef struct lxlsx_chart_pattern {

    /** The pattern foreground color. See @ref working_with_colors. */
    lxlsx_color_t fg_color;

    /** The pattern background color. See @ref working_with_colors. */
    lxlsx_color_t bg_color;

    /** The pattern type. See #lxlsx_chart_pattern_type. */
    uint8_t type;

} lxlsx_chart_pattern;

/**
 * @brief Struct to represent a chart font.
 *
 * See @ref lxlsx_chart_fonts.
 */
typedef struct lxlsx_chart_font {

    /** The chart font name, such as "Arial" or "Calibri". */
    const char *name;

    /** The chart font size. The default is 11. */
    double size;

    /** The chart font bold property. Set to 0 or 1. */
    uint8_t bold;

    /** The chart font italic property. Set to 0 or 1. */
    uint8_t italic;

    /** The chart font underline property. Set to 0 or 1. */
    uint8_t underline;

    /** The chart font rotation property. Range: -90 to 90, and 270, 271 and 360:
     *
     *  - The angles -90 to 90 are the normal range shown in the Excel user interface.
     *  - The angle 270 gives a stacked (top to bottom) alignment.
     *  - The angle 271 gives a stacked alignment for East Asian fonts.
     *  - The angle 360 gives an explicit angle of 0 to override the y axis default.
     * */
    int32_t rotation;

    /** The chart font color. See @ref working_with_colors. */
    lxlsx_color_t color;

    /** The chart font pitch family property. Rarely required, set to 0. */
    uint8_t pitch_family;

    /** The chart font character set property. Rarely required, set to 0. */
    uint8_t charset;

    /** The chart font baseline property. Rarely required, set to 0. */
    int8_t baseline;

} lxlsx_chart_font;

/**
 * @brief Struct to represent Excel chart element layout dimensions.
 *
 * Excel supports manual positioning of elements such as the chart axis labels,
 * the chart legend, the chart plot area and the chart title. The
 * `lxlsx_chart_layout` struct represents the layout dimension for these elements.
 *
 * The layout units used by Excel are relative units expressed as a percentage
 * of the chart dimensions and are double values in the range `0.0 < x <= 1.0`.
 * Excel calculates these dimensions as shown below:
 *
 * @image html lxlsx_chart_layout.png
 *
 * With reference to the above figure the layout units are calculated as
 * follows:
 *
 * ```text
 *     x = a / W
 *     y = b / H
 * ```
 *
 * These units are cumbersome and can vary depending on other elements in the
 * chart such as text lengths. However, these are the units that are required by
 * Excel to allow relative positioning. Some trial and error is generally
 * required.
 *
 * For the chart `lxlsx_chart_plotarea_set_layout()` and `lxlsx_chart_legend_set_layout()`
 * functions you can also set the width and height based on the following
 * calculation:
 *
 * ```text
 *     width  = w / W
 *     height = h / H
 * ```
 *
 * For other text based objects the width and height are changed by the font
 * dimensions.
 *
 * The chart functions that support `lxlsx_chart_layout` are:
 *
 * - `lxlsx_chart_title_set_layout()`
 * - `lxlsx_chart_legend_set_layout()`
 * - `lxlsx_chart_plotarea_set_layout()`
 * - `lxlsx_chart_axis_set_name_layout()`
 *
 */
typedef struct lxlsx_chart_layout {

    /** The x offset in the range `0.0 < x <= 1.0` */
    double x;

    /** The y offset in the range `0.0 < y <= 1.0` */
    double y;

    /** The width of the plotarea or legend in the range `0.0 < x <= 1.0` */
    double width;

    /** The height of the plotarea or legend in the range `0.0 < x <= 1.0` */
    double height;

    uint8_t has_inner;

} lxlsx_chart_layout;

typedef struct lxlsx_chart_marker {

    uint8_t type;
    uint8_t size;
    lxlsx_chart_line *line;
    lxlsx_chart_fill *fill;
    lxlsx_chart_pattern *pattern;

} lxlsx_chart_marker;

typedef struct lxlsx_chart_legend {

    lxlsx_chart_font *font;
    uint8_t position;
    lxlsx_chart_layout *layout;

} lxlsx_chart_legend;

typedef struct lxlsx_chart_title {

    char *name;
    lxlsx_row_t row;
    lxlsx_col_t col;
    lxlsx_chart_font *font;
    uint8_t off;
    uint8_t is_horizontal;
    uint8_t ignore_cache;
    uint8_t has_overlay;

    /* We use a range to hold the title formula properties even though it
     * will only have 1 point in order to re-use similar functions.*/
    lxlsx_series_range *range;

    struct lxlsx_series_data_point data_point;
    lxlsx_chart_layout *layout;

} lxlsx_chart_title;

/**
 * @brief Struct to represent an Excel chart data point.
 *
 * The lxlsx_chart_point used to set the line, fill and pattern of one or more
 * points in a chart data series. See @ref lxlsx_chart_points.
 */
typedef struct lxlsx_chart_point {

    /** The line/border for the chart point. See @ref lxlsx_chart_lines. */
    lxlsx_chart_line *line;

    /** The fill for the chart point. See @ref lxlsx_chart_fills. */
    lxlsx_chart_fill *fill;

    /** The pattern for the chart point. See @ref lxlsx_chart_patterns.*/
    lxlsx_chart_pattern *pattern;

} lxlsx_chart_point;

/**
 * @brief Struct to represent an Excel chart data label.
 *
 * The lxlsx_chart_data_label struct is used to represent a data label in a
 * chart series so that custom properties can be set for it.
 */
typedef struct lxlsx_chart_data_label {

    /** The string or formula value for the data label. See
     *  @ref lxlsx_chart_custom_labels. */
    const char *value;

    /** Option to hide/delete the data label from the chart series.
     *  See @ref lxlsx_chart_custom_labels. */
    uint8_t hide;

    /** The font properties for the chart data label. @ref lxlsx_chart_fonts. */
    lxlsx_chart_font *font;

    /** The line/border for the chart data label. See @ref lxlsx_chart_lines. */
    lxlsx_chart_line *line;

    /** The fill for the chart data label. See @ref lxlsx_chart_fills. */
    lxlsx_chart_fill *fill;

    /** The pattern for the chart data label. See @ref lxlsx_chart_patterns.*/
    lxlsx_chart_pattern *pattern;

} lxlsx_chart_data_label;

/* Internal version of lxlsx_chart_data_label with more metadata. */
typedef struct lxlsx_chart_custom_label {

    char *value;
    uint8_t hide;
    lxlsx_chart_font *font;
    lxlsx_chart_line *line;
    lxlsx_chart_fill *fill;
    lxlsx_chart_pattern *pattern;

    /* We use a range to hold the label formula properties even though it
     * will only have 1 point in order to re-use similar functions.*/
    lxlsx_series_range *range;

    struct lxlsx_series_data_point data_point;

} lxlsx_chart_custom_label;

/**
 * @brief Define how blank values are displayed in a chart.
 */
typedef enum lxlsx_chart_blank {

    /** Show empty chart cells as gaps in the data. The default. */
    LXLSX_CHART_BLANKS_AS_GAP,

    /** Show empty chart cells as zeros. */
    LXLSX_CHART_BLANKS_AS_ZERO,

    /** Show empty chart cells as connected. Only for charts with lines. */
    LXLSX_CHART_BLANKS_AS_CONNECTED
} lxlsx_chart_blank;

enum lxlsx_chart_position {
    LXLSX_CHART_AXIS_RIGHT,
    LXLSX_CHART_AXIS_LEFT,
    LXLSX_CHART_AXIS_TOP,
    LXLSX_CHART_AXIS_BOTTOM
};

enum lxlsx_chart_layout_type {
    LXLSX_CHART_LAYOUT_TITLE,
    LXLSX_CHART_LAYOUT_LEGEND,
    LXLSX_CHART_LAYOUT_PLOTAREA,
    LXLSX_CHART_LAYOUT_AXIS_NAME
};

/**
 * @brief Type/amount of data series error bar.
 */
typedef enum lxlsx_chart_error_bar_type {
    /** Error bar type: Standard error. */
    LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR,

    /** Error bar type: Fixed value. */
    LXLSX_CHART_ERROR_BAR_TYPE_FIXED,

    /** Error bar type: Percentage. */
    LXLSX_CHART_ERROR_BAR_TYPE_PERCENTAGE,

    /** Error bar type: Standard deviation(s). */
    LXLSX_CHART_ERROR_BAR_TYPE_STD_DEV
} lxlsx_chart_error_bar_type;

/**
 * @brief Direction for a data series error bar.
 */
typedef enum lxlsx_chart_error_bar_direction {

    /** Error bar extends in both directions. The default. */
    LXLSX_CHART_ERROR_BAR_DIR_BOTH,

    /** Error bar extends in positive direction. */
    LXLSX_CHART_ERROR_BAR_DIR_PLUS,

    /** Error bar extends in negative direction. */
    LXLSX_CHART_ERROR_BAR_DIR_MINUS
} lxlsx_chart_error_bar_direction;

/**
 * @brief Direction for a data series error bar.
 */
typedef enum lxlsx_chart_error_bar_axis {
    /** X axis error bar. */
    LXLSX_CHART_ERROR_BAR_AXIS_X,

    /** Y axis error bar. */
    LXLSX_CHART_ERROR_BAR_AXIS_Y
} lxlsx_chart_error_bar_axis;

/**
 * @brief End cap styles for a data series error bar.
 */
typedef enum lxlsx_chart_error_bar_cap {
    /** Flat end cap. The default. */
    LXLSX_CHART_ERROR_BAR_END_CAP,

    /** No end cap. */
    LXLSX_CHART_ERROR_BAR_NO_CAP
} lxlsx_chart_error_bar_cap;

typedef struct lxlsx_series_error_bars {
    uint8_t type;
    uint8_t direction;
    uint8_t endcap;
    uint8_t has_value;
    uint8_t is_set;
    uint8_t is_x;
    uint8_t lxlsx_chart_group;
    double value;
    lxlsx_chart_line *line;

} lxlsx_series_error_bars;

/**
 * @brief Series trendline/regression types.
 */
typedef enum lxlsx_chart_trendline_type {
    /** Trendline type: Linear. */
    LXLSX_CHART_TRENDLINE_TYPE_LINEAR,

    /** Trendline type: Logarithm. */
    LXLSX_CHART_TRENDLINE_TYPE_LOG,

    /** Trendline type: Polynomial. */
    LXLSX_CHART_TRENDLINE_TYPE_POLY,

    /** Trendline type: Power. */
    LXLSX_CHART_TRENDLINE_TYPE_POWER,

    /** Trendline type: Exponential. */
    LXLSX_CHART_TRENDLINE_TYPE_EXP,

    /** Trendline type: Moving Average. */
    LXLSX_CHART_TRENDLINE_TYPE_AVERAGE
} lxlsx_chart_trendline_type;

/**
 * @brief Struct to represent an Excel chart data series.
 *
 * The lxlsx_chart_series is created using the lxlsx_chart_add_series function. It is
 * used in functions that modify a chart series but the members of the struct
 * aren't modified directly.
 */
typedef struct lxlsx_chart_series {

    lxlsx_series_range *categories;
    lxlsx_series_range *values;
    lxlsx_chart_title title;
    lxlsx_chart_line *line;
    lxlsx_chart_fill *fill;
    lxlsx_chart_pattern *pattern;
    lxlsx_chart_marker *marker;
    lxlsx_chart_point *points;
    lxlsx_chart_custom_label *data_labels;
    uint16_t point_count;
    uint16_t data_label_count;

    uint8_t smooth;
    uint8_t invert_if_negative;

    /* Data label parameters. */
    uint8_t has_labels;
    uint8_t show_labels_value;
    uint8_t show_labels_category;
    uint8_t show_labels_name;
    uint8_t show_labels_leader;
    uint8_t show_labels_legend;
    uint8_t show_labels_percent;
    uint8_t label_position;
    uint8_t label_separator;
    uint8_t default_label_position;
    char *label_num_format;
    lxlsx_chart_font *label_font;
    lxlsx_chart_line *label_line;
    lxlsx_chart_fill *label_fill;
    lxlsx_chart_pattern *label_pattern;

    lxlsx_series_error_bars *x_error_bars;
    lxlsx_series_error_bars *y_error_bars;

    uint8_t has_trendline;
    uint8_t has_trendline_forecast;
    uint8_t has_trendline_equation;
    uint8_t has_trendline_r_squared;
    uint8_t has_trendline_intercept;
    uint8_t trendline_type;
    uint8_t trendline_value;
    double trendline_forward;
    double trendline_backward;
    uint8_t trendline_value_type;
    char *trendline_name;
    lxlsx_chart_line *trendline_line;
    double trendline_intercept;

    STAILQ_ENTRY (lxlsx_chart_series) list_pointers;

} lxlsx_chart_series;

/* Struct for major/minor axis gridlines. */
typedef struct lxlsx_chart_gridline {

    uint8_t visible;
    lxlsx_chart_line *line;

} lxlsx_chart_gridline;

/**
 * @brief Struct to represent an Excel chart axis.
 *
 * The lxlsx_chart_axis struct is used in functions that modify a chart axis
 * but the members of the struct aren't modified directly.
 */
typedef struct lxlsx_chart_axis {

    lxlsx_chart_title title;

    char *num_format;
    char *default_num_format;
    uint8_t source_linked;

    uint8_t major_tick_mark;
    uint8_t minor_tick_mark;
    uint8_t is_horizontal;

    lxlsx_chart_gridline major_gridlines;
    lxlsx_chart_gridline minor_gridlines;

    lxlsx_chart_font *num_font;
    lxlsx_chart_line *line;
    lxlsx_chart_fill *fill;
    lxlsx_chart_pattern *pattern;

    uint8_t is_category;
    uint8_t is_date;
    uint8_t is_value;
    uint8_t axis_position;
    uint8_t position_axis;
    uint8_t label_position;
    uint8_t label_align;
    uint8_t hidden;
    uint8_t reverse;

    uint8_t has_min;
    double min;
    uint8_t has_max;
    double max;

    uint8_t has_major_unit;
    double major_unit;
    uint8_t has_minor_unit;
    double minor_unit;

    uint16_t interval_unit;
    uint16_t interval_tick;

    uint16_t log_base;

    uint8_t display_units;
    uint8_t display_units_visible;

    uint8_t has_crossing;
    uint8_t crossing_min;
    uint8_t crossing_max;
    double crossing;

} lxlsx_chart_axis;

/**
 * @brief Struct to represent an Excel chart.
 *
 * The members of the lxlsx_chart struct aren't modified directly. Instead
 * the chart properties are set by calling the functions shown in chart.h.
 */
typedef struct lxlsx_chart {

    FILE *file;

    uint8_t type;
    uint8_t subtype;
    uint16_t series_index;

    void (*write_chart_type)(struct lxlsx_chart *);
    void (*write_plot_area)(struct lxlsx_chart *);

    /**
     * A pointer to the chart x_axis object which can be used in functions
     * that configures the X axis.
     */
    lxlsx_chart_axis *x_axis;

    /**
     * A pointer to the chart y_axis object which can be used in functions
     * that configures the Y axis.
     */
    lxlsx_chart_axis *y_axis;

    lxlsx_chart_title title;

    uint32_t id;
    uint32_t axis_id_1;
    uint32_t axis_id_2;
    uint32_t axis_id_3;
    uint32_t axis_id_4;

    uint8_t in_use;
    uint8_t lxlsx_chart_group;
    uint8_t cat_has_num_fmt;
    uint8_t is_chartsheet;

    uint8_t has_horiz_cat_axis;
    uint8_t has_horiz_val_axis;

    uint8_t style_id;
    uint16_t rotation;
    uint16_t hole_size;

    uint8_t no_title;
    uint8_t has_overlap;
    int8_t overlap_y1;
    int8_t overlap_y2;
    uint16_t gap_y1;
    uint16_t gap_y2;

    uint8_t grouping;
    uint8_t default_cross_between;

    lxlsx_chart_legend legend;
    int16_t *delete_series;
    uint16_t delete_series_count;
    lxlsx_chart_marker *default_marker;

    lxlsx_chart_line *chartarea_line;
    lxlsx_chart_fill *chartarea_fill;
    lxlsx_chart_pattern *chartarea_pattern;

    lxlsx_chart_line *plotarea_line;
    lxlsx_chart_fill *plotarea_fill;
    lxlsx_chart_layout *plotarea_layout;
    lxlsx_chart_pattern *plotarea_pattern;

    uint8_t has_drop_lines;
    lxlsx_chart_line *drop_lines_line;

    uint8_t has_high_low_lines;
    lxlsx_chart_line *high_low_lines_line;

    struct lxlsx_chart_series_list *series_list;

    uint8_t has_table;
    uint8_t has_table_vertical;
    uint8_t has_table_horizontal;
    uint8_t has_table_outline;
    uint8_t has_table_legend_keys;
    lxlsx_chart_font *lxlsx_table_font;

    uint8_t show_blanks_as;
    uint8_t show_hidden_data;

    uint8_t has_up_down_bars;
    lxlsx_chart_line *up_bar_line;
    lxlsx_chart_line *down_bar_line;
    lxlsx_chart_fill *up_bar_fill;
    lxlsx_chart_fill *down_bar_fill;

    uint8_t default_label_position;
    uint8_t is_protected;

    STAILQ_ENTRY (lxlsx_chart) ordered_list_pointers;
    STAILQ_ENTRY (lxlsx_chart) list_pointers;

} lxlsx_chart;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_chart *lxlsx_chart_new(uint8_t type);
void lxlsx_chart_free(lxlsx_chart *chart);
void lxlsx_chart_assemble_xml_file(lxlsx_chart *chart);

/**
 * @brief Add a data series to a chart.
 *
 * @param chart      Pointer to a lxlsx_chart instance to be configured.
 * @param categories The range of categories in the data series.
 * @param values     The range of values in the data series.
 *
 * @return A lxlsx_chart_series object pointer.
 *
 * In Excel a chart **series** is a collection of information that defines
 * which data is plotted such as the categories and values. It is also used to
 * define the formatting for the data.
 *
 * For an libxlsxwriter chart object the `%lxlsx_chart_add_series()` function is
 * used to set the categories and values of the series:
 *
 * @code
 *     lxlsx_chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
 * @endcode
 *
 *
 * The series parameters are:
 *
 * - `categories`: This sets the chart category labels. The category is more
 *   or less the same as the X axis. In most Excel chart types the
 *   `categories` property is optional and the chart will just assume a
 *   sequential series from `1..n`:
 *
 * @code
 *     // The NULL category will default to 1 to 5 like in Excel.
 *     lxlsx_chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
 * @endcode
 *
 *  - `values`: This is the most important property of a series and is the
 *    only mandatory option for every chart object. This parameter links the
 *    chart with the worksheet data that it displays.
 *
 * The `categories` and `values` should be a string formula like
 * `"=Sheet1!$A$2:$A$7"` in the same way it is represented in Excel. This is
 * convenient when recreating a chart from an example in Excel but it is
 * trickier to generate programmatically. For these cases you can set the
 * `categories` and `values` to `NULL` and use the
 * `lxlsx_chart_series_set_categories()` and `lxlsx_chart_series_set_values()` functions:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart, NULL, NULL);
 *
 *     // Configure the series using a syntax that is easier to define programmatically.
 *     lxlsx_chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
 *     lxlsx_chart_series_set_values(    series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
 * @endcode
 *
 * As shown in the previous example the return value from
 * `%lxlsx_chart_add_series()` is a lxlsx_chart_series pointer. This can be used in
 * other functions that configure a series.
 *
 *
 * More than one series can be added to a chart. The series numbering and
 * order in the Excel chart will be the same as the order in which they are
 * added in libxlsxwriter:
 *
 * @code
 *    lxlsx_chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
 *    lxlsx_chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5");
 *    lxlsx_chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5");
 * @endcode
 *
 * It is also possible to specify non-contiguous ranges:
 *
 * @code
 *    lxlsx_chart_add_series(
 *        chart,
 *        "=(Sheet1!$A$1:$A$9,Sheet1!$A$14:$A$25)",
 *        "=(Sheet1!$B$1:$B$9,Sheet1!$B$14:$B$25)"
 *    );
 * @endcode
 *
 */
lxlsx_chart_series *lxlsx_chart_add_series(lxlsx_chart *chart,
                                   const char *categories,
                                   const char *values);

/**
 * @brief Set a series "categories" range using row and column values.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `lxlsx_chart_add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%lxlsx_chart_series_set_categories()` function is an alternative method that
 * is easier to generate programmatically. It requires that you set the
 * `categories` and `values` parameters in `lxlsx_chart_add_series()`to `NULL` and
 * then set them using row and column values in
 * `lxlsx_chart_series_set_categories()` and `lxlsx_chart_series_set_values()`:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart, NULL, NULL);
 *
 *     // Configure the series ranges programmatically.
 *     lxlsx_chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
 *     lxlsx_chart_series_set_values(    series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
 * @endcode
 *
 */
void lxlsx_chart_series_set_categories(lxlsx_chart_series *series,
                                 const char *sheetname, lxlsx_row_t first_row,
                                 lxlsx_col_t first_col, lxlsx_row_t last_row,
                                 lxlsx_col_t last_col);

/**
 * @brief Set a series "values" range using row and column values.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `lxlsx_chart_add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%lxlsx_chart_series_set_values()` function is an alternative method that is
 * easier to generate programmatically. See the documentation for
 * `lxlsx_chart_series_set_categories()` above.
 */
void lxlsx_chart_series_set_values(lxlsx_chart_series *series, const char *sheetname,
                             lxlsx_row_t first_row, lxlsx_col_t first_col,
                             lxlsx_row_t last_row, lxlsx_col_t last_col);

/**
 * @brief Set the name of a chart series range.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param name   The series name.
 *
 * The `%lxlsx_chart_series_set_name` function is used to set the name for a chart
 * data series. The series name in Excel is displayed in the chart legend and
 * in the formula bar. The name property is optional and if it isn't supplied
 * it will default to `Series 1..n`.
 *
 * The function applies to a #lxlsx_chart_series object created using
 * `lxlsx_chart_add_series()`:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     lxlsx_chart_series_set_name(series, "Quarterly budget data");
 * @endcode
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     lxlsx_chart_series_set_name(series, "=Sheet1!$B$1");
 * @endcode
 *
 * See also the `lxlsx_chart_series_set_name_range()` function to see how to set the
 * name formula programmatically.
 */
void lxlsx_chart_series_set_name(lxlsx_chart_series *series, const char *name);

/**
 * @brief Set a series name formula using row and column values.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%lxlsx_chart_series_set_name_range()` function can be used to set a series
 * name range and is an alternative to using `lxlsx_chart_series_set_name()` and a
 * string formula:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     lxlsx_chart_series_set_name_range(series, "Sheet1", 0, 2); // "=Sheet1!$C$1"
 * @endcode
 */
void lxlsx_chart_series_set_name_range(lxlsx_chart_series *series,
                                 const char *sheetname, lxlsx_row_t row,
                                 lxlsx_col_t col);
/**
 * @brief Set the line properties for a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param line   A #lxlsx_chart_line struct.
 *
 * Set the line/border properties of a chart series:
 *
 * @code
 *     lxlsx_chart_line line = {.color = LXLSX_COLOR_RED};
 *
 *     lxlsx_chart_series_set_line(series1, &line);
 *     lxlsx_chart_series_set_line(series2, &line);
 *     lxlsx_chart_series_set_line(series3, &line);
 * @endcode
 *
 * @image html lxlsx_chart_series_set_line.png
 *
 * For more information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_series_set_line(lxlsx_chart_series *series, lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param fill   A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of a chart series:
 *
 * @code
 *     lxlsx_chart_fill fill1 = {.color = LXLSX_COLOR_RED};
 *     lxlsx_chart_fill fill2 = {.color = LXLSX_COLOR_YELLOW};
 *     lxlsx_chart_fill fill3 = {.color = LXLSX_COLOR_GREEN};
 *
 *     lxlsx_chart_series_set_fill(series1, &fill1);
 *     lxlsx_chart_series_set_fill(series2, &fill2);
 *     lxlsx_chart_series_set_fill(series3, &fill3);
 * @endcode
 *
 * @image html lxlsx_chart_series_set_fill.png
 *
 * For more information see @ref lxlsx_chart_fills.
 */
void lxlsx_chart_series_set_fill(lxlsx_chart_series *series, lxlsx_chart_fill *fill);

/**
 * @brief Invert the fill color for negative series values.
 *
 * @param series  A series object created via `lxlsx_chart_add_series()`.
 *
 * Invert the fill color for negative values. Usually only applicable to
 * column and bar charts.
 *
 * @code
 *     lxlsx_chart_series_set_invert_if_negative(series);
 * @endcode
 *
 */
void lxlsx_chart_series_set_invert_if_negative(lxlsx_chart_series *series);

/**
 * @brief Set the pattern properties for a chart series.
 *
 * @param series  A series object created via `lxlsx_chart_add_series()`.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of a chart series:
 *
 * @code
 *     lxlsx_chart_pattern pattern1 = {.type = LXLSX_CHART_PATTERN_SHINGLE,
 *                                   .fg_color = 0x804000,
 *                                   .bg_color = 0XC68C53};
 *
 *     lxlsx_chart_pattern pattern2 = {.type = LXLSX_CHART_PATTERN_HORIZONTAL_BRICK,
 *                                   .fg_color = 0XB30000,
 *                                   .bg_color = 0XFF6666};
 *
 *     lxlsx_chart_series_set_pattern(series1, &pattern1);
 *     lxlsx_chart_series_set_pattern(series2, &pattern2);
 *
 * @endcode
 *
 * @image html lxlsx_chart_pattern.png
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 */
void lxlsx_chart_series_set_pattern(lxlsx_chart_series *series,
                              lxlsx_chart_pattern *pattern);

/**
 * @brief Set the data marker type for a series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param type   The marker type, see #lxlsx_chart_marker_type.
 *
 * In Excel a chart marker is used to distinguish data points in a plotted
 * series. In general only Line and Scatter and Radar chart types use
 * markers. The libxlsxwriter chart types that can have markers are:
 *
 * - #LXLSX_CHART_LINE
 * - #LXLSX_CHART_SCATTER
 * - #LXLSX_CHART_SCATTER_STRAIGHT
 * - #LXLSX_CHART_SCATTER_STRAIGHT_WITH_MARKERS
 * - #LXLSX_CHART_SCATTER_SMOOTH
 * - #LXLSX_CHART_SCATTER_SMOOTH_WITH_MARKERS
 * - #LXLSX_CHART_RADAR
 * - #LXLSX_CHART_RADAR_WITH_MARKERS
 *
 * The chart types with `MARKERS` in the name have markers with default colors
 * and shapes turned on by default but it is possible using the various
 * `lxlsx_chart_series_set_marker_xxx()` functions below to change these defaults. It
 * is also possible to turn on an off markers.
 *
 * The `%lxlsx_chart_series_set_marker_type()` function is used to specify the
 * type of the series marker:
 *
 * @code
 *     lxlsx_chart_series_set_marker_type(series, LXLSX_CHART_MARKER_DIAMOND);
 * @endcode
 *
 * @image html lxlsx_chart_marker1.png
 *
 * The available marker types defined by #lxlsx_chart_marker_type are:
 *
 * - #LXLSX_CHART_MARKER_AUTOMATIC
 * - #LXLSX_CHART_MARKER_NONE
 * - #LXLSX_CHART_MARKER_SQUARE
 * - #LXLSX_CHART_MARKER_DIAMOND
 * - #LXLSX_CHART_MARKER_TRIANGLE
 * - #LXLSX_CHART_MARKER_X
 * - #LXLSX_CHART_MARKER_STAR
 * - #LXLSX_CHART_MARKER_SHORT_DASH
 * - #LXLSX_CHART_MARKER_LONG_DASH
 * - #LXLSX_CHART_MARKER_CIRCLE
 * - #LXLSX_CHART_MARKER_PLUS
 *
 * The `#LXLSX_CHART_MARKER_NONE` type can be used to turn off default markers:
 *
 * @code
 *     lxlsx_chart_series_set_marker_type(series, LXLSX_CHART_MARKER_NONE);
 * @endcode
 *
 * @image html lxlsx_chart_series_set_marker_none.png
 *
 * The `#LXLSX_CHART_MARKER_AUTOMATIC` type is a special case which turns on a
 * marker using the default marker style for the particular series. If
 * automatic is on then other marker properties such as size, line or fill
 * cannot be set.
 */
void lxlsx_chart_series_set_marker_type(lxlsx_chart_series *series, uint8_t type);

/**
 * @brief Set the size of a data marker for a series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param size   The size of the marker.
 *
 * The `%lxlsx_chart_series_set_marker_size()` function is used to specify the
 * size of the series marker:
 *
 * @code
 *     lxlsx_chart_series_set_marker_type(series, LXLSX_CHART_MARKER_CIRCLE);
 *     lxlsx_chart_series_set_marker_size(series, 10);
 * @endcode
 *
 * @image html lxlsx_chart_series_set_marker_size.png
 *
 */
void lxlsx_chart_series_set_marker_size(lxlsx_chart_series *series, uint8_t size);

/**
 * @brief Set the line properties for a chart series marker.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param line   A #lxlsx_chart_line struct.
 *
 * Set the line/border properties of a chart marker:
 *
 * @code
 *     lxlsx_chart_line line = {.color = LXLSX_COLOR_BLACK};
 *     lxlsx_chart_fill fill = {.color = LXLSX_COLOR_RED};
 *
 *     lxlsx_chart_series_set_marker_type(series, LXLSX_CHART_MARKER_SQUARE);
 *     lxlsx_chart_series_set_marker_size(series, 8);
 *
 *     lxlsx_chart_series_set_marker_line(series, &line);
 *     lxlsx_chart_series_set_marker_fill(series, &fill);
 * @endcode
 *
 * @image html lxlsx_chart_marker2.png
 *
 * For more information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_series_set_marker_line(lxlsx_chart_series *series,
                                  lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for a chart series marker.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param fill   A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of a chart marker:
 *
 * @code
 *     lxlsx_chart_series_set_marker_fill(series, &fill);
 * @endcode
 *
 * See the example and image above and also see @ref lxlsx_chart_fills.
 */
void lxlsx_chart_series_set_marker_fill(lxlsx_chart_series *series,
                                  lxlsx_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chart series marker.
 *
 * @param series  A series object created via `lxlsx_chart_add_series()`.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of a chart marker:
 *
 * @code
 *     lxlsx_chart_series_set_marker_pattern(series, &pattern);
 * @endcode
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 */
void lxlsx_chart_series_set_marker_pattern(lxlsx_chart_series *series,
                                     lxlsx_chart_pattern *pattern);

/**
 * @brief Set the formatting for points in the series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param points An NULL terminated array of #lxlsx_chart_point pointers.
 *
 * @return A #lxlsx_error.
 *
 * In general formatting is applied to an entire series in a chart. However,
 * it is occasionally required to format individual points in a series. In
 * particular this is required for Pie/Doughnut charts where each segment is
 * represented by a point.
 *
 * @dontinclude lxlsx_chart_pie_colors.c
 * @skip Add the data series
 * @until lxlsx_chart_series_set_points
 *
 * @image html lxlsx_chart_points1.png
 *
 * @note The array of #lxlsx_chart_point pointers should be NULL terminated
 * as shown in the example.
 *
 * For more details see @ref lxlsx_chart_points
 */
lxlsx_error lxlsx_chart_series_set_points(lxlsx_chart_series *series,
                                  lxlsx_chart_point *points[]);

/**
 * @brief Smooth a line or scatter chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param smooth Turn off/on the line smoothing. (0/1)
 *
 * The `lxlsx_chart_series_set_smooth()` function is used to set the smooth property
 * of a line series. It is only applicable to the line and scatter chart
 * types:
 *
 * @code
 *     lxlsx_chart_series_set_smooth(series2, LXLSX_TRUE);
 * @endcode
 *
 * @image html lxlsx_chart_smooth.png
 *
 *
 */
void lxlsx_chart_series_set_smooth(lxlsx_chart_series *series, uint8_t smooth);

/**
 * @brief Add data labels to a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_labels()` function is used to turn on data labels
 * for a chart series. Data labels indicate the values of the plotted data
 * points.
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels1.png
 *
 * By default data labels are displayed in Excel with only the values shown:
 *
 * @image html lxlsx_chart_data_labels2.png
 *
 * However, it is possible to configure other display options, as shown
 * in the functions below.
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels(lxlsx_chart_series *series);

/**
 * @brief Set the display options for the labels of a data series.
 *
 * @param series        A series object created via `lxlsx_chart_add_series()`.
 * @param show_name     Turn on/off the series name in the label caption.
 * @param show_category Turn on/off the category name in the label caption.
 * @param show_value    Turn on/off the value in the label caption.
 *
 * The `%lxlsx_chart_series_set_labels_options()` function is used to set the
 * parameters that are displayed in the series data label:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_options(series, LXLSX_TRUE, LXLSX_TRUE, LXLSX_TRUE);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels3.png
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_options(lxlsx_chart_series *series,
                                     uint8_t show_name, uint8_t show_category,
                                     uint8_t show_value);

/** @brief Set the properties for data labels in a series.
*
* @param series      A series object created via `lxlsx_chart_add_series()`.
* @param data_labels An NULL terminated array of #lxlsx_chart_data_label pointers.
*
* @return A #lxlsx_error.
*
* The `%lxlsx_chart_series_set_labels_custom()` function is used to set the properties
* for data labels in a series. It can also be used to delete individual data
* labels in a series.
*
* In general properties are set for all the data labels in a chart
* series. However, it is also possible to set properties for individual data
* labels in a series using `%lxlsx_chart_series_set_labels_custom()`.
*
* The `%lxlsx_chart_series_set_labels_custom()` function takes a pointer to an array
* of #lxlsx_chart_data_label pointers. The list should be `NULL` terminated:
*
* @code
*     // Add the series data labels.
*     lxlsx_chart_series_set_labels(series);
*
*     // Create some custom labels.
*     lxlsx_chart_data_label data_label1 = {.value = "Jan"};
*     lxlsx_chart_data_label data_label2 = {.value = "Feb"};
*     lxlsx_chart_data_label data_label3 = {.value = "Mar"};
*     lxlsx_chart_data_label data_label4 = {.value = "Apr"};
*     lxlsx_chart_data_label data_label5 = {.value = "May"};
*     lxlsx_chart_data_label data_label6 = {.value = "Jun"};
*
*     // Create an array of label pointers. NULL indicates the end of the array.
*     lxlsx_chart_data_label *data_labels[] = {
*         &data_label1,
*         &data_label2,
*         &data_label3,
*         &data_label4,
*         &data_label5,
*         &data_label6,
*         NULL
*     };
*
*     // Set the custom labels.
*     lxlsx_chart_series_set_labels_custom(series, data_labels);
* @endcode
*
* @image html lxlsx_chart_data_labels18.png
*
* @note The array of #lxlsx_chart_point pointers should be NULL terminated as
* shown in the example. Any #lxlsx_chart_data_label items set to a default
* initialization or omitted from the list will be assigned the default data
* label value.
*
* For more details see @ref lxlsx_chart_custom_labels.
*/
lxlsx_error lxlsx_chart_series_set_labels_custom(lxlsx_chart_series *series, lxlsx_chart_data_label
                                         *data_labels[]);

/**
 * @brief Set the separator for the data label captions.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param separator The separator for the data label options:
 *                  #lxlsx_chart_label_separator.
 *
 * The `%lxlsx_chart_series_set_labels_separator()` function is used to change the
 * separator between multiple data label items. The default options is a comma
 * separator as shown in the previous example.
 *
 * The available options are:
 *
 * - #LXLSX_CHART_LABEL_SEPARATOR_SEMICOLON: semicolon separator.
 * - #LXLSX_CHART_LABEL_SEPARATOR_PERIOD: a period (dot) separator.
 * - #LXLSX_CHART_LABEL_SEPARATOR_NEWLINE: a newline separator.
 * - #LXLSX_CHART_LABEL_SEPARATOR_SPACE: a space separator.
 *
 * For example:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_options(series, LXLSX_TRUE, LXLSX_TRUE, LXLSX_TRUE);
 *     lxlsx_chart_series_set_labels_separator(series, LXLSX_CHART_LABEL_SEPARATOR_NEWLINE);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels4.png
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_separator(lxlsx_chart_series *series,
                                       uint8_t separator);

/**
 * @brief Set the data label position for a series.
 *
 * @param series   A series object created via `lxlsx_chart_add_series()`.
 * @param position The data label position: #lxlsx_chart_label_position.
 *
 * The `%lxlsx_chart_series_set_labels_position()` function sets the position of
 * the labels in the data series:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_position(series, LXLSX_CHART_LABEL_POSITION_ABOVE);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels5.png
 *
 * In Excel the allowable data label positions vary for different chart
 * types. The allowable, and default, positions are:
 *
 * | Position                              | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
 * | :------------------------------------ | :------------ | :------------ | :------------ | :------------ |
 * | #LXLSX_CHART_LABEL_POSITION_CENTER      | Yes           | Yes           | Yes           | Yes (default) |
 * | #LXLSX_CHART_LABEL_POSITION_RIGHT       | Yes (default) |               |               |               |
 * | #LXLSX_CHART_LABEL_POSITION_LEFT        | Yes           |               |               |               |
 * | #LXLSX_CHART_LABEL_POSITION_ABOVE       | Yes           |               |               |               |
 * | #LXLSX_CHART_LABEL_POSITION_BELOW       | Yes           |               |               |               |
 * | #LXLSX_CHART_LABEL_POSITION_INSIDE_BASE |               | Yes           |               |               |
 * | #LXLSX_CHART_LABEL_POSITION_INSIDE_END  |               | Yes           | Yes           |               |
 * | #LXLSX_CHART_LABEL_POSITION_OUTSIDE_END |               | Yes (default) | Yes           |               |
 * | #LXLSX_CHART_LABEL_POSITION_BEST_FIT    |               |               | Yes (default) |               |
 *
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_position(lxlsx_chart_series *series,
                                      uint8_t position);

/**
 * @brief Set leader lines for Pie and Doughnut charts.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_labels_leader_line()` function  is used to turn on
 * leader lines for the data label of a series. It is mainly used for pie
 * or doughnut charts:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_leader_line(series);
 * @endcode
 *
 * @note Even when leader lines are turned on they aren't automatically
 *       visible in Excel or XlsxWriter. Due to an Excel limitation
 *       (or design) leader lines only appear if the data label is moved
 *       manually or if the data labels are very close and need to be
 *       adjusted automatically.
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_leader_line(lxlsx_chart_series *series);

/**
 * @brief Set the legend key for a data label in a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_labels_legend()` function is used to set the
 * legend key for a data series:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_legend(series);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels6.png
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_legend(lxlsx_chart_series *series);

/**
 * @brief Set the percentage for a Pie/Doughnut data point.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_labels_percentage()` function is used to turn on
 * the display of data labels as a percentage for a series. It is mainly
 * used for pie charts:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_options(series, LXLSX_FALSE, LXLSX_FALSE, LXLSX_FALSE);
 *     lxlsx_chart_series_set_labels_percentage(series);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels7.png
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_percentage(lxlsx_chart_series *series);

/**
 * @brief Set the number format for chart data labels in a series.
 *
 * @param series     A series object created via `lxlsx_chart_add_series()`.
 * @param num_format The number format string.
 *
 * The `%lxlsx_chart_series_set_labels_num_format()` function is used to set the
 * number format for data labels:
 *
 * @code
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_num_format(series, "$0.00");
 * @endcode
 *
 * @image html lxlsx_chart_data_labels8.png
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `lxlsx_format_set_num_format()`.
 *
 * For more information see @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_num_format(lxlsx_chart_series *series,
                                        const char *num_format);

/**
 * @brief Set the font properties for chart data labels in a series
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param font   A pointer to a chart #lxlsx_chart_font font struct.
 *
 *
 * The `%lxlsx_chart_series_set_labels_font()` function is used to set the font
 * for data labels:
 *
 * @code
 *     lxlsx_chart_font font = {.name = "Consolas", .color = LXLSX_COLOR_RED};
 *
 *     lxlsx_chart_series_set_labels(series);
 *     lxlsx_chart_series_set_labels_font(series, &font);
 * @endcode
 *
 * @image html lxlsx_chart_data_labels9.png
 *
 * For more information see @ref lxlsx_chart_fonts and @ref lxlsx_chart_labels.
 *
 */
void lxlsx_chart_series_set_labels_font(lxlsx_chart_series *series,
                                  lxlsx_chart_font *font);

/**
 * @brief Set the line properties for the data labels in a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param line   A #lxlsx_chart_line struct.
 *
 * Set the line/border properties of the data labels in a chart series:
 *
 * @code
 *     lxlsx_chart_line line = {.color = LXLSX_COLOR_RED};
 *     lxlsx_chart_fill fill = {.color = LXLSX_COLOR_YELLOW};
 *
 *     lxlsx_chart_series_set_labels_line(series, &line);
 *     lxlsx_chart_series_set_labels_fill(series, &fill);
 *
 * @endcode
 *
 * @image html lxlsx_chart_data_labels24.png
 *
 * For more information see @ref lxlsx_chart_lines and @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_line(lxlsx_chart_series *series,
                                  lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for the data labels in a chart series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param fill   A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of the data labels in a chart series:
 *
 * @code
 *     lxlsx_chart_fill fill = {.color = LXLSX_COLOR_YELLOW};
 *
 *     lxlsx_chart_series_set_labels_fill(series, &fill);
 * @endcode
 *
 * See the example and image above and also see @ref lxlsx_chart_fills and
 * @ref lxlsx_chart_labels.
 */
void lxlsx_chart_series_set_labels_fill(lxlsx_chart_series *series,
                                  lxlsx_chart_fill *fill);

/**
 * @brief Set the pattern properties for the data labels in a chart series.
 *
 * @param series  A series object created via `lxlsx_chart_add_series()`.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of the data labels in a chart series:
 *
 * @code
 *     lxlsx_chart_series_set_labels_pattern(series, &pattern);
 * @endcode
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 */
void lxlsx_chart_series_set_labels_pattern(lxlsx_chart_series *series,
                                     lxlsx_chart_pattern *pattern);

/**
 * @brief Turn on a trendline for a chart data series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param type   The type of trendline: #lxlsx_chart_trendline_type.
 * @param value  The order/period value for polynomial and moving average
 *               trendlines.
 *
 * A trendline can be added to a chart series to indicate trends in the data
 * such as a moving average or a polynomial fit. The trendlines types are
 * shown in the following Excel dialog:
 *
 * @image html lxlsx_chart_trendline0.png
 *
 * The `%lxlsx_chart_series_set_trendline()` function turns on these trendlines for
 * a data series:
 *
 * @code
 *     chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_LINE);
 *     series = lxlsx_chart_add_series(chart, NULL, "Sheet1!$A$1:$A$6");
 *
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 * @endcode
 *
 * @image html lxlsx_chart_trendline2.png
 *
 * The `value` parameter corresponds to *order* for a polynomial trendline
 * and *period* for a Moving Average trendline. It both cases it must be >= 2.
 * The `value` parameter  is ignored for all other trendlines:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_AVERAGE, 2);
 * @endcode
 *
 * @image html lxlsx_chart_trendline3.png
 *
 * The allowable values for the the trendline `type` are:
 *
 * - #LXLSX_CHART_TRENDLINE_TYPE_LINEAR: Linear trendline.
 * - #LXLSX_CHART_TRENDLINE_TYPE_LOG: Logarithm trendline.
 * - #LXLSX_CHART_TRENDLINE_TYPE_POLY: Polynomial trendline. The `value`
 *   parameter corresponds to *order*.
 * - #LXLSX_CHART_TRENDLINE_TYPE_POWER: Power trendline.
 * - #LXLSX_CHART_TRENDLINE_TYPE_EXP: Exponential trendline.
 * - #LXLSX_CHART_TRENDLINE_TYPE_AVERAGE: Moving Average trendline. The `value`
 *   parameter corresponds to *period*.
 *
 * Other trendline options, such as those shown in the following Excel
 * dialog, can be set using the functions below.
 *
 * @image html lxlsx_chart_trendline1.png
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline(lxlsx_chart_series *series, uint8_t type,
                                uint8_t value);

/**
 * @brief Set the trendline forecast for a chart data series.
 *
 * @param series   A series object created via `lxlsx_chart_add_series()`.
 * @param forward  The forward period.
 * @param backward The backwards period.
 *
 * The `%lxlsx_chart_series_set_trendline_forecast()` function sets the forward
 * and backward forecast periods for the trendline:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_forecast(series, 0.5, 0.5);
 * @endcode
 *
 * @image html lxlsx_chart_trendline4.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline_forecast(lxlsx_chart_series *series,
                                         double forward, double backward);

/**
 * @brief Display the equation of a trendline for a chart data series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_trendline_equation()` function displays the
 * equation of the trendline on the chart:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_equation(series);
 * @endcode
 *
 * @image html lxlsx_chart_trendline5.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline_equation(lxlsx_chart_series *series);

/**
 * @brief Display the R squared value of a trendline for a chart data series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 *
 * The `%lxlsx_chart_series_set_trendline_r_squared()` function displays the
 * R-squared value for the trendline on the chart:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_r_squared(series);
 * @endcode
 *
 * @image html lxlsx_chart_trendline6.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline_r_squared(lxlsx_chart_series *series);

/**
 * @brief Set the trendline Y-axis intercept for a chart data series.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param intercept Y-axis intercept value.
 *
 * The `%lxlsx_chart_series_set_trendline_intercept()` function sets the Y-axis
 * intercept for the trendline:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_equation(series);
 *     lxlsx_chart_series_set_trendline_intercept(series, 0.8);
 * @endcode
 *
 * @image html lxlsx_chart_trendline7.png
 *
 * As can be seen from the equation on the chart the intercept point
 * (when X=0) is the same as the value set in the equation.
 *
 * @note The intercept feature is only available in Excel for Exponential,
 *       Linear and Polynomial trendline types.
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline_intercept(lxlsx_chart_series *series,
                                          double intercept);

/**
 * @brief Set the trendline name for a chart data series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param name   The name of the trendline to display in the legend.
 *
 * The `%lxlsx_chart_series_set_trendline_name()` function sets the name of the
 * trendline that is displayed in the chart legend. In the examples above
 * the trendlines are displayed with default names like "Linear (Series 1)"
 * and "2 per Mov. Avg. (Series 1)". If these names are too verbose or not
 * descriptive enough you can set your own trendline name:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_name(series, "My trendline");
 * @endcode
 *
 * @image html lxlsx_chart_trendline8.png
 *
 * It is often preferable to turn off the trendline caption in the legend.
 * This is down in Excel by deleting the trendline name from the legend.
 * In libxlsxwriter this is done using the `lxlsx_chart_legend_delete_series()`
 * function to delete the zero based series numbers:
 *
 * @code
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *
 *     // Delete the series name for the second series (=1 in zero base).
 *     // The -1 value indicates the end of the array of values.
 *     int16_t names[] = {1, -1};
 *     lxlsx_chart_legend_delete_series(chart, names);
 * @endcode
 *
 * @image html lxlsx_chart_trendline9.png
 *
 * For more information see @ref lxlsx_chart_trendlines.
 */
void lxlsx_chart_series_set_trendline_name(lxlsx_chart_series *series,
                                     const char *name);

/**
 * @brief Set the trendline line properties for a chart data series.
 *
 * @param series A series object created via `lxlsx_chart_add_series()`.
 * @param line   A #lxlsx_chart_line struct.
 *
 * The `%lxlsx_chart_series_set_trendline_line()` function is used to set the line
 * properties of a trendline:
 *
 * @code
 *     lxlsx_chart_line line = {.color     = LXLSX_COLOR_RED,
 *                            .dash_type = LXLSX_CHART_LINE_DASH_LONG_DASH};
 *
 *     lxlsx_chart_series_set_trendline(series, LXLSX_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     lxlsx_chart_series_set_trendline_line(series, &line);
 * @endcode
 *
 * @image html lxlsx_chart_trendline10.png
 *
 * For more information see @ref lxlsx_chart_trendlines and @ref lxlsx_chart_lines.
 */
void lxlsx_chart_series_set_trendline_line(lxlsx_chart_series *series,
                                     lxlsx_chart_line *line);
/**
 * @brief           Get a pointer to X or Y error bars from a chart series.
 *
 * @param series    A series object created via `lxlsx_chart_add_series()`.
 * @param axis_type The axis type (X or Y): #lxlsx_chart_error_bar_axis.
 *
 * The `%lxlsx_chart_series_get_error_bars()` function returns a pointer to the
 * error bars of a series based on the type of #lxlsx_chart_error_bar_axis:
 *
 * @code
 *     lxlsx_series_error_bars *x_error_bars;
 *     lxlsx_series_error_bars *y_error_bars;
 *
 *     x_error_bars = lxlsx_chart_series_get_error_bars(series, LXLSX_CHART_ERROR_BAR_AXIS_X);
 *     y_error_bars = lxlsx_chart_series_get_error_bars(series, LXLSX_CHART_ERROR_BAR_AXIS_Y);
 *
 *     // Use the error bar pointers.
 *     lxlsx_chart_series_set_error_bars(x_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 *
 *     lxlsx_chart_series_set_error_bars(y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * Note, the series error bars can also be accessed directly:
 *
 * @code
 *     // Equivalent to the above example, without function calls.
 *     lxlsx_chart_series_set_error_bars(series->x_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 *
 *     lxlsx_chart_series_set_error_bars(series->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * @return Pointer to the series error bars, or NULL if not found.
 */

lxlsx_series_error_bars *lxlsx_chart_series_get_error_bars(lxlsx_chart_series *series, lxlsx_chart_error_bar_axis
                                                   axis_type);

/**
 * Set the X or Y error bars for a chart series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param type       The type of error bar: #lxlsx_chart_error_bar_type.
 * @param value      The error value.
 *
 * Error bars can be added to a chart series to indicate error bounds in the
 * data. The error bars can be vertical `y_error_bars` (the most common type)
 * or horizontal `x_error_bars` (for Bar and Scatter charts only).
 *
 * @image html lxlsx_chart_error_bars0.png
 *
 * The `%lxlsx_chart_series_set_error_bars()` function sets the error bar type
 * and value associated with the type:
 *
 * @code
 *     lxlsx_chart_series *series = lxlsx_chart_add_series(chart,
 *                                                 "=Sheet1!$A$1:$A$5",
 *                                                 "=Sheet1!$B$1:$B$5");
 *
 *     lxlsx_chart_series_set_error_bars(series->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * @image html lxlsx_chart_error_bars1.png
 *
 * The error bar types that be used are:
 *
 * - #LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR: Standard error.
 * - #LXLSX_CHART_ERROR_BAR_TYPE_FIXED: Fixed value.
 * - #LXLSX_CHART_ERROR_BAR_TYPE_PERCENTAGE: Percentage.
 * - #LXLSX_CHART_ERROR_BAR_TYPE_STD_DEV: Standard deviation(s).
 *
 * @note Custom error bars are not currently supported.
 *
 * All error bar types, apart from Standard error, should have a valid
 * value to set the error range:
 *
 * @code
 *     lxlsx_chart_series_set_error_bars(series1->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_FIXED, 2);
 *
 *     lxlsx_chart_series_set_error_bars(series2->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_PERCENTAGE, 5);
 *
 *     lxlsx_chart_series_set_error_bars(series3->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 * @endcode
 *
 * For the Standard error type the value is ignored.
 *
 * For more information see @ref lxlsx_chart_error_bars.
 */
void lxlsx_chart_series_set_error_bars(lxlsx_series_error_bars *error_bars,
                                 uint8_t type, double value);

/**
 * @brief Set the direction (up, down or both) of the error bars for a chart
 *        series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param direction  The bar direction: #lxlsx_chart_error_bar_direction.
 *
 * The `%lxlsx_chart_series_set_error_bars_direction()` function sets the
 * direction of the error bars:
 *
 * @code
 *     lxlsx_chart_series_set_error_bars(series->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     lxlsx_chart_series_set_error_bars_direction(series->y_error_bars,
 *                                           LXLSX_CHART_ERROR_BAR_DIR_PLUS);
 * @endcode
 *
 * @image html lxlsx_chart_error_bars2.png
 *
 * The valid directions are:
 *
 * - #LXLSX_CHART_ERROR_BAR_DIR_BOTH: Error bar extends in both directions.
 *   The default.
 * - #LXLSX_CHART_ERROR_BAR_DIR_PLUS: Error bar extends in positive direction.
 * - #LXLSX_CHART_ERROR_BAR_DIR_MINUS: Error bar extends in negative direction.
 *
 * For more information see @ref lxlsx_chart_error_bars.
 */
void lxlsx_chart_series_set_error_bars_direction(lxlsx_series_error_bars *error_bars,
                                           uint8_t direction);

/**
 * @brief Set the end cap type for the error bars of a chart series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param endcap     The error bar end cap type: #lxlsx_chart_error_bar_cap .
 *
 * The `%lxlsx_chart_series_set_error_bars_endcap()` function sets the end cap
 * type for the error bars:
 *
 * @code
 *     lxlsx_chart_series_set_error_bars(series->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     lxlsx_chart_series_set_error_bars_endcap(series->y_error_bars,
                                          LXLSX_CHART_ERROR_BAR_NO_CAP);
 * @endcode
 *
 * @image html lxlsx_chart_error_bars3.png
 *
 * The valid values are:
 *
 * - #LXLSX_CHART_ERROR_BAR_END_CAP: Flat end cap. The default.
 * - #LXLSX_CHART_ERROR_BAR_NO_CAP: No end cap.
 *
 * For more information see @ref lxlsx_chart_error_bars.
 */
void lxlsx_chart_series_set_error_bars_endcap(lxlsx_series_error_bars *error_bars,
                                        uint8_t endcap);

/**
 * @brief Set the line properties for a chart series error bars.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param line       A #lxlsx_chart_line struct.
 *
 * The `%lxlsx_chart_series_set_error_bars_line()` function sets the line
 * properties for the error bars:
 *
 * @code
 *     lxlsx_chart_line line = {.color     = LXLSX_COLOR_RED,
 *                            .dash_type = LXLSX_CHART_LINE_DASH_ROUND_DOT};
 *
 *     lxlsx_chart_series_set_error_bars(series->y_error_bars,
 *                                 LXLSX_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     lxlsx_chart_series_set_error_bars_line(series->y_error_bars, &line);
 * @endcode
 *
 * @image html lxlsx_chart_error_bars4.png
 *
 * For more information see @ref lxlsx_chart_lines and @ref lxlsx_chart_error_bars.
 */
void lxlsx_chart_series_set_error_bars_line(lxlsx_series_error_bars *error_bars,
                                      lxlsx_chart_line *line);

/**
 * @brief           Get an axis pointer from a chart.
 *
 * @param chart     Pointer to a lxlsx_chart instance to be configured.
 * @param axis_type The axis type (X or Y): #lxlsx_chart_axis_type.
 *
 * The `%lxlsx_chart_axis_get()` function returns a pointer to a chart axis based
 * on the  #lxlsx_chart_axis_type:
 *
 * @code
 *     lxlsx_chart_axis *x_axis = lxlsx_chart_axis_get(chart, LXLSX_CHART_AXIS_TYPE_X);
 *     lxlsx_chart_axis *y_axis = lxlsx_chart_axis_get(chart, LXLSX_CHART_AXIS_TYPE_Y);
 *
 *     // Use the axis pointer in other functions.
 *     lxlsx_chart_axis_major_gridlines_set_visible(x_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_major_gridlines_set_visible(y_axis, LXLSX_TRUE);
 * @endcode
 *
 * Note, the axis pointer can also be accessed directly:
 *
 * @code
 *     // Equivalent to the above example, without function calls.
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->x_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->y_axis, LXLSX_TRUE);
 * @endcode
 *
 * @return Pointer to the chart axis, or NULL if not found.
 */
lxlsx_chart_axis *lxlsx_chart_axis_get(lxlsx_chart *chart,
                               lxlsx_chart_axis_type axis_type);

/**
 * @brief Set the name caption of the an axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param name The name caption of the axis.
 *
 * The `%lxlsx_chart_axis_set_name()` function sets the name (also known as title or
 * caption) for an axis. It can be used for the X or Y axes. The name is
 * displayed below an X axis and to the side of a Y axis.
 *
 * @code
 *     lxlsx_chart_axis_set_name(chart->x_axis, "Earnings per Quarter");
 *     lxlsx_chart_axis_set_name(chart->y_axis, "US Dollars (Millions)");
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_name.png
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     lxlsx_chart_axis_set_name(chart->x_axis, "=Sheet1!$B$1");
 * @endcode
 *
 * See also the `lxlsx_chart_axis_set_name_range()` function to see how to set the
 * name formula programmatically.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_name(lxlsx_chart_axis *axis, const char *name);

/**
 * @brief Set a chart axis name formula using row and column values.
 *
 * @param axis      A pointer to a chart #lxlsx_chart_axis object.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%lxlsx_chart_axis_set_name_range()` function can be used to set an axis name
 * range and is an alternative to using `lxlsx_chart_axis_set_name()` and a string
 * formula:
 *
 * @code
 *     lxlsx_chart_axis_set_name_range(chart->x_axis, "Sheet1", 1, 0);
 *     lxlsx_chart_axis_set_name_range(chart->y_axis, "Sheet1", 2, 0);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_name_range(lxlsx_chart_axis *axis, const char *sheetname,
                               lxlsx_row_t row, lxlsx_col_t col);

/**
 * @brief Set the manual position of the chart axis name.
 *
 * @param axis   A pointer to a chart #lxlsx_chart_axis object.
 * @param layout A pointer to a chart #lxlsx_chart_layout struct.
 *
 * This function is used to simulate setting the manual position of a chart
 * axis name. See @ref lxlsx_chart_layout for more information.
 */
void lxlsx_chart_axis_set_name_layout(lxlsx_chart_axis *axis,
                                lxlsx_chart_layout *layout);

/**
 * @brief Set the font properties for a chart axis name.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param font A pointer to a chart #lxlsx_chart_font font struct.
 *
 * The `%lxlsx_chart_axis_set_name_font()` function is used to set the font of an
 * axis name:
 *
 * @code
 *     lxlsx_chart_font font = {.bold = LXLSX_TRUE, .color = LXLSX_COLOR_BLUE};
 *
 *     lxlsx_chart_axis_set_name(chart->x_axis, "Yearly data");
 *     lxlsx_chart_axis_set_name_font(chart->x_axis, &font);
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_name_font.png
 *
 * For more information see @ref lxlsx_chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_name_font(lxlsx_chart_axis *axis, lxlsx_chart_font *font);

/**
 * @brief Set the font properties for the numbers of a chart axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param font A pointer to a chart #lxlsx_chart_font font struct.
 *
 * The `%lxlsx_chart_axis_set_num_font()` function is used to set the font of the
 * numbers on an axis:
 *
 * @code
 *     lxlsx_chart_font font = {.bold = LXLSX_TRUE, .color = LXLSX_COLOR_BLUE};
 *
 *     lxlsx_chart_axis_set_num_font(chart->x_axis, &font1);
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_num_font.png
 *
 * For more information see @ref lxlsx_chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_num_font(lxlsx_chart_axis *axis, lxlsx_chart_font *font);

/**
 * @brief Set the number format for a chart axis.
 *
 * @param axis       A pointer to a chart #lxlsx_chart_axis object.
 * @param num_format The number format string.
 *
 * The `%lxlsx_chart_axis_set_num_format()` function is used to set the format of
 * the numbers on an axis:
 *
 * @code
 *     lxlsx_chart_axis_set_num_format(chart->x_axis, "0.00%");
 *     lxlsx_chart_axis_set_num_format(chart->y_axis, "$#,##0.00");
 * @endcode
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `lxlsx_format_set_num_format()`.
 *
 * @image html lxlsx_chart_axis_num_format.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_num_format(lxlsx_chart_axis *axis, const char *num_format);

/**
 * @brief Set the line properties for a chart axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param line A #lxlsx_chart_line struct.
 *
 * Set the line properties of a chart axis:
 *
 * @code
 *     // Hide the Y axis.
 *     lxlsx_chart_line line = {.none = LXLSX_TRUE};
 *
 *     lxlsx_chart_axis_set_line(chart->y_axis, &line);
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_line.png
 *
 * For more information see @ref lxlsx_chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_line(lxlsx_chart_axis *axis, lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for a chart axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param fill A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of a chart axis:
 *
 * @code
 *     lxlsx_chart_fill fill = {.color = LXLSX_COLOR_YELLOW};
 *
 *     lxlsx_chart_axis_set_fill(chart->y_axis, &fill);
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_fill.png
 *
 * For more information see @ref lxlsx_chart_fills.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_fill(lxlsx_chart_axis *axis, lxlsx_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chart axis.
 *
 * @param axis    A pointer to a chart #lxlsx_chart_axis object.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of a chart axis:
 *
 * @code
 *     lxlsx_chart_axis_set_pattern(chart->y_axis, &pattern);
 * @endcode
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_pattern(lxlsx_chart_axis *axis, lxlsx_chart_pattern *pattern);

/**
 * @brief Reverse the order of the axis categories or values.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 *
 * Reverse the order of the axis categories or values:
 *
 * @code
 *     lxlsx_chart_axis_set_reverse(chart->x_axis);
 * @endcode
 *
 * @image html lxlsx_chart_reverse.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_reverse(lxlsx_chart_axis *axis);

/**
 * @brief Set the position that the axis will cross the opposite axis.
 *
 * @param axis  A pointer to a chart #lxlsx_chart_axis object.
 * @param value The category or value that the axis crosses at.
 *
 * Set the position that the axis will cross the opposite axis:
 *
 * @code
 *     lxlsx_chart_axis_set_crossing(chart->x_axis, 3);
 *     lxlsx_chart_axis_set_crossing(chart->y_axis, 8);
 * @endcode
 *
 * @image html lxlsx_chart_crossing1.png
 *
 * If crossing is omitted (the default) the crossing will be set automatically
 * by Excel based on the chart data.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_crossing(lxlsx_chart_axis *axis, double value);

/**
 * @brief Set the opposite axis crossing position as the axis maximum.
 *
 * @param axis  A pointer to a chart #lxlsx_chart_axis object.
 *
 * Set the position that the opposite axis will cross as the axis maximum.
 * The default axis crossing position is generally the axis minimum so this
 * function can be used to reverse the location of the axes without reversing
 * the number sequence:
 *
 * @code
 *     lxlsx_chart_axis_set_crossing_max(chart->x_axis);
 *     lxlsx_chart_axis_set_crossing_max(chart->y_axis);
 * @endcode
 *
 * @image html lxlsx_chart_crossing2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_crossing_max(lxlsx_chart_axis *axis);

/**
 * @brief Set the opposite axis crossing position as the axis minimum.
 *
 * @param axis  A pointer to a chart #lxlsx_chart_axis object.
 *
 * Set the position that the opposite axis will cross as the axis minimum.
 * The default axis crossing position is generally the axis minimum so this
 * function can be used to reverse the location of the axes without reversing
 * the number sequence:
 *
 * @code
 *     lxlsx_chart_axis_set_crossing_min(chart->x_axis);
 *     lxlsx_chart_axis_set_crossing_min(chart->y_axis);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_crossing_min(lxlsx_chart_axis *axis);

/**
 * @brief Turn off/hide an axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 *
 * Turn off, hide, a chart axis:
 *
 * @code
 *     lxlsx_chart_axis_off(chart->x_axis);
 * @endcode
 *
 * @image html lxlsx_chart_axis_off.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_off(lxlsx_chart_axis *axis);

/**
 * @brief Position a category axis on or between the axis tick marks.
 *
 * @param axis     A pointer to a chart #lxlsx_chart_axis object.
 * @param position A #lxlsx_chart_axis_tick_position value.
 *
 * Position a category axis horizontally on, or between, the axis tick marks.
 *
 * There are two allowable values:
 *
 * - #LXLSX_CHART_AXIS_POSITION_ON_TICK
 * - #LXLSX_CHART_AXIS_POSITION_BETWEEN
 *
 * @code
 *     lxlsx_chart_axis_set_position(chart->x_axis, LXLSX_CHART_AXIS_POSITION_BETWEEN);
 * @endcode
 *
 * @image html lxlsx_chart_axis_set_position.png
 *
 * **Axis types**: This function is applicable to category axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_position(lxlsx_chart_axis *axis, uint8_t position);

/**
 * @brief Position the axis labels.
 *
 * @param axis     A pointer to a chart #lxlsx_chart_axis object.
 * @param position A #lxlsx_chart_axis_label_position value.
 *
 * Position the axis labels for the chart. The labels are the numbers, or
 * strings or dates, on the axis that indicate the categories or values of
 * the axis.
 *
 * For example:
 *
 * @code
 *     lxlsx_chart_axis_set_label_position(chart->x_axis, LXLSX_CHART_AXIS_LABEL_POSITION_HIGH);
 *     lxlsx_chart_axis_set_label_position(chart->y_axis, LXLSX_CHART_AXIS_LABEL_POSITION_HIGH);
 * @endcode
 *
 * @image html lxlsx_chart_label_position2.png
 *
 * The allowable values:
 *
 * - #LXLSX_CHART_AXIS_LABEL_POSITION_NEXT_TO - The default.
 * - #LXLSX_CHART_AXIS_LABEL_POSITION_HIGH - Also right for vertical axes.
 * - #LXLSX_CHART_AXIS_LABEL_POSITION_LOW - Also left for vertical axes.
 * - #LXLSX_CHART_AXIS_LABEL_POSITION_NONE
 *
 * @image html lxlsx_chart_label_position1.png
 *
 * The #LXLSX_CHART_AXIS_LABEL_POSITION_NONE turns off the axis labels. This
 * is slightly different from `lxlsx_chart_axis_off()` which also turns off the
 * labels but also turns off tick marks.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_label_position(lxlsx_chart_axis *axis, uint8_t position);

/**
 * @brief Set the alignment of the axis labels.
 *
 * @param axis  A pointer to a chart #lxlsx_chart_axis object.
 * @param align A #lxlsx_chart_axis_label_alignment value.
 *
 * Position the category axis labels for the chart. The labels are the
 * numbers, or strings or dates, on the axis that indicate the categories
 * of the axis.
 *
 * The allowable values:
 *
 * - #LXLSX_CHART_AXIS_LABEL_ALIGN_CENTER - Align label center (default).
 * - #LXLSX_CHART_AXIS_LABEL_ALIGN_LEFT - Align label left.
 * - #LXLSX_CHART_AXIS_LABEL_ALIGN_RIGHT - Align label right.
 *
 * @code
 *     lxlsx_chart_axis_set_label_align(chart->x_axis, LXLSX_CHART_AXIS_LABEL_ALIGN_RIGHT);
 * @endcode
 *
 * **Axis types**: This function is applicable to category axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_label_align(lxlsx_chart_axis *axis, uint8_t align);

/**
 * @brief Set the minimum value for a chart axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param min  Minimum value for chart axis. Value axes only.
 *
 * Set the minimum value for the axis range.
 *
 * @code
 *     lxlsx_chart_axis_set_min(chart->y_axis, -4);
 *     lxlsx_chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * @image html lxlsx_chart_max_min.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_min(lxlsx_chart_axis *axis, double min);

/**
 * @brief Set the maximum value for a chart axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param max  Maximum value for chart axis. Value axes only.
 *
 * Set the maximum value for the axis range.
 *
 * @code
 *     lxlsx_chart_axis_set_min(chart->y_axis, -4);
 *     lxlsx_chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * See the above image.
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_max(lxlsx_chart_axis *axis, double max);

/**
 * @brief Set the log base of the axis range.
 *
 * @param axis     A pointer to a chart #lxlsx_chart_axis object.
 * @param log_base The log base for value axis. Value axes only.
 *
 * Set the log base for the axis:
 *
 * @code
 *     lxlsx_chart_axis_set_log_base(chart->y_axis, 10);
 * @endcode
 *
 * @image html lxlsx_chart_log_base.png
 *
 * The allowable range of values for the log base in Excel is between 2 and
 * 1000.
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_log_base(lxlsx_chart_axis *axis, uint16_t log_base);

/**
 * @brief Set the major axis tick mark type.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param type The tick mark type, defined by #lxlsx_chart_tick_mark.
 *
 * Set the type of the major axis tick mark:
 *
 * @code
 *     lxlsx_chart_axis_set_major_tick_mark(chart->x_axis, LXLSX_CHART_AXIS_TICK_MARK_CROSSING);
 *     lxlsx_chart_axis_set_minor_tick_mark(chart->x_axis, LXLSX_CHART_AXIS_TICK_MARK_INSIDE);
 *
 *     lxlsx_chart_axis_set_major_tick_mark(chart->x_axis, LXLSX_CHART_AXIS_TICK_MARK_OUTSIDE);
 *     lxlsx_chart_axis_set_minor_tick_mark(chart->y_axis, LXLSX_CHART_AXIS_TICK_MARK_INSIDE);
 *
 *     // Hide the default gridlines so the tick marks are visible.
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->y_axis, LXLSX_FALSE);
 * @endcode
 *
 * @image html lxlsx_chart_tick_marks.png
 *
 * The tick mark types are:
 *
 * - #LXLSX_CHART_AXIS_TICK_MARK_NONE
 * - #LXLSX_CHART_AXIS_TICK_MARK_INSIDE
 * - #LXLSX_CHART_AXIS_TICK_MARK_OUTSIDE
 * - #LXLSX_CHART_AXIS_TICK_MARK_CROSSING
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_major_tick_mark(lxlsx_chart_axis *axis, uint8_t type);

/**
 * @brief Set the minor axis tick mark type.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param type The tick mark type, defined by #lxlsx_chart_tick_mark.
 *
 * Set the type of the minor axis tick mark:
 *
 * @code
 *     lxlsx_chart_axis_set_minor_tick_mark(chart->x_axis, LXLSX_CHART_AXIS_TICK_MARK_INSIDE);
 * @endcode
 *
 * See the image and example above.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_minor_tick_mark(lxlsx_chart_axis *axis, uint8_t type);

/**
 * @brief Set the interval between category values.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param unit The interval between the categories.
 *
 * Set the interval between the category values. The default interval is 1
 * which gives the intervals shown in the charts above:
 *
 *     1, 2, 3, 4, 5, etc.
 *
 * Setting it to 2 gives:
 *
 *     1, 3, 5, 7, etc.
 *
 * For example:
 *
 * @code
 *     lxlsx_chart_axis_set_interval_unit(chart->x_axis, 2);
 * @endcode
 *
 * @image html lxlsx_chart_set_interval1.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_interval_unit(lxlsx_chart_axis *axis, uint16_t unit);

/**
 * @brief Set the interval between category tick marks.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param unit The interval between the category ticks.
 *
 * Set the interval between the category tick marks. The default interval is 1
 * between each category but it can be set to other integer values:
 *
 * @code
 *     lxlsx_chart_axis_set_interval_tick(chart->x_axis, 2);
 * @endcode
 *
 * @image html lxlsx_chart_set_interval2.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_interval_tick(lxlsx_chart_axis *axis, uint16_t unit);

/**
 * @brief Set the increment of the major units in the axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param unit The increment of the major units.
 *
 * Set the increment of the major units in the axis range.
 *
 * @code
 *     // Turn on the minor gridline (it is off by default).
 *     lxlsx_chart_axis_minor_gridlines_set_visible(chart->y_axis, LXLSX_TRUE);
 *
 *     lxlsx_chart_axis_set_major_unit(chart->y_axis, 4);
 *     lxlsx_chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * @image html lxlsx_chart_set_major_units.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_major_unit(lxlsx_chart_axis *axis, double unit);

/**
 * @brief Set the increment of the minor units in the axis.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param unit The increment of the minor units.
 *
 * Set the increment of the minor units in the axis range.
 *
 * @code
 *     lxlsx_chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * See the image above
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_minor_unit(lxlsx_chart_axis *axis, double unit);

/**
 * @brief Set the display units for a value axis.
 *
 * @param axis  A pointer to a chart #lxlsx_chart_axis object.
 * @param units The display units: #lxlsx_chart_axis_display_unit.
 *
 * Set the display units for the axis. This can be useful if the axis numbers
 * are very large but you don't want to represent them in scientific notation:
 *
 * @code
 *     lxlsx_chart_axis_set_display_units(chart->x_axis, LXLSX_CHART_AXIS_UNITS_THOUSANDS);
 *     lxlsx_chart_axis_set_display_units(chart->y_axis, LXLSX_CHART_AXIS_UNITS_MILLIONS);
 * @endcode
 *
 * @image html lxlsx_chart_display_units.png
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_display_units(lxlsx_chart_axis *axis, uint8_t units);

/**
 * @brief Turn on/off the display units for a value axis.

 * @param axis    A pointer to a chart #lxlsx_chart_axis object.
 * @param visible Turn off/on the display units. (0/1)
 *
 * Turn on or off the display units for the axis. This option is set on
 * automatically by `lxlsx_chart_axis_set_display_units()`.
 *
 * @code
 *     lxlsx_chart_axis_set_display_units_visible(chart->y_axis, LXLSX_TRUE);
 * @endcode
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_set_display_units_visible(lxlsx_chart_axis *axis,
                                          uint8_t visible);

/**
 * @brief Turn on/off the major gridlines for an axis.
 *
 * @param axis    A pointer to a chart #lxlsx_chart_axis object.
 * @param visible Turn off/on the major gridline. (0/1)
 *
 * Turn on or off the major gridlines for an X or Y axis. In most Excel charts
 * the Y axis major gridlines are on by default and the X axis major
 * gridlines are off by default.
 *
 * Example:
 *
 * @code
 *     // Reverse the normal visible/hidden gridlines for a column chart.
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->x_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->y_axis, LXLSX_FALSE);
 * @endcode
 *
 * @image html lxlsx_chart_gridline1.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_major_gridlines_set_visible(lxlsx_chart_axis *axis,
                                            uint8_t visible);

/**
 * @brief Turn on/off the minor gridlines for an axis.
 *
 * @param axis    A pointer to a chart #lxlsx_chart_axis object.
 * @param visible Turn off/on the minor gridline. (0/1)
 *
 * Turn on or off the minor gridlines for an X or Y axis. In most Excel charts
 * the X and Y axis minor gridlines are off by default.
 *
 * Example, turn on all major and minor gridlines:
 *
 * @code
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->x_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_minor_gridlines_set_visible(chart->x_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_major_gridlines_set_visible(chart->y_axis, LXLSX_TRUE);
 *     lxlsx_chart_axis_minor_gridlines_set_visible(chart->y_axis, LXLSX_TRUE);
 * @endcode
 *
 * @image html lxlsx_chart_gridline2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_minor_gridlines_set_visible(lxlsx_chart_axis *axis,
                                            uint8_t visible);

/**
 * @brief Set the line properties for the chart axis major gridlines.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param line A #lxlsx_chart_line struct.
 *
 * Format the line properties of the major gridlines of a chart:
 *
 * @code
 *     lxlsx_chart_line line1 = {.color = LXLSX_COLOR_RED,
 *                             .width = 0.5,
 *                             .dash_type = LXLSX_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     lxlsx_chart_line line2 = {.color = LXLSX_COLOR_YELLOW};
 *
 *     lxlsx_chart_line line3 = {.width = 1.25,
 *                             .dash_type = LXLSX_CHART_LINE_DASH_DASH};
 *
 *     lxlsx_chart_line line4 = {.color =  0x00B050};
 *
 *     lxlsx_chart_axis_major_gridlines_set_line(chart->x_axis, &line1);
 *     lxlsx_chart_axis_minor_gridlines_set_line(chart->x_axis, &line2);
 *     lxlsx_chart_axis_major_gridlines_set_line(chart->y_axis, &line3);
 *     lxlsx_chart_axis_minor_gridlines_set_line(chart->y_axis, &line4);
 * @endcode
 *
 * @image html lxlsx_chart_gridline3.png
 *
 * For more information see @ref lxlsx_chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_major_gridlines_set_line(lxlsx_chart_axis *axis,
                                         lxlsx_chart_line *line);

/**
 * @brief Set the line properties for the chart axis minor gridlines.
 *
 * @param axis A pointer to a chart #lxlsx_chart_axis object.
 * @param line A #lxlsx_chart_line struct.
 *
 * Format the line properties of the minor gridlines of a chart, see the
 * example above.
 *
 * For more information see @ref lxlsx_chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void lxlsx_chart_axis_minor_gridlines_set_line(lxlsx_chart_axis *axis,
                                         lxlsx_chart_line *line);

/**
 * @brief Set the title of the chart.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param name  The chart title name.
 *
 * The `%lxlsx_chart_title_set_name()` function sets the name (title) for the
 * chart. The name is displayed above the chart.
 *
 * @code
 *     lxlsx_chart_title_set_name(chart, "Year End Results");
 * @endcode
 *
 * @image html lxlsx_chart_title_set_name.png
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     lxlsx_chart_title_set_name(chart, "=Sheet1!$B$1");
 * @endcode
 *
 * See also the `lxlsx_chart_title_set_name_range()` function to see how to set the
 * name formula programmatically.
 *
 * The Excel default is to have no chart title.
 */
void lxlsx_chart_title_set_name(lxlsx_chart *chart, const char *name);

/**
 * @brief Set a chart title formula using row and column values.
 *
 * @param chart     Pointer to a lxlsx_chart instance to be configured.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%lxlsx_chart_title_set_name_range()` function can be used to set a chart
 * title range and is an alternative to using `lxlsx_chart_title_set_name()` and a
 * string formula:
 *
 * @code
 *     lxlsx_chart_title_set_name_range(chart, "Sheet1", 1, 0);
 * @endcode
 */
void lxlsx_chart_title_set_name_range(lxlsx_chart *chart, const char *sheetname,
                                lxlsx_row_t row, lxlsx_col_t col);

/**
 * @brief  Set the font properties for a chart title.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param font  A pointer to a chart #lxlsx_chart_font font struct.
 *
 * The `%lxlsx_chart_title_set_name_font()` function is used to set the font of a
 * chart title:
 *
 * @code
 *     lxlsx_chart_font font = {.color = LXLSX_COLOR_BLUE};
 *
 *     lxlsx_chart_title_set_name(chart, "Year End Results");
 *     lxlsx_chart_title_set_name_font(chart, &font);
 * @endcode
 *
 * @image html lxlsx_chart_title_set_name_font.png
 *
 * In Excel a chart title font is bold by default (as shown in the image
 * above). To turn off bold in the font you cannot use #LXLSX_FALSE (0) since
 * that is indistinguishable from an uninitialized value. Instead you should
 * use #LXLSX_EXPLICIT_FALSE:
 *
 * @code
 *     lxlsx_chart_font font = {.bold = LXLSX_EXPLICIT_FALSE, .color = LXLSX_COLOR_BLUE};
 *
 *     lxlsx_chart_title_set_name(chart, "Year End Results");
 *     lxlsx_chart_title_set_name_font(chart, &font);
 * @endcode
 *
 * @image html lxlsx_chart_title_set_name_font2.png
 *
 * For more information see @ref lxlsx_chart_fonts.
 */
void lxlsx_chart_title_set_name_font(lxlsx_chart *chart, lxlsx_chart_font *font);

/**
 * @brief Turn off an automatic chart title.
 *
 * @param chart  Pointer to a lxlsx_chart instance to be configured.
 *
 * In general in Excel a chart title isn't displayed unless the user
 * explicitly adds one. However, Excel adds an automatic chart title to charts
 * with a single series and a user defined series name. The
 * `lxlsx_chart_title_off()` function allows you to turn off this automatic chart
 * title:
 *
 * @code
 *     lxlsx_chart_title_off(chart);
 * @endcode
 */
void lxlsx_chart_title_off(lxlsx_chart *chart);

/**
 * @brief Set the manual position of the chart title.
 *
 * @param chart  Pointer to a lxlsx_chart instance to be configured.
 * @param layout A pointer to a chart #lxlsx_chart_layout struct.
 *
 * This function is used to simulate setting the manual position of the chart
 * title. See @ref lxlsx_chart_layout for more information.
 */
void lxlsx_chart_title_set_layout(lxlsx_chart *chart, lxlsx_chart_layout *layout);

/**
 * @brief Allow the chart title to overlay the chart.
 *
 * @param chart   Pointer to a lxlsx_chart instance to be configured.
 * @param overlay Turn off/on the overlay. (0/1)
 *
 * This option allows the chart title to overlay the chart when the
 * `lxlsx_chart_title_set_layout()` function.
 */
void lxlsx_chart_title_set_overlay(lxlsx_chart *chart, uint8_t overlay);

/**
 * @brief Set the position of the chart legend.
 *
 * @param chart    Pointer to a lxlsx_chart instance to be configured.
 * @param position The #lxlsx_chart_legend_position value for the legend.
 *
 * The `%lxlsx_chart_legend_set_position()` function is used to set the chart
 * legend to one of the #lxlsx_chart_legend_position values:
 *
 *     LXLSX_CHART_LEGEND_NONE
 *     LXLSX_CHART_LEGEND_RIGHT
 *     LXLSX_CHART_LEGEND_LEFT
 *     LXLSX_CHART_LEGEND_TOP
 *     LXLSX_CHART_LEGEND_BOTTOM
 *     LXLSX_CHART_LEGEND_TOP_RIGHT
 *     LXLSX_CHART_LEGEND_OVERLAY_RIGHT
 *     LXLSX_CHART_LEGEND_OVERLAY_LEFT
 *     LXLSX_CHART_LEGEND_OVERLAY_TOP_RIGHT
 *
 * For example:
 *
 * @code
 *     lxlsx_chart_legend_set_position(chart, LXLSX_CHART_LEGEND_BOTTOM);
 * @endcode
 *
 * @image html lxlsx_chart_legend_bottom.png
 *
 * This function can also be used to turn off a chart legend:
 *
 * @code
 *     lxlsx_chart_legend_set_position(chart, LXLSX_CHART_LEGEND_NONE);
 * @endcode
 *
 * @image html lxlsx_chart_legend_none.png
 *
 */
void lxlsx_chart_legend_set_position(lxlsx_chart *chart, uint8_t position);

/**
 * @brief Set the manual layout of the chart legend.
 *
 * @param chart  Pointer to a lxlsx_chart instance to be configured.
 * @param layout A pointer to a chart #lxlsx_chart_layout struct.
 *
 * This function is used to simulate setting the manual position of the chart
 * legend. See @ref lxlsx_chart_layout for more information.
 */
void lxlsx_chart_legend_set_layout(lxlsx_chart *chart, lxlsx_chart_layout *layout);

/**
 * @brief Set the font properties for a chart legend.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param font  A pointer to a chart #lxlsx_chart_font font struct.
 *
 * The `%lxlsx_chart_legend_set_font()` function is used to set the font of a
 * chart legend:
 *
 * @code
 *     lxlsx_chart_font font = {.bold = LXLSX_TRUE, .color = LXLSX_COLOR_BLUE};
 *
 *     lxlsx_chart_legend_set_font(chart, &font);
 * @endcode
 *
 * @image html lxlsx_chart_legend_set_font.png
 *
 * For more information see @ref lxlsx_chart_fonts.
 */
void lxlsx_chart_legend_set_font(lxlsx_chart *chart, lxlsx_chart_font *font);

/**
 * @brief Remove one or more series from the the legend.
 *
 * @param chart         Pointer to a lxlsx_chart instance to be configured.
 * @param delete_series An array of zero-indexed values to delete from series.
 *
 * @return A #lxlsx_error.
 *
 * The `%lxlsx_chart_legend_delete_series()` function allows you to remove/hide one
 * or more series in a chart legend (the series will still display on the chart).
 *
 * This function takes an array of one or more zero indexed series
 * numbers. The array should be terminated with -1.
 *
 * For example to remove the first and third zero-indexed series from the
 * legend of a chart with 3 series:
 *
 * @code
 *     int16_t series[] = {0, 2, -1};
 *
 *     lxlsx_chart_legend_delete_series(chart, series);
 * @endcode
 *
 * @image html lxlsx_chart_legend_delete.png
 */
lxlsx_error lxlsx_chart_legend_delete_series(lxlsx_chart *chart,
                                     int16_t delete_series[]);

/**
 * @brief Set the line properties for a chartarea.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param line  A #lxlsx_chart_line struct.
 *
 * Set the line/border properties of a chartarea. In Excel the chartarea
 * is the background area behind the chart:
 *
 * @code
 *     lxlsx_chart_line line = {.none  = LXLSX_TRUE};
 *     lxlsx_chart_fill fill = {.color = LXLSX_COLOR_RED};
 *
 *     lxlsx_chart_chartarea_set_line(chart, &line);
 *     lxlsx_chart_chartarea_set_fill(chart, &fill);
 * @endcode
 *
 * @image html lxlsx_chart_chartarea.png
 *
 * For more information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_chartarea_set_line(lxlsx_chart *chart, lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for a chartarea.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param fill  A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of a chartarea:
 *
 * @code
 *     lxlsx_chart_chartarea_set_fill(chart, &fill);
 * @endcode
 *
 * See the example and image above.
 *
 * For more information see @ref lxlsx_chart_fills.
 */
void lxlsx_chart_chartarea_set_fill(lxlsx_chart *chart, lxlsx_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chartarea.
 *
 * @param chart   Pointer to a lxlsx_chart instance to be configured.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of a chartarea:
 *
 * @code
 *     lxlsx_chart_chartarea_set_pattern(series1, &pattern);
 * @endcode
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 */
void lxlsx_chart_chartarea_set_pattern(lxlsx_chart *chart,
                                 lxlsx_chart_pattern *pattern);

/**
 * @brief Set the line properties for a plotarea.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param line  A #lxlsx_chart_line struct.
 *
 * Set the line/border properties of a plotarea. In Excel the plotarea is
 * the area between the axes on which the chart series are plotted:
 *
 * @code
 *     lxlsx_chart_line line = {.color     = LXLSX_COLOR_RED,
 *                            .width     = 2,
 *                            .dash_type = LXLSX_CHART_LINE_DASH_DASH};
 *     lxlsx_chart_fill fill = {.color     = 0xFFFFC2};
 *
 *     lxlsx_chart_plotarea_set_line(chart, &line);
 *     lxlsx_chart_plotarea_set_fill(chart, &fill);
 *
 * @endcode
 *
 * @image html lxlsx_chart_plotarea.png
 *
 * For more information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_plotarea_set_line(lxlsx_chart *chart, lxlsx_chart_line *line);

/**
 * @brief Set the fill properties for a plotarea.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param fill  A #lxlsx_chart_fill struct.
 *
 * Set the fill properties of a plotarea:
 *
 * @code
 *     lxlsx_chart_plotarea_set_fill(chart, &fill);
 * @endcode
 *
 * See the example and image above.
 *
 * For more information see @ref lxlsx_chart_fills.
 */
void lxlsx_chart_plotarea_set_fill(lxlsx_chart *chart, lxlsx_chart_fill *fill);

/**
 * @brief Set the pattern properties for a plotarea.
 *
 * @param chart   Pointer to a lxlsx_chart instance to be configured.
 * @param pattern A #lxlsx_chart_pattern struct.
 *
 * Set the pattern properties of a plotarea:
 *
 * @code
 *     lxlsx_chart_plotarea_set_pattern(series1, &pattern);
 * @endcode
 *
 * For more information see #lxlsx_chart_pattern_type and @ref lxlsx_chart_patterns.
 */
void lxlsx_chart_plotarea_set_pattern(lxlsx_chart *chart, lxlsx_chart_pattern *pattern);

/**
 * @brief Set the manual layout of the chart plotarea.
 *
 * @param chart  Pointer to a lxlsx_chart instance to be configured.
 * @param layout A pointer to a chart #lxlsx_chart_layout struct.
 *
 * This function is used to simulate setting the manual position of the chart
 * plotarea. See @ref lxlsx_chart_layout for more information.
 */
void lxlsx_chart_plotarea_set_layout(lxlsx_chart *chart, lxlsx_chart_layout *layout);

/**
 * @brief Set the chart style type.
 *
 * @param chart    Pointer to a lxlsx_chart instance to be configured.
 * @param style_id An index representing the chart style, 1 - 48.
 *
 * The `%lxlsx_chart_set_style()` function is used to set the style of the chart to
 * one of the 48 built-in styles available on the "Design" tab in Excel 2007:
 *
 * @code
 *     lxlsx_chart_set_style(chart, 37)
 * @endcode
 *
 * @image html lxlsx_chart_style.png
 *
 * The style index number is counted from 1 on the top left in the Excel
 * dialog. The default style is 2.
 *
 * **Note:**
 *
 * In Excel 2013 the Styles section of the "Design" tab in Excel shows what
 * were referred to as "Layouts" in previous versions of Excel. These layouts
 * are not defined in the file format. They are a collection of modifications
 * to the base chart type. They can not be defined by the `lxlsx_chart_set_style()``
 * function.
 *
 */
void lxlsx_chart_set_style(lxlsx_chart *chart, uint8_t style_id);

/**
 * @brief Turn on a data table below the horizontal axis.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 *
 * The `%lxlsx_chart_set_table()` function adds a data table below the horizontal
 * axis with the data used to plot the chart:
 *
 * @code
 *     // Turn on the data table with default options.
 *     lxlsx_chart_set_table(chart);
 * @endcode
 *
 * @image html lxlsx_chart_data_table1.png
 *
 * The data table can only be shown with Bar, Column, Line and Area charts.
 *
 */
void lxlsx_chart_set_table(lxlsx_chart *chart);

/**
 * @brief Turn on/off grid options for a chart data table.
 *
 * @param chart       Pointer to a lxlsx_chart instance to be configured.
 * @param horizontal  Turn on/off the horizontal grid lines in the table.
 * @param vertical    Turn on/off the vertical grid lines in the table.
 * @param outline     Turn on/off the outline lines in the table.
 * @param legend_keys Turn on/off the legend keys in the table.
 *
 * The `%lxlsx_chart_set_table_grid()` function turns on/off grid options for a
 * chart data table. The data table grid options in Excel are shown in the
 * dialog below:
 *
 * @image html lxlsx_chart_data_table3.png
 *
 * These options can be passed to the `%lxlsx_chart_set_table_grid()` function.
 * The values for a default chart are:
 *
 * - `horizontal`: On.
 * - `vertical`: On.
 * - `outline`:  On.
 * - `legend_keys`: Off.
 *
 * Example:
 *
 * @code
 *     // Turn on the data table with default options.
 *     lxlsx_chart_set_table(chart);
 *
 *     // Turn on all grid lines and the grid legend.
 *     lxlsx_chart_set_table_grid(chart, LXLSX_TRUE, LXLSX_TRUE, LXLSX_TRUE, LXLSX_TRUE);
 *
 *     // Turn off the legend since it is show in the table.
 *     lxlsx_chart_legend_set_position(chart, LXLSX_CHART_LEGEND_NONE);
 *
 * @endcode
 *
 * @image html lxlsx_chart_data_table2.png
 *
 * The data table can only be shown with Bar, Column, Line and Area charts.
 *
 */
void lxlsx_chart_set_table_grid(lxlsx_chart *chart, uint8_t horizontal,
                          uint8_t vertical, uint8_t outline,
                          uint8_t legend_keys);

void lxlsx_chart_set_table_font(lxlsx_chart *chart, lxlsx_chart_font *font);

/**
 * @brief Turn on up-down bars for the chart.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 *
 * The `%lxlsx_chart_set_up_down_bars()` function adds Up-Down bars to Line charts
 * to indicate the difference between the first and last data series:
 *
 * @code
 *     lxlsx_chart_set_up_down_bars(chart);
 * @endcode
 *
 * @image html lxlsx_chart_data_tools4.png
 *
 * Up-Down bars are only available in Line charts. By default Up-Down bars are
 * black and white like in the above example. To format the border or fill
 * of the bars see the `lxlsx_chart_set_up_down_bars_format()` function below.
 */
void lxlsx_chart_set_up_down_bars(lxlsx_chart *chart);

/**
 * @brief Turn on up-down bars for the chart, with formatting.
 *
 * @param chart         Pointer to a lxlsx_chart instance to be configured.
 * @param up_bar_line   A #lxlsx_chart_line struct for the up-bar border.
 * @param up_bar_fill   A #lxlsx_chart_fill struct for the up-bar fill.
 * @param down_bar_line A #lxlsx_chart_line struct for the down-bar border.
 * @param down_bar_fill A #lxlsx_chart_fill struct for the down-bar fill.
 *
 * The `%lxlsx_chart_set_up_down_bars_format()` function adds Up-Down bars to Line
 * charts to indicate the difference between the first and last data series.
 * It also allows the up and down bars to be formatted:
 *
 * @code
 *     lxlsx_chart_line line      = {.color = LXLSX_COLOR_BLACK};
 *     lxlsx_chart_fill up_fill   = {.color = 0x00B050};
 *     lxlsx_chart_fill down_fill = {.color = LXLSX_COLOR_RED};
 *
 *     lxlsx_chart_set_up_down_bars_format(chart, &line, &up_fill, &line, &down_fill);
 * @endcode
 *
 * @image html lxlsx_chart_up_down_bars.png
 *
 * Up-Down bars are only available in Line charts.
 * For more format information  see @ref lxlsx_chart_lines and @ref lxlsx_chart_fills.
 */
void lxlsx_chart_set_up_down_bars_format(lxlsx_chart *chart,
                                   lxlsx_chart_line *up_bar_line,
                                   lxlsx_chart_fill *up_bar_fill,
                                   lxlsx_chart_line *down_bar_line,
                                   lxlsx_chart_fill *down_bar_fill);

/**
 * @brief Turn on and format Drop Lines for a chart.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param line  A #lxlsx_chart_line struct.
 *
 * The `%lxlsx_chart_set_drop_lines()` function adds Drop Lines to charts to
 * show the Category value of points in the data:
 *
 * @code
 *     lxlsx_chart_set_drop_lines(chart, NULL);
 * @endcode
 *
 * @image html lxlsx_chart_data_tools6.png
 *
 * It is possible to format the Drop Line line properties if required:
 *
 * @code
 *     lxlsx_chart_line line = {.color     = LXLSX_COLOR_RED,
 *                            .dash_type = LXLSX_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     lxlsx_chart_set_drop_lines(chart, &line);
 * @endcode
 *
 * Drop Lines are only available in Line and Area charts.
 * For more format information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_set_drop_lines(lxlsx_chart *chart, lxlsx_chart_line *line);

/**
 * @brief Turn on and format high-low Lines for a chart.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param line  A #lxlsx_chart_line struct.
 *
 * The `%lxlsx_chart_set_high_low_lines()` function adds High-Low Lines to charts
 * to show the Category value of points in the data:
 *
 * @code
 *     lxlsx_chart_set_high_low_lines(chart, NULL);
 * @endcode
 *
 * @image html lxlsx_chart_data_tools5.png
 *
 * It is possible to format the High-Low Line line properties if required:
 *
 * @code
 *     lxlsx_chart_line line = {.color     = LXLSX_COLOR_RED,
 *                            .dash_type = LXLSX_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     lxlsx_chart_set_high_low_lines(chart, &line);
 * @endcode
 *
 * High-Low Lines are only available in Line charts.
 * For more format information see @ref lxlsx_chart_lines.
 */
void lxlsx_chart_set_high_low_lines(lxlsx_chart *chart, lxlsx_chart_line *line);

/**
 * @brief Set the overlap between series in a Bar/Column chart.
 *
 * @param chart   Pointer to a lxlsx_chart instance to be configured.
 * @param overlap The overlap between the series. -100 to 100.
 *
 * The `%lxlsx_chart_set_series_overlap()` function sets the overlap between series
 * in Bar and Column charts.
 *
 * @code
 *     lxlsx_chart_set_series_overlap(chart, -50);
 * @endcode
 *
 * @image html lxlsx_chart_overlap.png
 *
 * The overlap value must be in the range `0 <= overlap <= 500`.
 * The default value is 0.
 *
 * This option is only available for Bar/Column charts.
 */
void lxlsx_chart_set_series_overlap(lxlsx_chart *chart, int8_t overlap);

/**
 * @brief Set the gap between series in a Bar/Column chart.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param gap   The gap between the series.  0 to 500.
 *
 * The `%lxlsx_chart_set_series_gap()` function sets the gap between series in
 * Bar and Column charts.
 *
 * @code
 *     lxlsx_chart_set_series_gap(chart, 400);
 * @endcode
 *
 * @image html lxlsx_chart_gap.png
 *
 * The gap value must be in the range `0 <= gap <= 500`. The default value
 * is 150.
 *
 * This option is only available for Bar/Column charts.
 */
void lxlsx_chart_set_series_gap(lxlsx_chart *chart, uint16_t gap);

/**
 * @brief Set the option for displaying blank data in a chart.
 *
 * @param chart    Pointer to a lxlsx_chart instance to be configured.
 * @param option The display option. A #lxlsx_chart_blank option.
 *
 * The `%lxlsx_chart_show_blanks_as()` function controls how blank data is displayed
 * in a chart:
 *
 * @code
 *     lxlsx_chart_show_blanks_as(chart, LXLSX_CHART_BLANKS_AS_CONNECTED);
 * @endcode
 *
 * The `option` parameter can have one of the following values:
 *
 * - #LXLSX_CHART_BLANKS_AS_GAP: Show empty chart cells as gaps in the data.
 *   This is the default option for Excel charts.
 * - #LXLSX_CHART_BLANKS_AS_ZERO: Show empty chart cells as zeros.
 * - #LXLSX_CHART_BLANKS_AS_CONNECTED: Show empty chart cells as connected.
 *   Only for charts with lines.
 */
void lxlsx_chart_show_blanks_as(lxlsx_chart *chart, uint8_t option);

/**
 * @brief Display data on charts from hidden rows or columns.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 *
 * Display data that is in hidden rows or columns on the chart:
 *
 * @code
 *     lxlsx_chart_show_hidden_data(chart);
 * @endcode
 */
void lxlsx_chart_show_hidden_data(lxlsx_chart *chart);

/**
 * @brief Set the Pie/Doughnut chart rotation.
 *
 * @param chart    Pointer to a lxlsx_chart instance to be configured.
 * @param rotation The angle of rotation.
 *
 * The `lxlsx_chart_set_rotation()` function is used to set the rotation of the
 * first segment of a Pie/Doughnut chart. This has the effect of rotating
 * the entire chart:
 *
 * @code
 *     lxlsx_chart_set_rotation(chart, 28);
 * @endcode
 *
 * The angle of rotation must be in the range `0 <= rotation <= 360`.
 *
 * This option is only available for Pie/Doughnut charts.
 *
 */
void lxlsx_chart_set_rotation(lxlsx_chart *chart, uint16_t rotation);

/**
 * @brief Set the Doughnut chart hole size.
 *
 * @param chart Pointer to a lxlsx_chart instance to be configured.
 * @param size  The hole size as a percentage.
 *
 * The `lxlsx_chart_set_hole_size()` function is used to set the hole size of a
 * Doughnut chart:
 *
 * @code
 *     lxlsx_chart_set_hole_size(chart, 33);
 * @endcode
 *
 * The hole size must be in the range `10 <= size <= 90`.
 *
 * This option is only available for Doughnut charts.
 *
 */
void lxlsx_chart_set_hole_size(lxlsx_chart *chart, uint8_t size);

lxlsx_error lxlsx_chart_add_data_cache(lxlsx_series_range *range, uint8_t *data,
                                   uint16_t rows, uint8_t cols, uint8_t col);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _chart_xml_declaration(lxlsx_chart *chart);
STATIC void _chart_write_legend(lxlsx_chart *chart);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_CHART_H__ */
