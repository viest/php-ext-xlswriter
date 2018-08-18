/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * chart - A libxlsxwriter library for creating Excel XLSX chart files.
 *
 */

/**
 * @page chart_page The Chart object
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
 * calling the `workbook_add_chart()` function from a Workbook object. For
 * example:
 *
 * @code
 *
 * #include "xlsxwriter.h"
 *
 * int main() {
 *
 *     lxw_workbook  *workbook  = new_workbook("chart.xlsx");
 *     lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
 *
 *     // User function to add data to worksheet, not shown here.
 *     write_worksheet_data(worksheet);
 *
 *     // Create a chart object.
 *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_COLUMN);
 *
 *     // In the simplest case we just add some value data series.
 *     // The NULL categories will default to 1 to 5 like in Excel.
 *     chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$5");
 *     chart_add_series(chart, NULL, "=Sheet1!$B$1:$B$5");
 *     chart_add_series(chart, NULL, "=Sheet1!$C$1:$C$5");
 *
 *     // Insert the chart into the worksheet
 *     worksheet_insert_chart(worksheet, CELL("B7"), chart);
 *
 *     return workbook_close(workbook);
 * }
 *
 * @endcode
 *
 * The chart in the worksheet will look like this:
 * @image html chart_simple.png
 *
 * The basic procedure for adding a chart to a worksheet is:
 *
 * 1. Create the chart with `workbook_add_chart()`.
 * 2. Add one or more data series to the chart which refers to data in the
 *    workbook using `chart_add_series()`.
 * 3. Configure the chart with the other available functions shown below.
 * 4. Insert the chart into a worksheet using `worksheet_insert_chart()`.
 *
 */

#ifndef __LXW_CHART_H__
#define __LXW_CHART_H__

#include <stdint.h>
#include <string.h>

#include "common.h"
#include "format.h"

STAILQ_HEAD(lxw_chart_series_list, lxw_chart_series);
STAILQ_HEAD(lxw_series_data_points, lxw_series_data_point);

#define LXW_CHART_NUM_FORMAT_LEN 128
#define LXW_CHART_DEFAULT_GAP 501

/**
 * @brief Available chart types.
 */
typedef enum lxw_chart_type {

    /** None. */
    LXW_CHART_NONE = 0,

    /** Area chart. */
    LXW_CHART_AREA,

    /** Area chart - stacked. */
    LXW_CHART_AREA_STACKED,

    /** Area chart - percentage stacked. */
    LXW_CHART_AREA_STACKED_PERCENT,

    /** Bar chart. */
    LXW_CHART_BAR,

    /** Bar chart - stacked. */
    LXW_CHART_BAR_STACKED,

    /** Bar chart - percentage stacked. */
    LXW_CHART_BAR_STACKED_PERCENT,

    /** Column chart. */
    LXW_CHART_COLUMN,

    /** Column chart - stacked. */
    LXW_CHART_COLUMN_STACKED,

    /** Column chart - percentage stacked. */
    LXW_CHART_COLUMN_STACKED_PERCENT,

    /** Doughnut chart. */
    LXW_CHART_DOUGHNUT,

    /** Line chart. */
    LXW_CHART_LINE,

    /** Pie chart. */
    LXW_CHART_PIE,

    /** Scatter chart. */
    LXW_CHART_SCATTER,

    /** Scatter chart - straight. */
    LXW_CHART_SCATTER_STRAIGHT,

    /** Scatter chart - straight with markers. */
    LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,

    /** Scatter chart - smooth. */
    LXW_CHART_SCATTER_SMOOTH,

    /** Scatter chart - smooth with markers. */
    LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,

    /** Radar chart. */
    LXW_CHART_RADAR,

    /** Radar chart - with markers. */
    LXW_CHART_RADAR_WITH_MARKERS,

    /** Radar chart - filled. */
    LXW_CHART_RADAR_FILLED
} lxw_chart_type;

/**
 * @brief Chart legend positions.
 */
typedef enum lxw_chart_legend_position {

    /** No chart legend. */
    LXW_CHART_LEGEND_NONE = 0,

    /** Chart legend positioned at right side. */
    LXW_CHART_LEGEND_RIGHT,

    /** Chart legend positioned at left side. */
    LXW_CHART_LEGEND_LEFT,

    /** Chart legend positioned at top. */
    LXW_CHART_LEGEND_TOP,

    /** Chart legend positioned at bottom. */
    LXW_CHART_LEGEND_BOTTOM,

    /** Chart legend overlaid at right side. */
    LXW_CHART_LEGEND_OVERLAY_RIGHT,

    /** Chart legend overlaid at left side. */
    LXW_CHART_LEGEND_OVERLAY_LEFT
} lxw_chart_legend_position;

/**
 * @brief Chart line dash types.
 *
 * The dash types are shown in the order that they appear in the Excel dialog.
 * See @ref chart_lines.
 */
typedef enum lxw_chart_line_dash_type {

    /** Solid. */
    LXW_CHART_LINE_DASH_SOLID = 0,

    /** Round Dot. */
    LXW_CHART_LINE_DASH_ROUND_DOT,

    /** Square Dot. */
    LXW_CHART_LINE_DASH_SQUARE_DOT,

    /** Dash. */
    LXW_CHART_LINE_DASH_DASH,

    /** Dash Dot. */
    LXW_CHART_LINE_DASH_DASH_DOT,

    /** Long Dash. */
    LXW_CHART_LINE_DASH_LONG_DASH,

    /** Long Dash Dot. */
    LXW_CHART_LINE_DASH_LONG_DASH_DOT,

    /** Long Dash Dot Dot. */
    LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT,

    /* These aren't available in the dialog but are used by Excel. */
    LXW_CHART_LINE_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
} lxw_chart_line_dash_type;

/**
 * @brief Chart marker types.
 */
typedef enum lxw_chart_marker_type {

    /** Automatic, series default, marker type. */
    LXW_CHART_MARKER_AUTOMATIC,

    /** No marker type. */
    LXW_CHART_MARKER_NONE,

    /** Square marker type. */
    LXW_CHART_MARKER_SQUARE,

    /** Diamond marker type. */
    LXW_CHART_MARKER_DIAMOND,

    /** Triangle marker type. */
    LXW_CHART_MARKER_TRIANGLE,

    /** X shape marker type. */
    LXW_CHART_MARKER_X,

    /** Star marker type. */
    LXW_CHART_MARKER_STAR,

    /** Short dash marker type. */
    LXW_CHART_MARKER_SHORT_DASH,

    /** Long dash marker type. */
    LXW_CHART_MARKER_LONG_DASH,

    /** Circle marker type. */
    LXW_CHART_MARKER_CIRCLE,

    /** Plus (+) marker type. */
    LXW_CHART_MARKER_PLUS
} lxw_chart_marker_type;

/**
 * @brief Chart pattern types.
 */
typedef enum lxw_chart_pattern_type {

    /** None pattern. */
    LXW_CHART_PATTERN_NONE,

    /** 5 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_5,

    /** 10 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_10,

    /** 20 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_20,

    /** 25 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_25,

    /** 30 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_30,

    /** 40 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_40,

    /** 50 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_50,

    /** 60 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_60,

    /** 70 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_70,

    /** 75 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_75,

    /** 80 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_80,

    /** 90 Percent pattern. */
    LXW_CHART_PATTERN_PERCENT_90,

    /** Light downward diagonal pattern. */
    LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,

    /** Light upward diagonal pattern. */
    LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,

    /** Dark downward diagonal pattern. */
    LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,

    /** Dark upward diagonal pattern. */
    LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL,

    /** Wide downward diagonal pattern. */
    LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,

    /** Wide upward diagonal pattern. */
    LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,

    /** Light vertical pattern. */
    LXW_CHART_PATTERN_LIGHT_VERTICAL,

    /** Light horizontal pattern. */
    LXW_CHART_PATTERN_LIGHT_HORIZONTAL,

    /** Narrow vertical pattern. */
    LXW_CHART_PATTERN_NARROW_VERTICAL,

    /** Narrow horizontal pattern. */
    LXW_CHART_PATTERN_NARROW_HORIZONTAL,

    /** Dark vertical pattern. */
    LXW_CHART_PATTERN_DARK_VERTICAL,

    /** Dark horizontal pattern. */
    LXW_CHART_PATTERN_DARK_HORIZONTAL,

    /** Dashed downward diagonal pattern. */
    LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,

    /** Dashed upward diagonal pattern. */
    LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,

    /** Dashed horizontal pattern. */
    LXW_CHART_PATTERN_DASHED_HORIZONTAL,

    /** Dashed vertical pattern. */
    LXW_CHART_PATTERN_DASHED_VERTICAL,

    /** Small confetti pattern. */
    LXW_CHART_PATTERN_SMALL_CONFETTI,

    /** Large confetti pattern. */
    LXW_CHART_PATTERN_LARGE_CONFETTI,

    /** Zigzag pattern. */
    LXW_CHART_PATTERN_ZIGZAG,

    /** Wave pattern. */
    LXW_CHART_PATTERN_WAVE,

    /** Diagonal brick pattern. */
    LXW_CHART_PATTERN_DIAGONAL_BRICK,

    /** Horizontal brick pattern. */
    LXW_CHART_PATTERN_HORIZONTAL_BRICK,

    /** Weave pattern. */
    LXW_CHART_PATTERN_WEAVE,

    /** Plaid pattern. */
    LXW_CHART_PATTERN_PLAID,

    /** Divot pattern. */
    LXW_CHART_PATTERN_DIVOT,

    /** Dotted grid pattern. */
    LXW_CHART_PATTERN_DOTTED_GRID,

    /** Dotted diamond pattern. */
    LXW_CHART_PATTERN_DOTTED_DIAMOND,

    /** Shingle pattern. */
    LXW_CHART_PATTERN_SHINGLE,

    /** Trellis pattern. */
    LXW_CHART_PATTERN_TRELLIS,

    /** Sphere pattern. */
    LXW_CHART_PATTERN_SPHERE,

    /** Small grid pattern. */
    LXW_CHART_PATTERN_SMALL_GRID,

    /** Large grid pattern. */
    LXW_CHART_PATTERN_LARGE_GRID,

    /** Small check pattern. */
    LXW_CHART_PATTERN_SMALL_CHECK,

    /** Large check pattern. */
    LXW_CHART_PATTERN_LARGE_CHECK,

    /** Outlined diamond pattern. */
    LXW_CHART_PATTERN_OUTLINED_DIAMOND,

    /** Solid diamond pattern. */
    LXW_CHART_PATTERN_SOLID_DIAMOND
} lxw_chart_pattern_type;

/**
 * @brief Chart data label positions.
 */
typedef enum lxw_chart_label_position {
    /** Series data label position: default position. */
    LXW_CHART_LABEL_POSITION_DEFAULT,

    /** Series data label position: center. */
    LXW_CHART_LABEL_POSITION_CENTER,

    /** Series data label position: right. */
    LXW_CHART_LABEL_POSITION_RIGHT,

    /** Series data label position: left. */
    LXW_CHART_LABEL_POSITION_LEFT,

    /** Series data label position: above. */
    LXW_CHART_LABEL_POSITION_ABOVE,

    /** Series data label position: below. */
    LXW_CHART_LABEL_POSITION_BELOW,

    /** Series data label position: inside base.  */
    LXW_CHART_LABEL_POSITION_INSIDE_BASE,

    /** Series data label position: inside end. */
    LXW_CHART_LABEL_POSITION_INSIDE_END,

    /** Series data label position: outside end. */
    LXW_CHART_LABEL_POSITION_OUTSIDE_END,

    /** Series data label position: best fit. */
    LXW_CHART_LABEL_POSITION_BEST_FIT
} lxw_chart_label_position;

/**
 * @brief Chart data label separator.
 */
typedef enum lxw_chart_label_separator {
    /** Series data label separator: comma (the default). */
    LXW_CHART_LABEL_SEPARATOR_COMMA,

    /** Series data label separator: semicolon. */
    LXW_CHART_LABEL_SEPARATOR_SEMICOLON,

    /** Series data label separator: period. */
    LXW_CHART_LABEL_SEPARATOR_PERIOD,

    /** Series data label separator: newline. */
    LXW_CHART_LABEL_SEPARATOR_NEWLINE,

    /** Series data label separator: space. */
    LXW_CHART_LABEL_SEPARATOR_SPACE
} lxw_chart_label_separator;

/**
 * @brief Chart axis types.
 */
typedef enum lxw_chart_axis_type {
    /** Chart X axis. */
    LXW_CHART_AXIS_TYPE_X,

    /** Chart Y axis. */
    LXW_CHART_AXIS_TYPE_Y
} lxw_chart_axis_type;

enum lxw_chart_subtype {

    LXW_CHART_SUBTYPE_NONE = 0,
    LXW_CHART_SUBTYPE_STACKED,
    LXW_CHART_SUBTYPE_STACKED_PERCENT
};

enum lxw_chart_grouping {
    LXW_GROUPING_CLUSTERED,
    LXW_GROUPING_STANDARD,
    LXW_GROUPING_PERCENTSTACKED,
    LXW_GROUPING_STACKED
};

/**
 * @brief Axis positions for category axes.
 */
typedef enum lxw_chart_axis_tick_position {

    LXW_CHART_AXIS_POSITION_DEFAULT,

    /** Position category axis on tick marks. */
    LXW_CHART_AXIS_POSITION_ON_TICK,

    /** Position category axis between tick marks. */
    LXW_CHART_AXIS_POSITION_BETWEEN
} lxw_chart_axis_tick_position;

/**
 * @brief Axis label positions.
 */
typedef enum lxw_chart_axis_label_position {

    /** Position the axis labels next to the axis. The default. */
    LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO,

    /** Position the axis labels at the top of the chart, for horizontal
     * axes, or to the right for vertical axes.*/
    LXW_CHART_AXIS_LABEL_POSITION_HIGH,

    /** Position the axis labels at the bottom of the chart, for horizontal
     * axes, or to the left for vertical axes.*/
    LXW_CHART_AXIS_LABEL_POSITION_LOW,

    /** Turn off the the axis labels. */
    LXW_CHART_AXIS_LABEL_POSITION_NONE
} lxw_chart_axis_label_position;

/**
 * @brief Display units for chart value axis.
 */
typedef enum lxw_chart_axis_display_unit {

    /** Axis display units: None. The default. */
    LXW_CHART_AXIS_UNITS_NONE,

    /** Axis display units: Hundreds. */
    LXW_CHART_AXIS_UNITS_HUNDREDS,

    /** Axis display units: Thousands. */
    LXW_CHART_AXIS_UNITS_THOUSANDS,

    /** Axis display units: Ten thousands. */
    LXW_CHART_AXIS_UNITS_TEN_THOUSANDS,

    /** Axis display units: Hundred thousands. */
    LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS,

    /** Axis display units: Millions. */
    LXW_CHART_AXIS_UNITS_MILLIONS,

    /** Axis display units: Ten millions. */
    LXW_CHART_AXIS_UNITS_TEN_MILLIONS,

    /** Axis display units: Hundred millions. */
    LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS,

    /** Axis display units: Billions. */
    LXW_CHART_AXIS_UNITS_BILLIONS,

    /** Axis display units: Trillions. */
    LXW_CHART_AXIS_UNITS_TRILLIONS
} lxw_chart_axis_display_unit;

/**
 * @brief Tick mark types for an axis.
 */
typedef enum lxw_chart_axis_tick_mark {

    /** Default tick mark for the chart axis. Usually outside. */
    LXW_CHART_AXIS_TICK_MARK_DEFAULT,

    /** No tick mark for the axis. */
    LXW_CHART_AXIS_TICK_MARK_NONE,

    /** Tick mark inside the axis only. */
    LXW_CHART_AXIS_TICK_MARK_INSIDE,

    /** Tick mark outside the axis only. */
    LXW_CHART_AXIS_TICK_MARK_OUTSIDE,

    /** Tick mark inside and outside the axis. */
    LXW_CHART_AXIS_TICK_MARK_CROSSING
} lxw_chart_tick_mark;

typedef struct lxw_series_range {
    char *formula;
    char *sheetname;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
    uint8_t ignore_cache;

    uint8_t has_string_cache;
    uint16_t num_data_points;
    struct lxw_series_data_points *data_cache;

} lxw_series_range;

typedef struct lxw_series_data_point {
    uint8_t is_string;
    double number;
    char *string;
    uint8_t no_data;

    STAILQ_ENTRY (lxw_series_data_point) list_pointers;

} lxw_series_data_point;

/**
 * @brief Struct to represent a chart line.
 *
 * See @ref chart_lines.
 */
typedef struct lxw_chart_line {

    /** The chart font color. See @ref working_with_colors. */
    lxw_color_t color;

    /** Turn off/hide line. Set to 0 or 1.*/
    uint8_t none;

    /** Width of the line in increments of 0.25. Default is 2.25. */
    float width;

    /** The line dash type. See #lxw_chart_line_dash_type. */
    uint8_t dash_type;

    /** Set the transparency of the line. 0 - 100. Default 0. */
    uint8_t transparency;

    /* Members for internal use only. */
    uint8_t has_color;

} lxw_chart_line;

/**
 * @brief Struct to represent a chart fill.
 *
 * See @ref chart_fills.
 */
typedef struct lxw_chart_fill {

    /** The chart font color. See @ref working_with_colors. */
    lxw_color_t color;

    /** Turn off/hide line. Set to 0 or 1.*/
    uint8_t none;

    /** Set the transparency of the fill. 0 - 100. Default 0. */
    uint8_t transparency;

    /* Members for internal use only. */
    uint8_t has_color;

} lxw_chart_fill;

/**
 * @brief Struct to represent a chart pattern.
 *
 * See @ref chart_patterns.
 */
typedef struct lxw_chart_pattern {

    /** The pattern foreground color. See @ref working_with_colors. */
    lxw_color_t fg_color;

    /** The pattern background color. See @ref working_with_colors. */
    lxw_color_t bg_color;

    /** The pattern type. See #lxw_chart_pattern_type. */
    uint8_t type;

    /* Members for internal use only. */
    uint8_t has_fg_color;
    uint8_t has_bg_color;

} lxw_chart_pattern;

/**
 * @brief Struct to represent a chart font.
 *
 * See @ref chart_fonts.
 */
typedef struct lxw_chart_font {

    /** The chart font name, such as "Arial" or "Calibri". */
    char *name;

    /** The chart font size. The default is 11. */
    double size;

    /** The chart font bold property. Set to 0 or 1. */
    uint8_t bold;

    /** The chart font italic property. Set to 0 or 1. */
    uint8_t italic;

    /** The chart font underline property. Set to 0 or 1. */
    uint8_t underline;

    /** The chart font rotation property. Range: -90 to 90. */
    int32_t rotation;

    /** The chart font color. See @ref working_with_colors. */
    lxw_color_t color;

    /** The chart font pitch family property. Rarely required. set to 0. */
    uint8_t pitch_family;

    /** The chart font character set property. Rarely required. set to 0. */
    uint8_t charset;

    /** The chart font baseline property. Rarely required. set to 0. */
    int8_t baseline;

    /* Members for internal use only. */

    uint8_t has_color;

} lxw_chart_font;

typedef struct lxw_chart_marker {

    uint8_t type;
    uint8_t size;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;

} lxw_chart_marker;

typedef struct lxw_chart_legend {

    lxw_chart_font *font;
    uint8_t position;

} lxw_chart_legend;

typedef struct lxw_chart_title {

    char *name;
    lxw_row_t row;
    lxw_col_t col;
    lxw_chart_font *font;
    uint8_t off;
    uint8_t is_horizontal;
    uint8_t ignore_cache;

    /* We use a range to hold the title formula properties even though it
     * will only have 1 point in order to re-use similar functions.*/
    lxw_series_range *range;

    struct lxw_series_data_point data_point;

} lxw_chart_title;

/**
 * @brief Struct to represent an Excel chart data point.
 *
 * The lxw_chart_point used to set the line, fill and pattern of one or more
 * points in a chart data series. See @ref chart_points.
 */
typedef struct lxw_chart_point {

    /** The line/border for the chart point. See @ref chart_lines. */
    lxw_chart_line *line;

    /** The fill for the chart point. See @ref chart_fills. */
    lxw_chart_fill *fill;

    /** The pattern for the chart point. See @ref chart_patterns.*/
    lxw_chart_pattern *pattern;

} lxw_chart_point;

/**
 * @brief Define how blank values are displayed in a chart.
 */
typedef enum lxw_chart_blank {

    /** Show empty chart cells as gaps in the data. The default. */
    LXW_CHART_BLANKS_AS_GAP,

    /** Show empty chart cells as zeros. */
    LXW_CHART_BLANKS_AS_ZERO,

    /** Show empty chart cells as connected. Only for charts with lines. */
    LXW_CHART_BLANKS_AS_CONNECTED
} lxw_chart_blank;

enum lxw_chart_position {
    LXW_CHART_AXIS_RIGHT,
    LXW_CHART_AXIS_LEFT,
    LXW_CHART_AXIS_TOP,
    LXW_CHART_AXIS_BOTTOM
};

/**
 * @brief Type/amount of data series error bar.
 */
typedef enum lxw_chart_error_bar_type {
    /** Error bar type: Standard error. */
    LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,

    /** Error bar type: Fixed value. */
    LXW_CHART_ERROR_BAR_TYPE_FIXED,

    /** Error bar type: Percentage. */
    LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,

    /** Error bar type: Standard deviation(s). */
    LXW_CHART_ERROR_BAR_TYPE_STD_DEV
} lxw_chart_error_bar_type;

/**
 * @brief Direction for a data series error bar.
 */
typedef enum lxw_chart_error_bar_direction {

    /** Error bar extends in both directions. The default. */
    LXW_CHART_ERROR_BAR_DIR_BOTH,

    /** Error bar extends in positive direction. */
    LXW_CHART_ERROR_BAR_DIR_PLUS,

    /** Error bar extends in negative direction. */
    LXW_CHART_ERROR_BAR_DIR_MINUS
} lxw_chart_error_bar_direction;

/**
 * @brief Direction for a data series error bar.
 */
typedef enum lxw_chart_error_bar_axis {
    /** X axis error bar. */
    LXW_CHART_ERROR_BAR_AXIS_X,

    /** Y axis error bar. */
    LXW_CHART_ERROR_BAR_AXIS_Y
} lxw_chart_error_bar_axis;

/**
 * @brief End cap styles for a data series error bar.
 */
typedef enum lxw_chart_error_bar_cap {
    /** Flat end cap. The default. */
    LXW_CHART_ERROR_BAR_END_CAP,

    /** No end cap. */
    LXW_CHART_ERROR_BAR_NO_CAP
} lxw_chart_error_bar_cap;

typedef struct lxw_series_error_bars {
    uint8_t type;
    uint8_t direction;
    uint8_t endcap;
    uint8_t has_value;
    uint8_t is_set;
    uint8_t is_x;
    uint8_t chart_group;
    double value;
    lxw_chart_line *line;

} lxw_series_error_bars;

/**
 * @brief Series trendline/regression types.
 */
typedef enum lxw_chart_trendline_type {
    /** Trendline type: Linear. */
    LXW_CHART_TRENDLINE_TYPE_LINEAR,

    /** Trendline type: Logarithm. */
    LXW_CHART_TRENDLINE_TYPE_LOG,

    /** Trendline type: Polynomial. */
    LXW_CHART_TRENDLINE_TYPE_POLY,

    /** Trendline type: Power. */
    LXW_CHART_TRENDLINE_TYPE_POWER,

    /** Trendline type: Exponential. */
    LXW_CHART_TRENDLINE_TYPE_EXP,

    /** Trendline type: Moving Average. */
    LXW_CHART_TRENDLINE_TYPE_AVERAGE
} lxw_chart_trendline_type;

/**
 * @brief Struct to represent an Excel chart data series.
 *
 * The lxw_chart_series is created using the chart_add_series function. It is
 * used in functions that modify a chart series but the members of the struct
 * aren't modified directly.
 */
typedef struct lxw_chart_series {

    lxw_series_range *categories;
    lxw_series_range *values;
    lxw_chart_title title;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;
    lxw_chart_marker *marker;
    lxw_chart_point *points;
    uint16_t point_count;

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
    lxw_chart_font *label_font;

    lxw_series_error_bars *x_error_bars;
    lxw_series_error_bars *y_error_bars;

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
    lxw_chart_line *trendline_line;
    double trendline_intercept;

    STAILQ_ENTRY (lxw_chart_series) list_pointers;

} lxw_chart_series;

/* Struct for major/minor axis gridlines. */
typedef struct lxw_chart_gridline {

    uint8_t visible;
    lxw_chart_line *line;

} lxw_chart_gridline;

/**
 * @brief Struct to represent an Excel chart axis.
 *
 * The lxw_chart_axis struct is used in functions that modify a chart axis
 * but the members of the struct aren't modified directly.
 */
typedef struct lxw_chart_axis {

    lxw_chart_title title;

    char *num_format;
    char *default_num_format;
    uint8_t source_linked;

    uint8_t major_tick_mark;
    uint8_t minor_tick_mark;
    uint8_t is_horizontal;

    lxw_chart_gridline major_gridlines;
    lxw_chart_gridline minor_gridlines;

    lxw_chart_font *num_font;
    lxw_chart_line *line;
    lxw_chart_fill *fill;
    lxw_chart_pattern *pattern;

    uint8_t is_category;
    uint8_t is_date;
    uint8_t is_value;
    uint8_t axis_position;
    uint8_t position_axis;
    uint8_t label_position;
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
    uint8_t crossing_max;
    double crossing;

} lxw_chart_axis;

/**
 * @brief Struct to represent an Excel chart.
 *
 * The members of the lxw_chart struct aren't modified directly. Instead
 * the chart properties are set by calling the functions shown in chart.h.
 */
typedef struct lxw_chart {

    FILE *file;

    uint8_t type;
    uint8_t subtype;
    uint16_t series_index;

    void (*write_chart_type) (struct lxw_chart *);
    void (*write_plot_area) (struct lxw_chart *);

    /**
     * A pointer to the chart x_axis object which can be used in functions
     * that configures the X axis.
     */
    lxw_chart_axis *x_axis;

    /**
     * A pointer to the chart y_axis object which can be used in functions
     * that configures the Y axis.
     */
    lxw_chart_axis *y_axis;

    lxw_chart_title title;

    uint32_t id;
    uint32_t axis_id_1;
    uint32_t axis_id_2;
    uint32_t axis_id_3;
    uint32_t axis_id_4;

    uint8_t in_use;
    uint8_t chart_group;
    uint8_t cat_has_num_fmt;

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

    lxw_chart_legend legend;
    int16_t *delete_series;
    uint16_t delete_series_count;
    lxw_chart_marker *default_marker;

    lxw_chart_line *chartarea_line;
    lxw_chart_fill *chartarea_fill;
    lxw_chart_pattern *chartarea_pattern;
    lxw_chart_line *plotarea_line;
    lxw_chart_fill *plotarea_fill;
    lxw_chart_pattern *plotarea_pattern;

    uint8_t has_drop_lines;
    lxw_chart_line *drop_lines_line;

    uint8_t has_high_low_lines;
    lxw_chart_line *high_low_lines_line;

    struct lxw_chart_series_list *series_list;

    uint8_t has_table;
    uint8_t has_table_vertical;
    uint8_t has_table_horizontal;
    uint8_t has_table_outline;
    uint8_t has_table_legend_keys;
    lxw_chart_font *table_font;

    uint8_t show_blanks_as;
    uint8_t show_hidden_data;

    uint8_t has_up_down_bars;
    lxw_chart_line *up_bar_line;
    lxw_chart_line *down_bar_line;
    lxw_chart_fill *up_bar_fill;
    lxw_chart_fill *down_bar_fill;

    uint8_t default_label_position;

    STAILQ_ENTRY (lxw_chart) ordered_list_pointers;
    STAILQ_ENTRY (lxw_chart) list_pointers;

} lxw_chart;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_chart *lxw_chart_new(uint8_t type);
void lxw_chart_free(lxw_chart *chart);
void lxw_chart_assemble_xml_file(lxw_chart *chart);

/**
 * @brief Add a data series to a chart.
 *
 * @param chart      Pointer to a lxw_chart instance to be configured.
 * @param categories The range of categories in the data series.
 * @param values     The range of values in the data series.
 *
 * @return A lxw_chart_series object pointer.
 *
 * In Excel a chart **series** is a collection of information that defines
 * which data is plotted such as the categories and values. It is also used to
 * define the formatting for the data.
 *
 * For an libxlsxwriter chart object the `%chart_add_series()` function is
 * used to set the categories and values of the series:
 *
 * @code
 *     chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
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
 *     chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
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
 * `chart_series_set_categories()` and `chart_series_set_values()` functions:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart, NULL, NULL);
 *
 *     // Configure the series using a syntax that is easier to define programmatically.
 *     chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
 *     chart_series_set_values(    series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
 * @endcode
 *
 * As shown in the previous example the return value from
 * `%chart_add_series()` is a lxw_chart_series pointer. This can be used in
 * other functions that configure a series.
 *
 *
 * More than one series can be added to a chart. The series numbering and
 * order in the Excel chart will be the same as the order in which they are
 * added in libxlsxwriter:
 *
 * @code
 *    chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
 *    chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5");
 *    chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5");
 * @endcode
 *
 * It is also possible to specify non-contiguous ranges:
 *
 * @code
 *    chart_add_series(
 *        chart,
 *        "=(Sheet1!$A$1:$A$9,Sheet1!$A$14:$A$25)",
 *        "=(Sheet1!$B$1:$B$9,Sheet1!$B$14:$B$25)"
 *    );
 * @endcode
 *
 */
lxw_chart_series *chart_add_series(lxw_chart *chart,
                                   const char *categories,
                                   const char *values);

/**
 * @brief Set a series "categories" range using row and column values.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `chart_add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%chart_series_set_categories()` function is an alternative method that
 * is easier to generate programmatically. It requires that you set the
 * `categories` and `values` parameters in `chart_add_series()`to `NULL` and
 * then set them using row and column values in
 * `chart_series_set_categories()` and `chart_series_set_values()`:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart, NULL, NULL);
 *
 *     // Configure the series ranges programmatically.
 *     chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
 *     chart_series_set_values(    series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
 * @endcode
 *
 */
void chart_series_set_categories(lxw_chart_series *series,
                                 const char *sheetname, lxw_row_t first_row,
                                 lxw_col_t first_col, lxw_row_t last_row,
                                 lxw_col_t last_col);

/**
 * @brief Set a series "values" range using row and column values.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the data range.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * The `categories` and `values` of a chart data series are generally set
 * using the `chart_add_series()` function and Excel range formulas like
 * `"=Sheet1!$A$2:$A$7"`.
 *
 * The `%chart_series_set_values()` function is an alternative method that is
 * easier to generate programmatically. See the documentation for
 * `chart_series_set_categories()` above.
 */
void chart_series_set_values(lxw_chart_series *series, const char *sheetname,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);

/**
 * @brief Set the name of a chart series range.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param name   The series name.
 *
 * The `%chart_series_set_name` function is used to set the name for a chart
 * data series. The series name in Excel is displayed in the chart legend and
 * in the formula bar. The name property is optional and if it isn't supplied
 * it will default to `Series 1..n`.
 *
 * The function applies to a #lxw_chart_series object created using
 * `chart_add_series()`:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     chart_series_set_name(series, "Quarterly budget data");
 * @endcode
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     chart_series_set_name(series, "=Sheet1!$B1$1");
 * @endcode
 *
 * See also the `chart_series_set_name_range()` function to see how to set the
 * name formula programmatically.
 */
void chart_series_set_name(lxw_chart_series *series, const char *name);

/**
 * @brief Set a series name formula using row and column values.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%chart_series_set_name_range()` function can be used to set a series
 * name range and is an alternative to using `chart_series_set_name()` and a
 * string formula:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart, NULL, "=Sheet1!$B$2:$B$7");
 *
 *     chart_series_set_name_range(series, "Sheet1", 0, 2); // "=Sheet1!$C$1"
 * @endcode
 */
void chart_series_set_name_range(lxw_chart_series *series,
                                 const char *sheetname, lxw_row_t row,
                                 lxw_col_t col);
/**
 * @brief Set the line properties for a chart series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param line   A #lxw_chart_line struct.
 *
 * Set the line/border properties of a chart series:
 *
 * @code
 *     lxw_chart_line line = {.color = LXW_COLOR_RED};
 *
 *     chart_series_set_line(series1, &line);
 *     chart_series_set_line(series2, &line);
 *     chart_series_set_line(series3, &line);
 * @endcode
 *
 * @image html chart_series_set_line.png
 *
 * For more information see @ref chart_lines.
 */
void chart_series_set_line(lxw_chart_series *series, lxw_chart_line *line);

/**
 * @brief Set the fill properties for a chart series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param fill   A #lxw_chart_fill struct.
 *
 * Set the fill properties of a chart series:
 *
 * @code
 *     lxw_chart_fill fill1 = {.color = LXW_COLOR_RED};
 *     lxw_chart_fill fill2 = {.color = LXW_COLOR_YELLOW};
 *     lxw_chart_fill fill3 = {.color = LXW_COLOR_GREEN};
 *
 *     chart_series_set_fill(series1, &fill1);
 *     chart_series_set_fill(series2, &fill2);
 *     chart_series_set_fill(series3, &fill3);
 * @endcode
 *
 * @image html chart_series_set_fill.png
 *
 * For more information see @ref chart_fills.
 */
void chart_series_set_fill(lxw_chart_series *series, lxw_chart_fill *fill);

/**
 * @brief Invert the fill color for negative series values.
 *
 * @param series  A series object created via `chart_add_series()`.
 *
 * Invert the fill color for negative values. Usually only applicable to
 * column and bar charts.
 *
 * @code
 *     chart_series_set_invert_if_negative(series);
 * @endcode
 *
 */
void chart_series_set_invert_if_negative(lxw_chart_series *series);

/**
 * @brief Set the pattern properties for a chart series.
 *
 * @param series  A series object created via `chart_add_series()`.
 * @param pattern A #lxw_chart_pattern struct.
 *
 * Set the pattern properties of a chart series:
 *
 * @code
 *     lxw_chart_pattern pattern1 = {.type = LXW_CHART_PATTERN_SHINGLE,
 *                                   .fg_color = 0x804000,
 *                                   .bg_color = 0XC68C53};
 *
 *     lxw_chart_pattern pattern2 = {.type = LXW_CHART_PATTERN_HORIZONTAL_BRICK,
 *                                   .fg_color = 0XB30000,
 *                                   .bg_color = 0XFF6666};
 *
 *     chart_series_set_pattern(series1, &pattern1);
 *     chart_series_set_pattern(series2, &pattern2);
 *
 * @endcode
 *
 * @image html chart_pattern.png
 *
 * For more information see #lxw_chart_pattern_type and @ref chart_patterns.
 */
void chart_series_set_pattern(lxw_chart_series *series,
                              lxw_chart_pattern *pattern);

/**
 * @brief Set the data marker type for a series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param type   The marker type, see #lxw_chart_marker_type.
 *
 * In Excel a chart marker is used to distinguish data points in a plotted
 * series. In general only Line and Scatter and Radar chart types use
 * markers. The libxlsxwriter chart types that can have markers are:
 *
 * - #LXW_CHART_LINE
 * - #LXW_CHART_SCATTER
 * - #LXW_CHART_SCATTER_STRAIGHT
 * - #LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
 * - #LXW_CHART_SCATTER_SMOOTH
 * - #LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
 * - #LXW_CHART_RADAR
 * - #LXW_CHART_RADAR_WITH_MARKERS
 *
 * The chart types with `MARKERS` in the name have markers with default colors
 * and shapes turned on by default but it is possible using the various
 * `chart_series_set_marker_xxx()` functions below to change these defaults. It
 * is also possible to turn on an off markers.
 *
 * The `%chart_series_set_marker_type()` function is used to specify the
 * type of the series marker:
 *
 * @code
 *     chart_series_set_marker_type(series, LXW_CHART_MARKER_DIAMOND);
 * @endcode
 *
 * @image html chart_marker1.png
 *
 * The available marker types defined by #lxw_chart_marker_type are:
 *
 * - #LXW_CHART_MARKER_AUTOMATIC
 * - #LXW_CHART_MARKER_NONE
 * - #LXW_CHART_MARKER_SQUARE
 * - #LXW_CHART_MARKER_DIAMOND
 * - #LXW_CHART_MARKER_TRIANGLE
 * - #LXW_CHART_MARKER_X
 * - #LXW_CHART_MARKER_STAR
 * - #LXW_CHART_MARKER_SHORT_DASH
 * - #LXW_CHART_MARKER_LONG_DASH
 * - #LXW_CHART_MARKER_CIRCLE
 * - #LXW_CHART_MARKER_PLUS
 *
 * The `#LXW_CHART_MARKER_NONE` type can be used to turn off default markers:
 *
 * @code
 *     chart_series_set_marker_type(series, LXW_CHART_MARKER_NONE);
 * @endcode
 *
 * @image html chart_series_set_marker_none.png
 *
 * The `#LXW_CHART_MARKER_AUTOMATIC` type is a special case which turns on a
 * marker using the default marker style for the particular series. If
 * automatic is on then other marker properties such as size, line or fill
 * cannot be set.
 */
void chart_series_set_marker_type(lxw_chart_series *series, uint8_t type);

/**
 * @brief Set the size of a data marker for a series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param size   The size of the marker.
 *
 * The `%chart_series_set_marker_size()` function is used to specify the
 * size of the series marker:
 *
 * @code
 *     chart_series_set_marker_type(series, LXW_CHART_MARKER_CIRCLE);
 *     chart_series_set_marker_size(series, 10);
 * @endcode
 *
 * @image html chart_series_set_marker_size.png
 *
 */
void chart_series_set_marker_size(lxw_chart_series *series, uint8_t size);

/**
 * @brief Set the line properties for a chart series marker.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param line   A #lxw_chart_line struct.
 *
 * Set the line/border properties of a chart marker:
 *
 * @code
 *     lxw_chart_line line = {.color = LXW_COLOR_BLACK};
 *     lxw_chart_fill fill = {.color = LXW_COLOR_RED};
 *
 *     chart_series_set_marker_type(series, LXW_CHART_MARKER_SQUARE);
 *     chart_series_set_marker_size(series, 8);
 *
 *     chart_series_set_marker_line(series, &line);
 *     chart_series_set_marker_fill(series, &fill);
 * @endcode
 *
 * @image html chart_marker2.png
 *
 * For more information see @ref chart_lines.
 */
void chart_series_set_marker_line(lxw_chart_series *series,
                                  lxw_chart_line *line);

/**
 * @brief Set the fill properties for a chart series marker.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param fill   A #lxw_chart_fill struct.
 *
 * Set the fill properties of a chart marker:
 *
 * @code
 *     chart_series_set_marker_fill(series, &fill);
 * @endcode
 *
 * See the example and image above and also see @ref chart_fills.
 */
void chart_series_set_marker_fill(lxw_chart_series *series,
                                  lxw_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chart series marker.
 *
 * @param series  A series object created via `chart_add_series()`.
 * @param pattern A #lxw_chart_pattern struct.
 *
 * Set the pattern properties of a chart marker:
 *
 * @code
 *     chart_series_set_marker_pattern(series, &pattern);
 * @endcode
 *
 * For more information see #lxw_chart_pattern_type and @ref chart_patterns.
 */
void chart_series_set_marker_pattern(lxw_chart_series *series,
                                     lxw_chart_pattern *pattern);

/**
 * @brief Set the formatting for points in the series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param points An NULL terminated array of #lxw_chart_point pointers.
 *
 * @return A #lxw_error.
 *
 * In general formatting is applied to an entire series in a chart. However,
 * it is occasionally required to format individual points in a series. In
 * particular this is required for Pie/Doughnut charts where each segment is
 * represented by a point.
 *
 * @dontinclude chart_pie_colors.c
 * @skip Add the data series
 * @until chart_series_set_points
 *
 * @image html chart_points1.png
 *
 * @note The array of #lxw_chart_point pointers should be NULL terminated
 * as shown in the example.
 *
 * For more details see @ref chart_points
 */
lxw_error chart_series_set_points(lxw_chart_series *series,
                                  lxw_chart_point *points[]);

/**
 * @brief Smooth a line or scatter chart series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param smooth Turn off/on the line smoothing. (0/1)
 *
 * The `chart_series_set_smooth()` function is used to set the smooth property
 * of a line series. It is only applicable to the line and scatter chart
 * types:
 *
 * @code
 *     chart_series_set_smooth(series2, LXW_TRUE);
 * @endcode
 *
 * @image html chart_smooth.png
 *
 *
 */
void chart_series_set_smooth(lxw_chart_series *series, uint8_t smooth);

/**
 * @brief Add data labels to a chart series.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_labels()` function is used to turn on data labels
 * for a chart series. Data labels indicate the values of the plotted data
 * points.
 *
 * @code
 *     chart_series_set_labels(series);
 * @endcode
 *
 * @image html chart_labels1.png
 *
 * By default data labels are displayed in Excel with only the values shown:
 *
 * @image html chart_labels2.png
 *
 * However, it is possible to configure other display options, as shown
 * in the functions below.
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels(lxw_chart_series *series);

/**
 * @brief Set the display options for the labels of a data series.
 *
 * @param series        A series object created via `chart_add_series()`.
 * @param show_name     Turn on/off the series name in the label caption.
 * @param show_category Turn on/off the category name in the label caption.
 * @param show_value    Turn on/off the value in the label caption.
 *
 * The `%chart_series_set_labels_options()` function is used to set the
 * parameters that are displayed in the series data label:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_options(series, LXW_TRUE, LXW_TRUE, LXW_TRUE);
 * @endcode
 *
 * @image html chart_labels3.png
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_options(lxw_chart_series *series,
                                     uint8_t show_name, uint8_t show_category,
                                     uint8_t show_value);

/**
 * @brief Set the separator for the data label captions.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param separator The separator for the data label options:
 *                  #lxw_chart_label_separator.
 *
 * The `%chart_series_set_labels_separator()` function is used to change the
 * separator between multiple data label items. The default options is a comma
 * separator as shown in the previous example.
 *
 * The available options are:
 *
 * - #LXW_CHART_LABEL_SEPARATOR_SEMICOLON: semicolon separator.
 * - #LXW_CHART_LABEL_SEPARATOR_PERIOD: a period (dot) separator.
 * - #LXW_CHART_LABEL_SEPARATOR_NEWLINE: a newline separator.
 * - #LXW_CHART_LABEL_SEPARATOR_SPACE: a space separator.
 *
 * For example:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_options(series, LXW_TRUE, LXW_TRUE, LXW_TRUE);
 *     chart_series_set_labels_separator(series, LXW_CHART_LABEL_SEPARATOR_NEWLINE);
 * @endcode
 *
 * @image html chart_labels4.png
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_separator(lxw_chart_series *series,
                                       uint8_t separator);

/**
 * @brief Set the data label position for a series.
 *
 * @param series   A series object created via `chart_add_series()`.
 * @param position The data label position: #lxw_chart_label_position.
 *
 * The `%chart_series_set_labels_position()` function sets the position of
 * the labels in the data series:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_position(series, LXW_CHART_LABEL_POSITION_ABOVE);
 * @endcode
 *
 * @image html chart_labels5.png
 *
 * In Excel the allowable data label positions vary for different chart
 * types. The allowable, and default, positions are:
 *
 * | Position                              | Line, Scatter | Bar, Column   | Pie, Doughnut | Area, Radar   |
 * | :------------------------------------ | :------------ | :------------ | :------------ | :------------ |
 * | #LXW_CHART_LABEL_POSITION_CENTER      | Yes           | Yes           | Yes           | Yes (default) |
 * | #LXW_CHART_LABEL_POSITION_RIGHT       | Yes (default) |               |               |               |
 * | #LXW_CHART_LABEL_POSITION_LEFT        | Yes           |               |               |               |
 * | #LXW_CHART_LABEL_POSITION_ABOVE       | Yes           |               |               |               |
 * | #LXW_CHART_LABEL_POSITION_BELOW       | Yes           |               |               |               |
 * | #LXW_CHART_LABEL_POSITION_INSIDE_BASE |               | Yes           |               |               |
 * | #LXW_CHART_LABEL_POSITION_INSIDE_END  |               | Yes           | Yes           |               |
 * | #LXW_CHART_LABEL_POSITION_OUTSIDE_END |               | Yes (default) | Yes           |               |
 * | #LXW_CHART_LABEL_POSITION_BEST_FIT    |               |               | Yes (default) |               |
 *
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_position(lxw_chart_series *series,
                                      uint8_t position);

/**
 * @brief Set leader lines for Pie and Doughnut charts.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_labels_leader_line()` function  is used to turn on
 * leader lines for the data label of a series. It is mainly used for pie
 * or doughnut charts:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_leader_line(series);
 * @endcode
 *
 * @note Even when leader lines are turned on they aren't automatically
 *       visible in Excel or XlsxWriter. Due to an Excel limitation
 *       (or design) leader lines only appear if the data label is moved
 *       manually or if the data labels are very close and need to be
 *       adjusted automatically.
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_leader_line(lxw_chart_series *series);

/**
 * @brief Set the legend key for a data label in a chart series.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_labels_legend()` function is used to set the
 * legend key for a data series:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_legend(series);
 * @endcode
 *
 * @image html chart_labels6.png
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_legend(lxw_chart_series *series);

/**
 * @brief Set the percentage for a Pie/Doughnut data point.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_labels_percentage()` function is used to turn on
 * the display of data labels as a percentage for a series. It is mainly
 * used for pie charts:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_options(series, LXW_FALSE, LXW_FALSE, LXW_FALSE);
 *     chart_series_set_labels_percentage(series);
 * @endcode
 *
 * @image html chart_labels7.png
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_percentage(lxw_chart_series *series);

/**
 * @brief Set the number format for chart data labels in a series.
 *
 * @param series     A series object created via `chart_add_series()`.
 * @param num_format The number format string.
 *
 * The `%chart_series_set_labels_num_format()` function is used to set the
 * number format for data labels:
 *
 * @code
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_num_format(series, "$0.00");
 * @endcode
 *
 * @image html chart_labels8.png
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `format_set_num_format()`.
 *
 * For more information see @ref chart_labels.
 */
void chart_series_set_labels_num_format(lxw_chart_series *series,
                                        const char *num_format);

/**
 * @brief Set the font properties for chart data labels in a series
 *
 * @param series A series object created via `chart_add_series()`.
 * @param font   A pointer to a chart #lxw_chart_font font struct.
 *
 *
 * The `%chart_series_set_labels_font()` function is used to set the font
 * for data labels:
 *
 * @code
 *     lxw_chart_font font = {.name = "Consolas", .color = LXW_COLOR_RED};
 *
 *     chart_series_set_labels(series);
 *     chart_series_set_labels_font(series, &font);
 * @endcode
 *
 * @image html chart_labels9.png
 *
 * For more information see @ref chart_fonts and @ref chart_labels.
 *
 */
void chart_series_set_labels_font(lxw_chart_series *series,
                                  lxw_chart_font *font);

/**
 * @brief Turn on a trendline for a chart data series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param type   The type of trendline: #lxw_chart_trendline_type.
 * @param value  The order/period value for polynomial and moving average
 *               trendlines.
 *
 * A trendline can be added to a chart series to indicate trends in the data
 * such as a moving average or a polynomial fit. The trendlines types are
 * shown in the following Excel dialog:
 *
 * @image html chart_trendline0.png
 *
 * The `%chart_series_set_trendline()` function turns on these trendlines for
 * a data series:
 *
 * @code
 *     chart = workbook_add_chart(workbook, LXW_CHART_LINE);
 *     series = chart_add_series(chart, NULL, "Sheet1!$A$1:$A$6");
 *
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 * @endcode
 *
 * @image html chart_trendline2.png
 *
 * The `value` parameter corresponds to *order* for a polynomial trendline
 * and *period* for a Moving Average trendline. It both cases it must be >= 2.
 * The `value` parameter  is ignored for all other trendlines:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_AVERAGE, 2);
 * @endcode
 *
 * @image html chart_trendline3.png
 *
 * The allowable values for the the trendline `type` are:
 *
 * - #LXW_CHART_TRENDLINE_TYPE_LINEAR: Linear trendline.
 * - #LXW_CHART_TRENDLINE_TYPE_LOG: Logarithm trendline.
 * - #LXW_CHART_TRENDLINE_TYPE_POLY: Polynomial trendline. The `value`
 *   parameter corresponds to *order*.
 * - #LXW_CHART_TRENDLINE_TYPE_POWER: Power trendline.
 * - #LXW_CHART_TRENDLINE_TYPE_EXP: Exponential trendline.
 * - #LXW_CHART_TRENDLINE_TYPE_AVERAGE: Moving Average trendline. The `value`
 *   parameter corresponds to *period*.
 *
 * Other trendline options, such as those shown in the following Excel
 * dialog, can be set using the functions below.
 *
 * @image html chart_trendline1.png
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline(lxw_chart_series *series, uint8_t type,
                                uint8_t value);

/**
 * @brief Set the trendline forecast for a chart data series.
 *
 * @param series   A series object created via `chart_add_series()`.
 * @param forward  The forward period.
 * @param backward The backwards period.
 *
 * The `%chart_series_set_trendline_forecast()` function sets the forward
 * and backward forecast periods for the trendline:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_forecast(series, 0.5, 0.5);
 * @endcode
 *
 * @image html chart_trendline4.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline_forecast(lxw_chart_series *series,
                                         double forward, double backward);

/**
 * @brief Display the equation of a trendline for a chart data series.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_trendline_equation()` function displays the
 * equation of the trendline on the chart:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_equation(series);
 * @endcode
 *
 * @image html chart_trendline5.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline_equation(lxw_chart_series *series);

/**
 * @brief Display the R squared value of a trendline for a chart data series.
 *
 * @param series A series object created via `chart_add_series()`.
 *
 * The `%chart_series_set_trendline_r_squared()` function displays the
 * R-squared value for the trendline on the chart:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_r_squared(series);
 * @endcode
 *
 * @image html chart_trendline6.png
 *
 * @note This feature isn't available for Moving Average in Excel.
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline_r_squared(lxw_chart_series *series);

/**
 * @brief Set the trendline Y-axis intercept for a chart data series.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param intercept Y-axis intercept value.
 *
 * The `%chart_series_set_trendline_intercept()` function sets the Y-axis
 * intercept for the trendline:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_equation(series);
 *     chart_series_set_trendline_intercept(series, 0.8);
 * @endcode
 *
 * @image html chart_trendline7.png
 *
 * As can be seen from the equation on the chart the intercept point
 * (when X=0) is the same as the value set in the equation.
 *
 * @note The intercept feature is only available in Excel for Exponential,
 *       Linear and Polynomial trendline types.
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline_intercept(lxw_chart_series *series,
                                          double intercept);

/**
 * @brief Set the trendline name for a chart data series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param name   The name of the trendline to display in the legend.
 *
 * The `%chart_series_set_trendline_name()` function sets the name of the
 * trendline that is displayed in the chart legend. In the examples above
 * the trendlines are displayed with default names like "Linear (Series 1)"
 * and "2 per Mov. Avg. (Series 1)". If these names are too verbose or not
 * descriptive enough you can set your own trendline name:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_name(series, "My trendline");
 * @endcode
 *
 * @image html chart_trendline8.png
 *
 * It is often preferable to turn off the trendline caption in the legend.
 * This is down in Excel by deleting the trendline name from the legend.
 * In libxlsxwriter this is done using the `chart_legend_delete_series()`
 * function to delete the zero based series numbers:
 *
 * @code
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *
 *     // Delete the series name for the second series (=1 in zero base).
 *     // The -1 value indicates the end of the array of values.
 *     int16_t names[] = {1, -1};
 *     chart_legend_delete_series(chart, names);
 * @endcode
 *
 * @image html chart_trendline9.png
 *
 * For more information see @ref chart_trendlines.
 */
void chart_series_set_trendline_name(lxw_chart_series *series,
                                     const char *name);

/**
 * @brief Set the trendline line properties for a chart data series.
 *
 * @param series A series object created via `chart_add_series()`.
 * @param line   A #lxw_chart_line struct.
 *
 * The `%chart_series_set_trendline_line()` function is used to set the line
 * properties of a trendline:
 *
 * @code
 *     lxw_chart_line line = {.color     = LXW_COLOR_RED,
 *                            .dash_type = LXW_CHART_LINE_DASH_LONG_DASH};
 *
 *     chart_series_set_trendline(series, LXW_CHART_TRENDLINE_TYPE_LINEAR, 0);
 *     chart_series_set_trendline_line(series, &line);
 * @endcode
 *
 * @image html chart_trendline10.png
 *
 * For more information see @ref chart_trendlines and @ref chart_lines.
 */
void chart_series_set_trendline_line(lxw_chart_series *series,
                                     lxw_chart_line *line);
/**
 * @brief           Get a pointer to X or Y error bars from a chart series.
 *
 * @param series    A series object created via `chart_add_series()`.
 * @param axis_type The axis type (X or Y): #lxw_chart_error_bar_axis.
 *
 * The `%chart_series_get_error_bars()` function returns a pointer to the
 * error bars of a series based on the type of #lxw_chart_error_bar_axis:
 *
 * @code
 *     lxw_series_error_bars *x_error_bars;
 *     lxw_series_error_bars *y_error_bars;
 *
 *     x_error_bars = chart_series_get_error_bars(series, LXW_CHART_ERROR_BAR_AXIS_X);
 *     y_error_bars = chart_series_get_error_bars(series, LXW_CHART_ERROR_BAR_AXIS_Y);
 *
 *     // Use the error bar pointers.
 *     chart_series_set_error_bars(x_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 *
 *     chart_series_set_error_bars(y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * Note, the series error bars can also be accessed directly:
 *
 * @code
 *     // Equivalent to the above example, without function calls.
 *     chart_series_set_error_bars(series->x_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 *
 *     chart_series_set_error_bars(series->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * @return Pointer to the series error bars, or NULL if not found.
 */

lxw_series_error_bars *chart_series_get_error_bars(lxw_chart_series *series, lxw_chart_error_bar_axis
                                                   axis_type);

/**
 * Set the X or Y error bars for a chart series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param type       The type of error bar: #lxw_chart_error_bar_type.
 * @param value      The error value.
 *
 * Error bars can be added to a chart series to indicate error bounds in the
 * data. The error bars can be vertical `y_error_bars` (the most common type)
 * or horizontal `x_error_bars` (for Bar and Scatter charts only).
 *
 * @image html chart_error_bars0.png
 *
 * The `%chart_series_set_error_bars()` function sets the error bar type
 * and value associated with the type:
 *
 * @code
 *     lxw_chart_series *series = chart_add_series(chart,
 *                                                 "=Sheet1!$A$1:$A$5",
 *                                                 "=Sheet1!$B$1:$B$5");
 *
 *     chart_series_set_error_bars(series->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 * @endcode
 *
 * @image html chart_error_bars1.png
 *
 * The error bar types that be used are:
 *
 * - #LXW_CHART_ERROR_BAR_TYPE_STD_ERROR: Standard error.
 * - #LXW_CHART_ERROR_BAR_TYPE_FIXED: Fixed value.
 * - #LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE: Percentage.
 * - #LXW_CHART_ERROR_BAR_TYPE_STD_DEV: Standard deviation(s).
 *
 * @note Custom error bars are not currently supported.
 *
 * All error bar types, apart from Standard error, should have a valid
 * value to set the error range:
 *
 * @code
 *     chart_series_set_error_bars(series1->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_FIXED, 2);
 *
 *     chart_series_set_error_bars(series2->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE, 5);
 *
 *     chart_series_set_error_bars(series3->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_DEV, 1);
 * @endcode
 *
 * For the Standard error type the value is ignored.
 *
 * For more information see @ref chart_error_bars.
 */
void chart_series_set_error_bars(lxw_series_error_bars *error_bars,
                                 uint8_t type, double value);

/**
 * @brief Set the direction (up, down or both) of the error bars for a chart
 *        series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param direction  The bar direction: #lxw_chart_error_bar_direction.
 *
 * The `%chart_series_set_error_bars_direction()` function sets the
 * direction of the error bars:
 *
 * @code
 *     chart_series_set_error_bars(series->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     chart_series_set_error_bars_direction(series->y_error_bars,
 *                                           LXW_CHART_ERROR_BAR_DIR_PLUS);
 * @endcode
 *
 * @image html chart_error_bars2.png
 *
 * The valid directions are:
 *
 * - #LXW_CHART_ERROR_BAR_DIR_BOTH: Error bar extends in both directions.
 *   The default.
 * - #LXW_CHART_ERROR_BAR_DIR_PLUS: Error bar extends in positive direction.
 * - #LXW_CHART_ERROR_BAR_DIR_MINUS: Error bar extends in negative direction.
 *
 * For more information see @ref chart_error_bars.
 */
void chart_series_set_error_bars_direction(lxw_series_error_bars *error_bars,
                                           uint8_t direction);

/**
 * @brief Set the end cap type for the error bars of a chart series.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param endcap     The error bar end cap type: #lxw_chart_error_bar_cap .
 *
 * The `%chart_series_set_error_bars_endcap()` function sets the end cap
 * type for the error bars:
 *
 * @code
 *     chart_series_set_error_bars(series->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     chart_series_set_error_bars_endcap(series->y_error_bars,
                                          LXW_CHART_ERROR_BAR_NO_CAP);
 * @endcode
 *
 * @image html chart_error_bars3.png
 *
 * The valid values are:
 *
 * - #LXW_CHART_ERROR_BAR_END_CAP: Flat end cap. The default.
 * - #LXW_CHART_ERROR_BAR_NO_CAP: No end cap.
 *
 * For more information see @ref chart_error_bars.
 */
void chart_series_set_error_bars_endcap(lxw_series_error_bars *error_bars,
                                        uint8_t endcap);

/**
 * @brief Set the line properties for a chart series error bars.
 *
 * @param error_bars A pointer to the series X or Y error bars.
 * @param line       A #lxw_chart_line struct.
 *
 * The `%chart_series_set_error_bars_line()` function sets the line
 * properties for the error bars:
 *
 * @code
 *     lxw_chart_line line = {.color     = LXW_COLOR_RED,
 *                            .dash_type = LXW_CHART_LINE_DASH_ROUND_DOT};
 *
 *     chart_series_set_error_bars(series->y_error_bars,
 *                                 LXW_CHART_ERROR_BAR_TYPE_STD_ERROR, 0);
 *
 *     chart_series_set_error_bars_line(series->y_error_bars, &line);
 * @endcode
 *
 * @image html chart_error_bars4.png
 *
 * For more information see @ref chart_lines and @ref chart_error_bars.
 */
void chart_series_set_error_bars_line(lxw_series_error_bars *error_bars,
                                      lxw_chart_line *line);

/**
 * @brief           Get an axis pointer from a chart.
 *
 * @param chart     Pointer to a lxw_chart instance to be configured.
 * @param axis_type The axis type (X or Y): #lxw_chart_axis_type.
 *
 * The `%chart_axis_get()` function returns a pointer to a chart axis based
 * on the  #lxw_chart_axis_type:
 *
 * @code
 *     lxw_chart_axis *x_axis = chart_axis_get(chart, LXW_CHART_AXIS_TYPE_X);
 *     lxw_chart_axis *y_axis = chart_axis_get(chart, LXW_CHART_AXIS_TYPE_Y);
 *
 *     // Use the axis pointer in other functions.
 *     chart_axis_major_gridlines_set_visible(x_axis, LXW_TRUE);
 *     chart_axis_major_gridlines_set_visible(y_axis, LXW_TRUE);
 * @endcode
 *
 * Note, the axis pointer can also be accessed directly:
 *
 * @code
 *     // Equivalent to the above example, without function calls.
 *     chart_axis_major_gridlines_set_visible(chart->x_axis, LXW_TRUE);
 *     chart_axis_major_gridlines_set_visible(chart->y_axis, LXW_TRUE);
 * @endcode
 *
 * @return Pointer to the chart axis, or NULL if not found.
 */
lxw_chart_axis *chart_axis_get(lxw_chart *chart,
                               lxw_chart_axis_type axis_type);

/**
 * @brief Set the name caption of the an axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param name The name caption of the axis.
 *
 * The `%chart_axis_set_name()` function sets the name (also known as title or
 * caption) for an axis. It can be used for the X or Y axes. The name is
 * displayed below an X axis and to the side of a Y axis.
 *
 * @code
 *     chart_axis_set_name(chart->x_axis, "Earnings per Quarter");
 *     chart_axis_set_name(chart->y_axis, "US Dollars (Millions)");
 * @endcode
 *
 * @image html chart_axis_set_name.png
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     chart_axis_set_name(chart->x_axis, "=Sheet1!$B1$1");
 * @endcode
 *
 * See also the `chart_axis_set_name_range()` function to see how to set the
 * name formula programmatically.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name(lxw_chart_axis *axis, const char *name);

/**
 * @brief Set a chart axis name formula using row and column values.
 *
 * @param axis      A pointer to a chart #lxw_chart_axis object.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%chart_axis_set_name_range()` function can be used to set an axis name
 * range and is an alternative to using `chart_axis_set_name()` and a string
 * formula:
 *
 * @code
 *     chart_axis_set_name_range(chart->x_axis, "Sheet1", 1, 0);
 *     chart_axis_set_name_range(chart->y_axis, "Sheet1", 2, 0);
 * @endcode
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name_range(lxw_chart_axis *axis, const char *sheetname,
                               lxw_row_t row, lxw_col_t col);

/**
 * @brief Set the font properties for a chart axis name.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param font A pointer to a chart #lxw_chart_font font struct.
 *
 * The `%chart_axis_set_name_font()` function is used to set the font of an
 * axis name:
 *
 * @code
 *     lxw_chart_font font = {.bold = LXW_TRUE, .color = LXW_COLOR_BLUE};
 *
 *     chart_axis_set_name(chart->x_axis, "Yearly data");
 *     chart_axis_set_name_font(chart->x_axis, &font);
 * @endcode
 *
 * @image html chart_axis_set_name_font.png
 *
 * For more information see @ref chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_name_font(lxw_chart_axis *axis, lxw_chart_font *font);

/**
 * @brief Set the font properties for the numbers of a chart axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param font A pointer to a chart #lxw_chart_font font struct.
 *
 * The `%chart_axis_set_num_font()` function is used to set the font of the
 * numbers on an axis:
 *
 * @code
 *     lxw_chart_font font = {.bold = LXW_TRUE, .color = LXW_COLOR_BLUE};
 *
 *     chart_axis_set_num_font(chart->x_axis, &font1);
 * @endcode
 *
 * @image html chart_axis_set_num_font.png
 *
 * For more information see @ref chart_fonts.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_num_font(lxw_chart_axis *axis, lxw_chart_font *font);

/**
 * @brief Set the number format for a chart axis.
 *
 * @param axis       A pointer to a chart #lxw_chart_axis object.
 * @param num_format The number format string.
 *
 * The `%chart_axis_set_num_format()` function is used to set the format of
 * the numbers on an axis:
 *
 * @code
 *     chart_axis_set_num_format(chart->x_axis, "0.00%");
 *     chart_axis_set_num_format(chart->y_axis, "$#,##0.00");
 * @endcode
 *
 * The number format is similar to the Worksheet Cell Format num_format,
 * see `format_set_num_format()`.
 *
 * @image html chart_axis_num_format.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_num_format(lxw_chart_axis *axis, const char *num_format);

/**
 * @brief Set the line properties for a chart axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param line A #lxw_chart_line struct.
 *
 * Set the line properties of a chart axis:
 *
 * @code
 *     // Hide the Y axis.
 *     lxw_chart_line line = {.none = LXW_TRUE};
 *
 *     chart_axis_set_line(chart->y_axis, &line);
 * @endcode
 *
 * @image html chart_axis_set_line.png
 *
 * For more information see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_line(lxw_chart_axis *axis, lxw_chart_line *line);

/**
 * @brief Set the fill properties for a chart axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param fill A #lxw_chart_fill struct.
 *
 * Set the fill properties of a chart axis:
 *
 * @code
 *     lxw_chart_fill fill = {.color = LXW_COLOR_YELLOW};
 *
 *     chart_axis_set_fill(chart->y_axis, &fill);
 * @endcode
 *
 * @image html chart_axis_set_fill.png
 *
 * For more information see @ref chart_fills.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_fill(lxw_chart_axis *axis, lxw_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chart axis.
 *
 * @param axis    A pointer to a chart #lxw_chart_axis object.
 * @param pattern A #lxw_chart_pattern struct.
 *
 * Set the pattern properties of a chart axis:
 *
 * @code
 *     chart_axis_set_pattern(chart->y_axis, &pattern);
 * @endcode
 *
 * For more information see #lxw_chart_pattern_type and @ref chart_patterns.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_pattern(lxw_chart_axis *axis, lxw_chart_pattern *pattern);

/**
 * @brief Reverse the order of the axis categories or values.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 *
 * Reverse the order of the axis categories or values:
 *
 * @code
 *     chart_axis_set_reverse(chart->x_axis);
 * @endcode
 *
 * @image html chart_reverse.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_reverse(lxw_chart_axis *axis);

/**
 * @brief Set the position that the axis will cross the opposite axis.
 *
 * @param axis  A pointer to a chart #lxw_chart_axis object.
 * @param value The category or value that the axis crosses at.
 *
 * Set the position that the axis will cross the opposite axis:
 *
 * @code
 *     chart_axis_set_crossing(chart->x_axis, 3);
 *     chart_axis_set_crossing(chart->y_axis, 8);
 * @endcode
 *
 * @image html chart_crossing1.png
 *
 * If crossing is omitted (the default) the crossing will be set automatically
 * by Excel based on the chart data.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_crossing(lxw_chart_axis *axis, double value);

/**
 * @brief Set the opposite axis crossing position as the axis maximum.
 *
 * @param axis  A pointer to a chart #lxw_chart_axis object.
 *
 * Set the position that the opposite axis will cross as the axis maximum.
 * The default axis crossing position is generally the axis minimum so this
 * function can be used to reverse the location of the axes without reversing
 * the number sequence:
 *
 * @code
 *     chart_axis_set_crossing_max(chart->x_axis);
 *     chart_axis_set_crossing_max(chart->y_axis);
 * @endcode
 *
 * @image html chart_crossing2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_crossing_max(lxw_chart_axis *axis);

/**
 * @brief Turn off/hide an axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 *
 * Turn off, hide, a chart axis:
 *
 * @code
 *     chart_axis_off(chart->x_axis);
 * @endcode
 *
 * @image html chart_axis_off.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_off(lxw_chart_axis *axis);

/**
 * @brief Position a category axis on or between the axis tick marks.
 *
 * @param axis     A pointer to a chart #lxw_chart_axis object.
 * @param position A #lxw_chart_axis_tick_position value.
 *
 * Position a category axis horizontally on, or between, the axis tick marks.
 *
 * There are two allowable values:
 *
 * - #LXW_CHART_AXIS_POSITION_ON_TICK
 * - #LXW_CHART_AXIS_POSITION_BETWEEN
 *
 * @code
 *     chart_axis_set_position(chart->x_axis, LXW_CHART_AXIS_POSITION_BETWEEN);
 * @endcode
 *
 * @image html chart_axis_set_position.png
 *
 * **Axis types**: This function is applicable to category axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_position(lxw_chart_axis *axis, uint8_t position);

/**
 * @brief Position the axis labels.
 *
 * @param axis     A pointer to a chart #lxw_chart_axis object.
 * @param position A #lxw_chart_axis_label_position value.
 *
 * Position the axis labels for the chart. The labels are the numbers, or
 * strings or dates, on the axis that indicate the categories or values of
 * the axis.
 *
 * For example:
 *
 * @code
 *     chart_axis_set_label_position(chart->x_axis, LXW_CHART_AXIS_LABEL_POSITION_HIGH);
       chart_axis_set_label_position(chart->y_axis, LXW_CHART_AXIS_LABEL_POSITION_HIGH);
 * @endcode
 *
 * @image html chart_label_position2.png
 *
 * The allowable values:
 *
 * - #LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO - The default.
 * - #LXW_CHART_AXIS_LABEL_POSITION_HIGH - Also right for vertical axes.
 * - #LXW_CHART_AXIS_LABEL_POSITION_LOW - Also left for vertical axes.
 * - #LXW_CHART_AXIS_LABEL_POSITION_NONE
 *
 * @image html chart_label_position1.png
 *
 * The #LXW_CHART_AXIS_LABEL_POSITION_NONE turns off the axis labels. This
 * is slightly different from `chart_axis_off()` which also turns off the
 * labels but also turns off tick marks.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_label_position(lxw_chart_axis *axis, uint8_t position);

/**
 * @brief Set the minimum value for a chart axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param min  Minimum value for chart axis. Value axes only.
 *
 * Set the minimum value for the axis range.
 *
 * @code
 *     chart_axis_set_min(chart->y_axis, -4);
 *     chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * @image html chart_max_min.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_min(lxw_chart_axis *axis, double min);

/**
 * @brief Set the maximum value for a chart axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param max  Maximum value for chart axis. Value axes only.
 *
 * Set the maximum value for the axis range.
 *
 * @code
 *     chart_axis_set_min(chart->y_axis, -4);
 *     chart_axis_set_max(chart->y_axis, 21);
 * @endcode
 *
 * See the above image.
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_max(lxw_chart_axis *axis, double max);

/**
 * @brief Set the log base of the axis range.
 *
 * @param axis     A pointer to a chart #lxw_chart_axis object.
 * @param log_base The log base for value axis. Value axes only.
 *
 * Set the log base for the axis:
 *
 * @code
 *     chart_axis_set_log_base(chart->y_axis, 10);
 * @endcode
 *
 * @image html chart_log_base.png
 *
 * The allowable range of values for the log base in Excel is between 2 and
 * 1000.
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_log_base(lxw_chart_axis *axis, uint16_t log_base);

/**
 * @brief Set the major axis tick mark type.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param type The tick mark type, defined by #lxw_chart_tick_mark.
 *
 * Set the type of the major axis tick mark:
 *
 * @code
 *     chart_axis_set_major_tick_mark(chart->x_axis, LXW_CHART_AXIS_TICK_MARK_CROSSING);
 *     chart_axis_set_minor_tick_mark(chart->x_axis, LXW_CHART_AXIS_TICK_MARK_INSIDE);
 *
 *     chart_axis_set_major_tick_mark(chart->x_axis, LXW_CHART_AXIS_TICK_MARK_OUTSIDE);
 *     chart_axis_set_minor_tick_mark(chart->y_axis, LXW_CHART_AXIS_TICK_MARK_INSIDE);
 *
 *     // Hide the default gridlines so the tick marks are visible.
 *     chart_axis_major_gridlines_set_visible(chart->y_axis, LXW_FALSE);
 * @endcode
 *
 * @image html chart_tick_marks.png
 *
 * The tick mark types are:
 *
 * - #LXW_CHART_AXIS_TICK_MARK_NONE
 * - #LXW_CHART_AXIS_TICK_MARK_INSIDE
 * - #LXW_CHART_AXIS_TICK_MARK_OUTSIDE
 * - #LXW_CHART_AXIS_TICK_MARK_CROSSING
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_major_tick_mark(lxw_chart_axis *axis, uint8_t type);

/**
 * @brief Set the minor axis tick mark type.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param type The tick mark type, defined by #lxw_chart_tick_mark.
 *
 * Set the type of the minor axis tick mark:
 *
 * @code
 *     chart_axis_set_minor_tick_mark(chart->x_axis, LXW_CHART_AXIS_TICK_MARK_INSIDE);
 * @endcode
 *
 * See the image and example above.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_minor_tick_mark(lxw_chart_axis *axis, uint8_t type);

/**
 * @brief Set the interval between category values.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
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
 *     chart_axis_set_interval_unit(chart->x_axis, 2);
 * @endcode
 *
 * @image html chart_set_interval1.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_interval_unit(lxw_chart_axis *axis, uint16_t unit);

/**
 * @brief Set the interval between category tick marks.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param unit The interval between the category ticks.
 *
 * Set the interval between the category tick marks. The default interval is 1
 * between each category but it can be set to other integer values:
 *
 * @code
 *     chart_axis_set_interval_tick(chart->x_axis, 2);
 * @endcode
 *
 * @image html chart_set_interval2.png
 *
 * **Axis types**: This function is applicable to category and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_interval_tick(lxw_chart_axis *axis, uint16_t unit);

/**
 * @brief Set the increment of the major units in the axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param unit The increment of the major units.
 *
 * Set the increment of the major units in the axis range.
 *
 * @code
 *     // Turn on the minor gridline (it is off by default).
 *     chart_axis_minor_gridlines_set_visible(chart->y_axis, LXW_TRUE);
 *
 *     chart_axis_set_major_unit(chart->y_axis, 4);
 *     chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * @image html chart_set_major_units.png
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_major_unit(lxw_chart_axis *axis, double unit);

/**
 * @brief Set the increment of the minor units in the axis.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param unit The increment of the minor units.
 *
 * Set the increment of the minor units in the axis range.
 *
 * @code
 *     chart_axis_set_minor_unit(chart->y_axis, 2);
 * @endcode
 *
 * See the image above
 *
 * **Axis types**: This function is applicable to value and date axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_minor_unit(lxw_chart_axis *axis, double unit);

/**
 * @brief Set the display units for a value axis.
 *
 * @param axis  A pointer to a chart #lxw_chart_axis object.
 * @param units The display units: #lxw_chart_axis_display_unit.
 *
 * Set the display units for the axis. This can be useful if the axis numbers
 * are very large but you don't want to represent them in scientific notation:
 *
 * @code
 *     chart_axis_set_display_units(chart->x_axis, LXW_CHART_AXIS_UNITS_THOUSANDS);
 *     chart_axis_set_display_units(chart->y_axis, LXW_CHART_AXIS_UNITS_MILLIONS);
 * @endcode
 *
 * @image html chart_display_units.png
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_display_units(lxw_chart_axis *axis, uint8_t units);

/**
 * @brief Turn on/off the display units for a value axis.

 * @param axis    A pointer to a chart #lxw_chart_axis object.
 * @param visible Turn off/on the display units. (0/1)
 *
 * Turn on or off the display units for the axis. This option is set on
 * automatically by `chart_axis_set_display_units()`.
 *
 * @code
 *     chart_axis_set_display_units_visible(chart->y_axis, LXW_TRUE);
 * @endcode
 *
 * **Axis types**: This function is applicable to value axes only.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_set_display_units_visible(lxw_chart_axis *axis,
                                          uint8_t visible);

/**
 * @brief Turn on/off the major gridlines for an axis.
 *
 * @param axis    A pointer to a chart #lxw_chart_axis object.
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
 *     chart_axis_major_gridlines_set_visible(chart->x_axis, LXW_TRUE);
 *     chart_axis_major_gridlines_set_visible(chart->y_axis, LXW_FALSE);
 * @endcode
 *
 * @image html chart_gridline1.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_major_gridlines_set_visible(lxw_chart_axis *axis,
                                            uint8_t visible);

/**
 * @brief Turn on/off the minor gridlines for an axis.
 *
 * @param axis    A pointer to a chart #lxw_chart_axis object.
 * @param visible Turn off/on the minor gridline. (0/1)
 *
 * Turn on or off the minor gridlines for an X or Y axis. In most Excel charts
 * the X and Y axis minor gridlines are off by default.
 *
 * Example, turn on all major and minor gridlines:
 *
 * @code
 *     chart_axis_major_gridlines_set_visible(chart->x_axis, LXW_TRUE);
 *     chart_axis_minor_gridlines_set_visible(chart->x_axis, LXW_TRUE);
 *     chart_axis_major_gridlines_set_visible(chart->y_axis, LXW_TRUE);
 *     chart_axis_minor_gridlines_set_visible(chart->y_axis, LXW_TRUE);
 * @endcode
 *
 * @image html chart_gridline2.png
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_minor_gridlines_set_visible(lxw_chart_axis *axis,
                                            uint8_t visible);

/**
 * @brief Set the line properties for the chart axis major gridlines.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param line A #lxw_chart_line struct.
 *
 * Format the line properties of the major gridlines of a chart:
 *
 * @code
 *     lxw_chart_line line1 = {.color = LXW_COLOR_RED,
 *                             .width = 0.5,
 *                             .dash_type = LXW_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     lxw_chart_line line2 = {.color = LXW_COLOR_YELLOW};
 *
 *     lxw_chart_line line3 = {.width = 1.25,
 *                             .dash_type = LXW_CHART_LINE_DASH_DASH};
 *
 *     lxw_chart_line line4 = {.color =  0x00B050};
 *
 *     chart_axis_major_gridlines_set_line(chart->x_axis, &line1);
 *     chart_axis_minor_gridlines_set_line(chart->x_axis, &line2);
 *     chart_axis_major_gridlines_set_line(chart->y_axis, &line3);
 *     chart_axis_minor_gridlines_set_line(chart->y_axis, &line4);
 * @endcode
 *
 * @image html chart_gridline3.png
 *
 * For more information see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_major_gridlines_set_line(lxw_chart_axis *axis,
                                         lxw_chart_line *line);

/**
 * @brief Set the line properties for the chart axis minor gridlines.
 *
 * @param axis A pointer to a chart #lxw_chart_axis object.
 * @param line A #lxw_chart_line struct.
 *
 * Format the line properties of the minor gridlines of a chart, see the
 * example above.
 *
 * For more information see @ref chart_lines.
 *
 * **Axis types**: This function is applicable to to all axes types.
 *                 See @ref ww_charts_axes.
 */
void chart_axis_minor_gridlines_set_line(lxw_chart_axis *axis,
                                         lxw_chart_line *line);

/**
 * @brief Set the title of the chart.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param name  The chart title name.
 *
 * The `%chart_title_set_name()` function sets the name (title) for the
 * chart. The name is displayed above the chart.
 *
 * @code
 *     chart_title_set_name(chart, "Year End Results");
 * @endcode
 *
 * @image html chart_title_set_name.png
 *
 * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
 * a cell in the workbook that contains the name:
 *
 * @code
 *     chart_title_set_name(chart, "=Sheet1!$B1$1");
 * @endcode
 *
 * See also the `chart_title_set_name_range()` function to see how to set the
 * name formula programmatically.
 *
 * The Excel default is to have no chart title.
 */
void chart_title_set_name(lxw_chart *chart, const char *name);

/**
 * @brief Set a chart title formula using row and column values.
 *
 * @param chart     Pointer to a lxw_chart instance to be configured.
 * @param sheetname The name of the worksheet that contains the cell range.
 * @param row       The zero indexed row number of the range.
 * @param col       The zero indexed column number of the range.
 *
 * The `%chart_title_set_name_range()` function can be used to set a chart
 * title range and is an alternative to using `chart_title_set_name()` and a
 * string formula:
 *
 * @code
 *     chart_title_set_name_range(chart, "Sheet1", 1, 0);
 * @endcode
 */
void chart_title_set_name_range(lxw_chart *chart, const char *sheetname,
                                lxw_row_t row, lxw_col_t col);

/**
 * @brief  Set the font properties for a chart title.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param font  A pointer to a chart #lxw_chart_font font struct.
 *
 * The `%chart_title_set_name_font()` function is used to set the font of a
 * chart title:
 *
 * @code
 *     lxw_chart_font font = {.bold = LXW_TRUE, .color = LXW_COLOR_BLUE};
 *
 *     chart_title_set_name(chart, "Year End Results");
 *     chart_title_set_name_font(chart, &font);
 * @endcode
 *
 * @image html chart_title_set_name_font.png
 *
 * For more information see @ref chart_fonts.
 */
void chart_title_set_name_font(lxw_chart *chart, lxw_chart_font *font);

/**
 * @brief Turn off an automatic chart title.
 *
 * @param chart  Pointer to a lxw_chart instance to be configured.
 *
 * In general in Excel a chart title isn't displayed unless the user
 * explicitly adds one. However, Excel adds an automatic chart title to charts
 * with a single series and a user defined series name. The
 * `chart_title_off()` function allows you to turn off this automatic chart
 * title:
 *
 * @code
 *     chart_title_off(chart);
 * @endcode
 */
void chart_title_off(lxw_chart *chart);

/**
 * @brief Set the position of the chart legend.
 *
 * @param chart    Pointer to a lxw_chart instance to be configured.
 * @param position The #lxw_chart_legend_position value for the legend.
 *
 * The `%chart_legend_set_position()` function is used to set the chart
 * legend to one of the #lxw_chart_legend_position values:
 *
 *     LXW_CHART_LEGEND_NONE
 *     LXW_CHART_LEGEND_RIGHT
 *     LXW_CHART_LEGEND_LEFT
 *     LXW_CHART_LEGEND_TOP
 *     LXW_CHART_LEGEND_BOTTOM
 *     LXW_CHART_LEGEND_OVERLAY_RIGHT
 *     LXW_CHART_LEGEND_OVERLAY_LEFT
 *
 * For example:
 *
 * @code
 *     chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM);
 * @endcode
 *
 * @image html chart_legend_bottom.png
 *
 * This function can also be used to turn off a chart legend:
 *
 * @code
 *     chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE);
 * @endcode
 *
 * @image html chart_legend_none.png
 *
 */
void chart_legend_set_position(lxw_chart *chart, uint8_t position);

/**
 * @brief Set the font properties for a chart legend.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param font  A pointer to a chart #lxw_chart_font font struct.
 *
 * The `%chart_legend_set_font()` function is used to set the font of a
 * chart legend:
 *
 * @code
 *     lxw_chart_font font = {.bold = LXW_TRUE, .color = LXW_COLOR_BLUE};
 *
 *     chart_legend_set_font(chart, &font);
 * @endcode
 *
 * @image html chart_legend_set_font.png
 *
 * For more information see @ref chart_fonts.
 */
void chart_legend_set_font(lxw_chart *chart, lxw_chart_font *font);

/**
 * @brief Remove one or more series from the the legend.
 *
 * @param chart         Pointer to a lxw_chart instance to be configured.
 * @param delete_series An array of zero-indexed values to delete from series.
 *
 * @return A #lxw_error.
 *
 * The `%chart_legend_delete_series()` function allows you to remove/hide one
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
 *     chart_legend_delete_series(chart, series);
 * @endcode
 *
 * @image html chart_legend_delete.png
 */
lxw_error chart_legend_delete_series(lxw_chart *chart,
                                     int16_t delete_series[]);

/**
 * @brief Set the line properties for a chartarea.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param line  A #lxw_chart_line struct.
 *
 * Set the line/border properties of a chartarea. In Excel the chartarea
 * is the background area behind the chart:
 *
 * @code
 *     lxw_chart_line line = {.none  = LXW_TRUE};
 *     lxw_chart_fill fill = {.color = LXW_COLOR_RED};
 *
 *     chart_chartarea_set_line(chart, &line);
 *     chart_chartarea_set_fill(chart, &fill);
 * @endcode
 *
 * @image html chart_chartarea.png
 *
 * For more information see @ref chart_lines.
 */
void chart_chartarea_set_line(lxw_chart *chart, lxw_chart_line *line);

/**
 * @brief Set the fill properties for a chartarea.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param fill  A #lxw_chart_fill struct.
 *
 * Set the fill properties of a chartarea:
 *
 * @code
 *     chart_chartarea_set_fill(chart, &fill);
 * @endcode
 *
 * See the example and image above.
 *
 * For more information see @ref chart_fills.
 */
void chart_chartarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);

/**
 * @brief Set the pattern properties for a chartarea.
 *
 * @param chart   Pointer to a lxw_chart instance to be configured.
 * @param pattern A #lxw_chart_pattern struct.
 *
 * Set the pattern properties of a chartarea:
 *
 * @code
 *     chart_chartarea_set_pattern(series1, &pattern);
 * @endcode
 *
 * For more information see #lxw_chart_pattern_type and @ref chart_patterns.
 */
void chart_chartarea_set_pattern(lxw_chart *chart,
                                 lxw_chart_pattern *pattern);

/**
 * @brief Set the line properties for a plotarea.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param line  A #lxw_chart_line struct.
 *
 * Set the line/border properties of a plotarea. In Excel the plotarea is
 * the area between the axes on which the chart series are plotted:
 *
 * @code
 *     lxw_chart_line line = {.color     = LXW_COLOR_RED,
 *                            .width     = 2,
 *                            .dash_type = LXW_CHART_LINE_DASH_DASH};
 *     lxw_chart_fill fill = {.color     = 0xFFFFC2};
 *
 *     chart_plotarea_set_line(chart, &line);
 *     chart_plotarea_set_fill(chart, &fill);
 *
 * @endcode
 *
 * @image html chart_plotarea.png
 *
 * For more information see @ref chart_lines.
 */
void chart_plotarea_set_line(lxw_chart *chart, lxw_chart_line *line);

/**
 * @brief Set the fill properties for a plotarea.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param fill  A #lxw_chart_fill struct.
 *
 * Set the fill properties of a plotarea:
 *
 * @code
 *     chart_plotarea_set_fill(chart, &fill);
 * @endcode
 *
 * See the example and image above.
 *
 * For more information see @ref chart_fills.
 */
void chart_plotarea_set_fill(lxw_chart *chart, lxw_chart_fill *fill);

/**
 * @brief Set the pattern properties for a plotarea.
 *
 * @param chart   Pointer to a lxw_chart instance to be configured.
 * @param pattern A #lxw_chart_pattern struct.
 *
 * Set the pattern properties of a plotarea:
 *
 * @code
 *     chart_plotarea_set_pattern(series1, &pattern);
 * @endcode
 *
 * For more information see #lxw_chart_pattern_type and @ref chart_patterns.
 */
void chart_plotarea_set_pattern(lxw_chart *chart, lxw_chart_pattern *pattern);

/**
 * @brief Set the chart style type.
 *
 * @param chart    Pointer to a lxw_chart instance to be configured.
 * @param style_id An index representing the chart style, 1 - 48.
 *
 * The `%chart_set_style()` function is used to set the style of the chart to
 * one of the 48 built-in styles available on the "Design" tab in Excel 2007:
 *
 * @code
 *     chart_set_style(chart, 37)
 * @endcode
 *
 * @image html chart_style.png
 *
 * The style index number is counted from 1 on the top left in the Excel
 * dialog. The default style is 2.
 *
 * **Note:**
 *
 * In Excel 2013 the Styles section of the "Design" tab in Excel shows what
 * were referred to as "Layouts" in previous versions of Excel. These layouts
 * are not defined in the file format. They are a collection of modifications
 * to the base chart type. They can not be defined by the `chart_set_style()``
 * function.
 *
 */
void chart_set_style(lxw_chart *chart, uint8_t style_id);

/**
 * @brief Turn on a data table below the horizontal axis.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 *
 * The `%chart_set_table()` function adds a data table below the horizontal
 * axis with the data used to plot the chart:
 *
 * @code
 *     // Turn on the data table with default options.
 *     chart_set_table(chart);
 * @endcode
 *
 * @image html chart_data_table1.png
 *
 * The data table can only be shown with Bar, Column, Line and Area charts.
 *
 */
void chart_set_table(lxw_chart *chart);

/**
 * @brief Turn on/off grid options for a chart data table.
 *
 * @param chart       Pointer to a lxw_chart instance to be configured.
 * @param horizontal  Turn on/off the horizontal grid lines in the table.
 * @param vertical    Turn on/off the vertical grid lines in the table.
 * @param outline     Turn on/off the outline lines in the table.
 * @param legend_keys Turn on/off the legend keys in the table.
 *
 * The `%chart_set_table_grid()` function turns on/off grid options for a
 * chart data table. The data table grid options in Excel are shown in the
 * dialog below:
 *
 * @image html chart_data_table3.png
 *
 * These options can be passed to the `%chart_set_table_grid()` function.
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
 *     chart_set_table(chart);
 *
 *     // Turn on all grid lines and the grid legend.
 *     chart_set_table_grid(chart, LXW_TRUE, LXW_TRUE, LXW_TRUE, LXW_TRUE);
 *
 *     // Turn off the legend since it is show in the table.
 *     chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE);
 *
 * @endcode
 *
 * @image html chart_data_table2.png
 *
 * The data table can only be shown with Bar, Column, Line and Area charts.
 *
 */
void chart_set_table_grid(lxw_chart *chart, uint8_t horizontal,
                          uint8_t vertical, uint8_t outline,
                          uint8_t legend_keys);

void chart_set_table_font(lxw_chart *chart, lxw_chart_font *font);

/**
 * @brief Turn on up-down bars for the chart.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 *
 * The `%chart_set_up_down_bars()` function adds Up-Down bars to Line charts
 * to indicate the difference between the first and last data series:
 *
 * @code
 *     chart_set_up_down_bars(chart);
 * @endcode
 *
 * @image html chart_data_tools4.png
 *
 * Up-Down bars are only available in Line charts. By default Up-Down bars are
 * black and white like in the above example. To format the border or fill
 * of the bars see the `chart_set_up_down_bars_format()` function below.
 */
void chart_set_up_down_bars(lxw_chart *chart);

/**
 * @brief Turn on up-down bars for the chart, with formatting.
 *
 * @param chart         Pointer to a lxw_chart instance to be configured.
 * @param up_bar_line   A #lxw_chart_line struct for the up-bar border.
 * @param up_bar_fill   A #lxw_chart_fill struct for the up-bar fill.
 * @param down_bar_line A #lxw_chart_line struct for the down-bar border.
 * @param down_bar_fill A #lxw_chart_fill struct for the down-bar fill.
 *
 * The `%chart_set_up_down_bars_format()` function adds Up-Down bars to Line
 * charts to indicate the difference between the first and last data series.
 * It also allows the up and down bars to be formatted:
 *
 * @code
 *     lxw_chart_line line      = {.color = LXW_COLOR_BLACK};
 *     lxw_chart_fill up_fill   = {.color = 0x00B050};
 *     lxw_chart_fill down_fill = {.color = LXW_COLOR_RED};
 *
 *     chart_set_up_down_bars_format(chart, &line, &up_fill, &line, &down_fill);
 * @endcode
 *
 * @image html chart_up_down_bars.png
 *
 * Up-Down bars are only available in Line charts.
 * For more format information  see @ref chart_lines and @ref chart_fills.
 */
void chart_set_up_down_bars_format(lxw_chart *chart,
                                   lxw_chart_line *up_bar_line,
                                   lxw_chart_fill *up_bar_fill,
                                   lxw_chart_line *down_bar_line,
                                   lxw_chart_fill *down_bar_fill);

/**
 * @brief Turn on and format Drop Lines for a chart.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param line  A #lxw_chart_line struct.
 *
 * The `%chart_set_drop_lines()` function adds Drop Lines to charts to
 * show the Category value of points in the data:
 *
 * @code
 *     chart_set_drop_lines(chart, NULL);
 * @endcode
 *
 * @image html chart_data_tools6.png
 *
 * It is possible to format the Drop Line line properties if required:
 *
 * @code
 *     lxw_chart_line line = {.color     = LXW_COLOR_RED,
 *                            .dash_type = LXW_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     chart_set_drop_lines(chart, &line);
 * @endcode
 *
 * Drop Lines are only available in Line and Area charts.
 * For more format information see @ref chart_lines.
 */
void chart_set_drop_lines(lxw_chart *chart, lxw_chart_line *line);

/**
 * @brief Turn on and format high-low Lines for a chart.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param line  A #lxw_chart_line struct.
 *
 * The `%chart_set_high_low_lines()` function adds High-Low Lines to charts
 * to show the Category value of points in the data:
 *
 * @code
 *     chart_set_high_low_lines(chart, NULL);
 * @endcode
 *
 * @image html chart_data_tools5.png
 *
 * It is possible to format the High-Low Line line properties if required:
 *
 * @code
 *     lxw_chart_line line = {.color     = LXW_COLOR_RED,
 *                            .dash_type = LXW_CHART_LINE_DASH_SQUARE_DOT};
 *
 *     chart_set_high_low_lines(chart, &line);
 * @endcode
 *
 * High-Low Lines are only available in Line charts.
 * For more format information see @ref chart_lines.
 */
void chart_set_high_low_lines(lxw_chart *chart, lxw_chart_line *line);

/**
 * @brief Set the overlap between series in a Bar/Column chart.
 *
 * @param chart   Pointer to a lxw_chart instance to be configured.
 * @param overlap The overlap between the series. -100 to 100.
 *
 * The `%chart_set_series_overlap()` function sets the overlap between series
 * in Bar and Column charts.
 *
 * @code
 *     chart_set_series_overlap(chart, -50);
 * @endcode
 *
 * @image html chart_overlap.png
 *
 * The overlap value must be in the range `0 <= overlap <= 500`.
 * The default value is 0.
 *
 * This option is only available for Bar/Column charts.
 */
void chart_set_series_overlap(lxw_chart *chart, int8_t overlap);

/**
 * @brief Set the gap between series in a Bar/Column chart.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param gap   The gap between the series.  0 to 500.
 *
 * The `%chart_set_series_gap()` function sets the gap between series in
 * Bar and Column charts.
 *
 * @code
 *     chart_set_series_gap(chart, 400);
 * @endcode
 *
 * @image html chart_gap.png
 *
 * The gap value must be in the range `0 <= gap <= 500`. The default value
 * is 150.
 *
 * This option is only available for Bar/Column charts.
 */
void chart_set_series_gap(lxw_chart *chart, uint16_t gap);

/**
 * @brief Set the option for displaying blank data in a chart.
 *
 * @param chart    Pointer to a lxw_chart instance to be configured.
 * @param option The display option. A #lxw_chart_blank option.
 *
 * The `%chart_show_blanks_as()` function controls how blank data is displayed
 * in a chart:
 *
 * @code
 *     chart_show_blanks_as(chart, LXW_CHART_BLANKS_AS_CONNECTED);
 * @endcode
 *
 * The `option` parameter can have one of the following values:
 *
 * - #LXW_CHART_BLANKS_AS_GAP: Show empty chart cells as gaps in the data.
 *   This is the default option for Excel charts.
 * - #LXW_CHART_BLANKS_AS_ZERO: Show empty chart cells as zeros.
 * - #LXW_CHART_BLANKS_AS_CONNECTED: Show empty chart cells as connected.
 *   Only for charts with lines.
 */
void chart_show_blanks_as(lxw_chart *chart, uint8_t option);

/**
 * @brief Display data on charts from hidden rows or columns.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 *
 * Display data that is in hidden rows or columns on the chart:
 *
 * @code
 *     chart_show_hidden_data(chart);
 * @endcode
 */
void chart_show_hidden_data(lxw_chart *chart);

/**
 * @brief Set the Pie/Doughnut chart rotation.
 *
 * @param chart    Pointer to a lxw_chart instance to be configured.
 * @param rotation The angle of rotation.
 *
 * The `chart_set_rotation()` function is used to set the rotation of the
 * first segment of a Pie/Doughnut chart. This has the effect of rotating
 * the entire chart:
 *
 * @code
 *     chart_set_rotation(chart, 28);
 * @endcode
 *
 * The angle of rotation must be in the range `0 <= rotation <= 360`.
 *
 * This option is only available for Pie/Doughnut charts.
 *
 */
void chart_set_rotation(lxw_chart *chart, uint16_t rotation);

/**
 * @brief Set the Doughnut chart hole size.
 *
 * @param chart Pointer to a lxw_chart instance to be configured.
 * @param size  The hole size as a percentage.
 *
 * The `chart_set_hole_size()` function is used to set the hole size of a
 * Doughnut chart:
 *
 * @code
 *     chart_set_hole_size(chart, 33);
 * @endcode
 *
 * The hole size must be in the range `10 <= size <= 90`.
 *
 * This option is only available for Doughnut charts.
 *
 */
void chart_set_hole_size(lxw_chart *chart, uint8_t size);

lxw_error lxw_chart_add_data_cache(lxw_series_range *range, uint8_t *data,
                                   uint16_t rows, uint8_t cols, uint8_t col);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _chart_xml_declaration(lxw_chart *chart);
STATIC void _chart_write_legend(lxw_chart *chart);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_CHART_H__ */
