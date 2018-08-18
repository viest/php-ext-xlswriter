/*****************************************************************************
 * chart - A library for creating Excel XLSX chart files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/chart.h"
#include "xlsxwriter/utility.h"

/*
 * Forward declarations.
 */

STATIC void _chart_initialize(lxw_chart *self, uint8_t type);
STATIC void _chart_axis_set_default_num_format(lxw_chart_axis *axis,
                                               char *num_format);

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Free a series range object.
 */
STATIC void
_chart_free_range(lxw_series_range *range)
{
    struct lxw_series_data_point *data_point;

    if (!range)
        return;

    if (range->data_cache) {
        while (!STAILQ_EMPTY(range->data_cache)) {
            data_point = STAILQ_FIRST(range->data_cache);
            free(data_point->string);
            STAILQ_REMOVE_HEAD(range->data_cache, list_pointers);

            free(data_point);
        }
        free(range->data_cache);
    }

    free(range->formula);
    free(range->sheetname);
    free(range);
}

STATIC void
_chart_free_points(lxw_chart_series *series)
{
    uint16_t index;

    for (index = 0; index < series->point_count; index++) {
        lxw_chart_point *point = &series->points[index];

        free(point->line);
        free(point->fill);
        free(point->pattern);
    }

    series->point_count = 0;
    free(series->points);
}

/*
 * Free a chart font object.
 */
STATIC void
_chart_free_font(lxw_chart_font *font)
{
    if (!font)
        return;

    free(font->name);
    free(font);
}

/*
 * Free a series object.
 */
STATIC void
_chart_series_free(lxw_chart_series *series)
{
    if (!series)
        return;

    free(series->title.name);
    free(series->line);
    free(series->fill);
    free(series->pattern);
    free(series->label_num_format);
    _chart_free_font(series->label_font);

    if (series->marker) {
        free(series->marker->line);
        free(series->marker->fill);
        free(series->marker->pattern);
        free(series->marker);
    }

    _chart_free_range(series->categories);
    _chart_free_range(series->values);
    _chart_free_range(series->title.range);
    _chart_free_points(series);

    if (series->x_error_bars) {
        free(series->x_error_bars->line);
        free(series->x_error_bars);
    }

    if (series->y_error_bars) {
        free(series->y_error_bars->line);
        free(series->y_error_bars);
    }

    free(series->trendline_line);
    free(series->trendline_name);

    free(series);
}

/*
 * Initialize the data cache in a range object.
 */
STATIC lxw_error
_chart_init_data_cache(lxw_series_range *range)
{
    /* Initialize the series range data cache. */
    range->data_cache = calloc(1, sizeof(struct lxw_series_data_points));
    RETURN_ON_MEM_ERROR(range->data_cache, LXW_ERROR_MEMORY_MALLOC_FAILED);
    STAILQ_INIT(range->data_cache);

    return LXW_NO_ERROR;
}

/*
 * Free a chart object.
 */
void
lxw_chart_free(lxw_chart *chart)
{
    lxw_chart_series *series;

    if (!chart)
        return;

    /* Chart series. */
    if (chart->series_list) {
        while (!STAILQ_EMPTY(chart->series_list)) {
            series = STAILQ_FIRST(chart->series_list);
            STAILQ_REMOVE_HEAD(chart->series_list, list_pointers);

            _chart_series_free(series);
        }

        free(chart->series_list);
    }

    /* X Axis. */
    if (chart->x_axis) {
        _chart_free_font(chart->x_axis->title.font);
        _chart_free_font(chart->x_axis->num_font);
        _chart_free_range(chart->x_axis->title.range);
        free(chart->x_axis->title.name);
        free(chart->x_axis->line);
        free(chart->x_axis->fill);
        free(chart->x_axis->pattern);
        free(chart->x_axis->major_gridlines.line);
        free(chart->x_axis->minor_gridlines.line);
        free(chart->x_axis->num_format);
        free(chart->x_axis->default_num_format);
        free(chart->x_axis);
    }

    /* Y Axis. */
    if (chart->y_axis) {
        _chart_free_font(chart->y_axis->title.font);
        _chart_free_font(chart->y_axis->num_font);
        _chart_free_range(chart->y_axis->title.range);
        free(chart->y_axis->title.name);
        free(chart->y_axis->line);
        free(chart->y_axis->fill);
        free(chart->y_axis->pattern);
        free(chart->y_axis->major_gridlines.line);
        free(chart->y_axis->minor_gridlines.line);
        free(chart->y_axis->num_format);
        free(chart->y_axis->default_num_format);
        free(chart->y_axis);
    }

    /* Chart title. */
    _chart_free_font(chart->title.font);
    _chart_free_range(chart->title.range);
    free(chart->title.name);

    /* Chart legend. */
    _chart_free_font(chart->legend.font);
    free(chart->delete_series);

    free(chart->default_marker);

    free(chart->chartarea_line);
    free(chart->chartarea_fill);
    free(chart->chartarea_pattern);
    free(chart->plotarea_line);
    free(chart->plotarea_fill);
    free(chart->plotarea_pattern);

    free(chart->drop_lines_line);
    free(chart->high_low_lines_line);

    free(chart->up_bar_line);
    free(chart->up_bar_fill);
    free(chart->down_bar_line);
    free(chart->down_bar_fill);

    _chart_free_font(chart->table_font);

    free(chart);
}

/*
 * Create a new chart object.
 */
lxw_chart *
lxw_chart_new(uint8_t type)
{
    lxw_chart *chart = calloc(1, sizeof(lxw_chart));
    GOTO_LABEL_ON_MEM_ERROR(chart, mem_error);

    chart->series_list = calloc(1, sizeof(struct lxw_chart_series_list));
    GOTO_LABEL_ON_MEM_ERROR(chart->series_list, mem_error);
    STAILQ_INIT(chart->series_list);

    chart->x_axis = calloc(1, sizeof(struct lxw_chart_axis));
    GOTO_LABEL_ON_MEM_ERROR(chart->x_axis, mem_error);

    chart->y_axis = calloc(1, sizeof(struct lxw_chart_axis));
    GOTO_LABEL_ON_MEM_ERROR(chart->y_axis, mem_error);

    chart->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(chart->title.range, mem_error);

    chart->x_axis->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(chart->x_axis->title.range, mem_error);

    chart->y_axis->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(chart->y_axis->title.range, mem_error);

    /* Initialize the ranges in the chart titles. */
    if (_chart_init_data_cache(chart->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(chart->x_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(chart->y_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    chart->type = type;
    chart->style_id = 2;
    chart->hole_size = 50;

    /* Set the default axis positions. */
    chart->x_axis->axis_position = LXW_CHART_AXIS_BOTTOM;
    chart->y_axis->axis_position = LXW_CHART_AXIS_LEFT;

    /* Set the default axis number formats. */
    _chart_axis_set_default_num_format(chart->x_axis, "General");
    _chart_axis_set_default_num_format(chart->y_axis, "General");

    chart->x_axis->major_gridlines.visible = LXW_FALSE;
    chart->y_axis->major_gridlines.visible = LXW_TRUE;

    chart->has_horiz_cat_axis = LXW_FALSE;
    chart->has_horiz_val_axis = LXW_TRUE;

    chart->legend.position = LXW_CHART_LEGEND_RIGHT;

    chart->gap_y1 = LXW_CHART_DEFAULT_GAP;
    chart->gap_y2 = LXW_CHART_DEFAULT_GAP;

    /* Initialize the chart specific properties. */
    _chart_initialize(chart, chart->type);

    return chart;

mem_error:
    lxw_chart_free(chart);
    return NULL;
}

/*
 * Create a copy of a user supplied font.
 */
STATIC lxw_chart_font *
_chart_convert_font_args(lxw_chart_font *user_font)
{
    lxw_chart_font *font;

    if (!user_font)
        return NULL;

    font = calloc(1, sizeof(struct lxw_chart_font));
    RETURN_ON_MEM_ERROR(font, NULL);

    /* Copy the user supplied properties. */
    font->name = lxw_strdup(user_font->name);
    font->size = user_font->size;
    font->bold = user_font->bold;
    font->italic = user_font->italic;
    font->underline = user_font->underline;
    font->rotation = user_font->rotation;
    font->color = user_font->color;
    font->pitch_family = user_font->pitch_family;
    font->charset = user_font->charset;
    font->baseline = user_font->baseline;

    /* Convert font size units. */
    if (font->size > 0.0)
        font->size = font->size * 100.0;

    /* Convert rotation into 60,000ths of a degree. */
    if (font->rotation)
        font->rotation = font->rotation * 60000;

    if (font->color) {
        font->color = lxw_format_check_color(font->color);
        font->has_color = LXW_TRUE;
    }

    return font;
}

/*
 * Create a copy of a user supplied line.
 */
STATIC lxw_chart_line *
_chart_convert_line_args(lxw_chart_line *user_line)
{
    lxw_chart_line *line;

    if (!user_line)
        return NULL;

    line = calloc(1, sizeof(struct lxw_chart_line));
    RETURN_ON_MEM_ERROR(line, NULL);

    /* Copy the user supplied properties. */
    line->color = user_line->color;
    line->none = user_line->none;
    line->width = user_line->width;
    line->dash_type = user_line->dash_type;
    line->transparency = user_line->transparency;

    if (line->color) {
        line->color = lxw_format_check_color(line->color);
        line->has_color = LXW_TRUE;
    }

    if (line->transparency > 100)
        line->transparency = 0;

    return line;
}

/*
 * Create a copy of a user supplied fill.
 */
STATIC lxw_chart_fill *
_chart_convert_fill_args(lxw_chart_fill *user_fill)
{
    lxw_chart_fill *fill;

    if (!user_fill)
        return NULL;

    fill = calloc(1, sizeof(struct lxw_chart_fill));
    RETURN_ON_MEM_ERROR(fill, NULL);

    /* Copy the user supplied properties. */
    fill->color = user_fill->color;
    fill->none = user_fill->none;
    fill->transparency = user_fill->transparency;

    if (fill->color) {
        fill->color = lxw_format_check_color(fill->color);
        fill->has_color = LXW_TRUE;
    }

    if (fill->transparency > 100)
        fill->transparency = 0;

    return fill;
}

/*
 * Create a copy of a user supplied pattern.
 */
STATIC lxw_chart_pattern *
_chart_convert_pattern_args(lxw_chart_pattern *user_pattern)
{
    lxw_chart_pattern *pattern;

    if (!user_pattern)
        return NULL;

    if (!user_pattern->type) {
        LXW_WARN("chart_xxx_set_pattern: 'type' must be specified");
        return NULL;
    }

    if (!user_pattern->fg_color) {
        LXW_WARN("chart_xxx_set_pattern: 'fg_color' must be specified");
        return NULL;
    }

    pattern = calloc(1, sizeof(struct lxw_chart_pattern));
    RETURN_ON_MEM_ERROR(pattern, NULL);

    /* Copy the user supplied properties. */
    pattern->fg_color = user_pattern->fg_color;
    pattern->bg_color = user_pattern->bg_color;
    pattern->type = user_pattern->type;

    pattern->fg_color = lxw_format_check_color(pattern->fg_color);
    pattern->has_fg_color = LXW_TRUE;

    if (pattern->bg_color) {
        pattern->bg_color = lxw_format_check_color(pattern->bg_color);
        pattern->has_bg_color = LXW_TRUE;
    }
    else {
        /* Default background color in Excel is white, when unspecified. */
        pattern->bg_color = LXW_COLOR_WHITE;
        pattern->has_bg_color = LXW_TRUE;
    }

    return pattern;
}

/*
 * Set a marker type for a series.
 */
STATIC void
_chart_set_default_marker_type(lxw_chart *self, uint8_t type)
{
    if (!self->default_marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        self->default_marker = marker;
    }

    self->default_marker->type = type;
}

/*
 * Set an axis number format.
 */
void
_chart_axis_set_default_num_format(lxw_chart_axis *axis, char *num_format)
{
    if (!num_format)
        return;

    /* Free any previously allocated resource. */
    free(axis->default_num_format);

    axis->default_num_format = lxw_strdup(num_format);
}

/*
 * Verify that a X/Y error bar property is support for the chart type.
 * All chart types, except Bar have Y error bars. Only Bar and Scatter
 * support X error bars.
 */
lxw_error
_chart_check_error_bars(lxw_series_error_bars *error_bars, char *property)
{
    /* Check that the error bar type has been set for all error bar
     * functions except the one that is used to set the type. */
    if (strlen(property) && !error_bars->is_set) {
        LXW_WARN_FORMAT1("chart_series_set_error_bars%s(): "
                         "error bar type must be set first using "
                         "chart_series_set_error_bars()", property);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (error_bars->is_x) {
        if (error_bars->chart_group != LXW_CHART_SCATTER
            && error_bars->chart_group != LXW_CHART_BAR) {

            LXW_WARN_FORMAT1("chart_series_set_error_bars%s(): "
                             "'X error bar' properties only available for"
                             " Scatter and Bar charts in Excel", property);

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }
    else {
        if (error_bars->chart_group == LXW_CHART_BAR) {
            LXW_WARN_FORMAT1("chart_series_set_error_bars%s(): "
                             "'Y error bar' properties not available for "
                             "Bar charts in Excel", property);

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Add unique ids for primary or secondary axes.
 */
STATIC void
_chart_add_axis_ids(lxw_chart *self)
{
    uint32_t chart_id = 50010000 + self->id;
    uint32_t axis_count = 1;

    self->axis_id_1 = chart_id + axis_count;
    self->axis_id_2 = self->axis_id_1 + 1;
}

/*
 * Utility function to set a chart range.
 */
STATIC void
_chart_set_range(lxw_series_range *range, const char *sheetname,
                 lxw_row_t first_row, lxw_col_t first_col,
                 lxw_row_t last_row, lxw_col_t last_col)
{
    char formula[LXW_MAX_FORMULA_RANGE_LENGTH] = { 0 };

    /* Set the range properties. */
    range->sheetname = lxw_strdup(sheetname);
    range->first_row = first_row;
    range->first_col = first_col;
    range->last_row = last_row;
    range->last_col = last_col;

    /* Free any existing range. */
    free(range->formula);

    /* Convert the range properties to a formula like: Sheet1!$A$1:$A$5. */
    lxw_rowcol_to_formula_abs(formula, sheetname,
                              first_row, first_col, last_row, last_col);

    range->formula = lxw_strdup(formula);
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
_chart_xml_declaration(lxw_chart *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <c:chartSpace> element.
 */
STATIC void
_chart_write_chart_space(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_c[] = LXW_SCHEMA_DRAWING "/chart";
    char xmlns_a[] = LXW_SCHEMA_DRAWING "/main";
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:c", xmlns_c);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:a", xmlns_a);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxw_xml_start_tag(self->file, "c:chartSpace", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lang> element.
 */
STATIC void
_chart_write_lang(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "en-US");

    lxw_xml_empty_tag(self->file, "c:lang", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:style> element.
 */
STATIC void
_chart_write_style(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Don"t write an element for the default style, 2. */
    if (self->style_id == 2)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", self->style_id);

    lxw_xml_empty_tag(self->file, "c:style", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:layout> element.
 */
STATIC void
_chart_write_layout(lxw_chart *self)
{
    lxw_xml_empty_tag(self->file, "c:layout", NULL);
}

/*
 * Write the <c:grouping> element.
 */
STATIC void
_chart_write_grouping(lxw_chart *self, uint8_t grouping)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (grouping == LXW_GROUPING_STANDARD)
        LXW_PUSH_ATTRIBUTES_STR("val", "standard");
    else if (grouping == LXW_GROUPING_PERCENTSTACKED)
        LXW_PUSH_ATTRIBUTES_STR("val", "percentStacked");
    else if (grouping == LXW_GROUPING_STACKED)
        LXW_PUSH_ATTRIBUTES_STR("val", "stacked");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "clustered");

    lxw_xml_empty_tag(self->file, "c:grouping", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:radarStyle> element.
 */
STATIC void
_chart_write_radar_style(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (self->type == LXW_CHART_RADAR_FILLED)
        LXW_PUSH_ATTRIBUTES_STR("val", "filled");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "marker");

    lxw_xml_empty_tag(self->file, "c:radarStyle", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:varyColors> element.
 */
STATIC void
_chart_write_vary_colors(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:varyColors", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:firstSliceAng> element.
 */
STATIC void
_chart_write_first_slice_ang(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", self->rotation);

    lxw_xml_empty_tag(self->file, "c:firstSliceAng", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:holeSize> element.
 */
STATIC void
_chart_write_hole_size(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", self->hole_size);

    lxw_xml_empty_tag(self->file, "c:holeSize", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:alpha> element.
 */
STATIC void
_chart_write_a_alpha(lxw_chart *self, uint8_t transparency)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint32_t val;

    LXW_INIT_ATTRIBUTES();

    val = (100 - transparency) * 1000;

    LXW_PUSH_ATTRIBUTES_INT("val", val);

    lxw_xml_empty_tag(self->file, "a:alpha", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:srgbClr> element.
 */
STATIC void
_chart_write_a_srgb_clr(lxw_chart *self, lxw_color_t color,
                        uint8_t transparency)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rgb_str[LXW_ATTR_32];

    LXW_INIT_ATTRIBUTES();

    lxw_snprintf(rgb_str, LXW_ATTR_32, "%06X", color & LXW_COLOR_MASK);
    LXW_PUSH_ATTRIBUTES_STR("val", rgb_str);

    if (transparency) {
        lxw_xml_start_tag(self->file, "a:srgbClr", &attributes);

        /* Write the a:alpha element. */
        _chart_write_a_alpha(self, transparency);

        lxw_xml_end_tag(self->file, "a:srgbClr");
    }
    else {
        lxw_xml_empty_tag(self->file, "a:srgbClr", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:solidFill> element.
 */
STATIC void
_chart_write_a_solid_fill(lxw_chart *self, lxw_color_t color,
                          uint8_t transparency)
{

    lxw_xml_start_tag(self->file, "a:solidFill", NULL);

    /* Write the a:srgbClr element. */
    _chart_write_a_srgb_clr(self, color, transparency);

    lxw_xml_end_tag(self->file, "a:solidFill");
}

/*
 * Write the <a:t> element.
 */
STATIC void
_chart_write_a_t(lxw_chart *self, char *name)
{
    lxw_xml_data_element(self->file, "a:t", name, NULL);
}

/*
 * Write the <a:endParaRPr> element.
 */
STATIC void
_chart_write_a_end_para_rpr(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("lang", "en-US");

    lxw_xml_empty_tag(self->file, "a:endParaRPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:defRPr> element.
 */
STATIC void
_chart_write_a_def_rpr(lxw_chart *self, lxw_chart_font *font)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint8_t has_color = LXW_FALSE;
    uint8_t has_latin = LXW_FALSE;
    uint8_t use_font_default = LXW_FALSE;

    LXW_INIT_ATTRIBUTES();

    if (font) {
        has_color = font->color || font->has_color;
        has_latin = font->name || font->pitch_family || font->charset;
        use_font_default = !(has_color || has_latin || font->baseline == -1);

        /* Set the font attributes. */
        if (font->size > 0.0)
            LXW_PUSH_ATTRIBUTES_DBL("sz", font->size);

        if (use_font_default || font->bold)
            LXW_PUSH_ATTRIBUTES_INT("b", font->bold & 0x1);

        if (use_font_default || font->italic)
            LXW_PUSH_ATTRIBUTES_INT("i", font->italic & 0x1);

        if (font->underline)
            LXW_PUSH_ATTRIBUTES_STR("u", "sng");

        if (font->baseline != -1)
            LXW_PUSH_ATTRIBUTES_INT("baseline", font->baseline);
    }

    /* There are sub-elements if the font name or color have changed. */
    if (has_latin || has_color) {

        lxw_xml_start_tag(self->file, "a:defRPr", &attributes);

        if (has_color) {
            _chart_write_a_solid_fill(self, font->color, LXW_FALSE);
        }

        if (has_latin) {
            /* Free and reuse the attribute list for the latin attributes. */
            LXW_FREE_ATTRIBUTES();

            if (font->name)
                LXW_PUSH_ATTRIBUTES_STR("typeface", font->name);

            if (font->pitch_family)
                LXW_PUSH_ATTRIBUTES_INT("pitchFamily", font->pitch_family);

            if (font->pitch_family || font->charset)
                LXW_PUSH_ATTRIBUTES_INT("charset", font->charset);

            /* Write the <a:latin> element. */
            lxw_xml_empty_tag(self->file, "a:latin", &attributes);
        }

        lxw_xml_end_tag(self->file, "a:defRPr");
    }
    else {
        lxw_xml_empty_tag(self->file, "a:defRPr", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:rPr> element.
 */
STATIC void
_chart_write_a_r_pr(lxw_chart *self, lxw_chart_font *font)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint8_t has_color = LXW_FALSE;
    uint8_t has_latin = LXW_FALSE;
    uint8_t use_font_default = LXW_FALSE;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("lang", "en-US");

    if (font) {
        has_color = font->color || font->has_color;
        has_latin = font->name || font->pitch_family || font->charset;
        use_font_default = !(has_color || has_latin || font->baseline == -1);

        /* Set the font attributes. */
        if (font->size > 0.0)
            LXW_PUSH_ATTRIBUTES_DBL("sz", font->size);

        if (use_font_default || font->bold)
            LXW_PUSH_ATTRIBUTES_INT("b", font->bold & 0x1);

        if (use_font_default || font->italic)
            LXW_PUSH_ATTRIBUTES_INT("i", font->italic & 0x1);

        if (font->underline)
            LXW_PUSH_ATTRIBUTES_STR("u", "sng");

        if (font->baseline != -1)
            LXW_PUSH_ATTRIBUTES_INT("baseline", font->baseline);
    }

    /* There are sub-elements if the font name or color have changed. */
    if (has_latin || has_color) {

        lxw_xml_start_tag(self->file, "a:rPr", &attributes);

        if (has_color) {
            _chart_write_a_solid_fill(self, font->color, LXW_FALSE);
        }

        if (has_latin) {
            /* Free and reuse the attribute list for the latin attributes. */
            LXW_FREE_ATTRIBUTES();

            if (font->name)
                LXW_PUSH_ATTRIBUTES_STR("typeface", font->name);

            if (font->pitch_family)
                LXW_PUSH_ATTRIBUTES_INT("pitchFamily", font->pitch_family);

            if (font->pitch_family || font->charset)
                LXW_PUSH_ATTRIBUTES_INT("charset", font->charset);

            /* Write the <a:latin> element. */
            lxw_xml_empty_tag(self->file, "a:latin", &attributes);
        }

        lxw_xml_end_tag(self->file, "a:rPr");
    }
    else {
        lxw_xml_empty_tag(self->file, "a:rPr", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:r> element.
 */
STATIC void
_chart_write_a_r(lxw_chart *self, char *name, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:r", NULL);

    /* Write the a:rPr element. */
    _chart_write_a_r_pr(self, font);

    /* Write the a:t element. */
    _chart_write_a_t(self, name);

    lxw_xml_end_tag(self->file, "a:r");
}

/*
 * Write the <a:pPr> element.
 */
STATIC void
_chart_write_a_p_pr_formula(lxw_chart *self, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:pPr", NULL);

    /* Write the a:defRPr element. */
    _chart_write_a_def_rpr(self, font);

    lxw_xml_end_tag(self->file, "a:pPr");
}

/*
 * Write the <a:pPr> element for pie chart legends.
 */
STATIC void
_chart_write_a_p_pr_pie(lxw_chart *self, lxw_chart_font *font)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("rtl", "0");

    lxw_xml_start_tag(self->file, "a:pPr", &attributes);

    /* Write the a:defRPr element. */
    _chart_write_a_def_rpr(self, font);

    lxw_xml_end_tag(self->file, "a:pPr");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:pPr> element.
 */
STATIC void
_chart_write_a_p_pr_rich(lxw_chart *self, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:pPr", NULL);

    /* Write the a:defRPr element. */
    _chart_write_a_def_rpr(self, font);

    lxw_xml_end_tag(self->file, "a:pPr");
}

/*
 * Write the <a:p> element.
 */
STATIC void
_chart_write_a_p_formula(lxw_chart *self, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:p", NULL);

    /* Write the a:pPr element. */
    _chart_write_a_p_pr_formula(self, font);

    /* Write the a:endParaRPr element. */
    _chart_write_a_end_para_rpr(self);

    lxw_xml_end_tag(self->file, "a:p");
}

/*
 * Write the <a:p> element for pie chart legends.
 */
STATIC void
_chart_write_a_p_pie(lxw_chart *self, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:p", NULL);

    /* Write the a:pPr element. */
    _chart_write_a_p_pr_pie(self, font);

    /* Write the a:endParaRPr element. */
    _chart_write_a_end_para_rpr(self);

    lxw_xml_end_tag(self->file, "a:p");
}

/*
 * Write the <a:p> element.
 */
STATIC void
_chart_write_a_p_rich(lxw_chart *self, char *name, lxw_chart_font *font)
{
    lxw_xml_start_tag(self->file, "a:p", NULL);

    /* Write the a:pPr element. */
    _chart_write_a_p_pr_rich(self, font);

    /* Write the a:r element. */
    _chart_write_a_r(self, name, font);

    lxw_xml_end_tag(self->file, "a:p");
}

/*
 * Write the <a:lstStyle> element.
 */
STATIC void
_chart_write_a_lst_style(lxw_chart *self)
{
    lxw_xml_empty_tag(self->file, "a:lstStyle", NULL);
}

/*
 * Write the <a:bodyPr> element.
 */
STATIC void
_chart_write_a_body_pr(lxw_chart *self, int32_t rotation,
                       uint8_t is_horizontal)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (rotation == 0 && is_horizontal)
        rotation = -5400000;

    if (rotation)
        LXW_PUSH_ATTRIBUTES_INT("rot", rotation);

    if (is_horizontal)
        LXW_PUSH_ATTRIBUTES_STR("vert", "horz");

    lxw_xml_empty_tag(self->file, "a:bodyPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:ptCount> element.
 */
STATIC void
_chart_write_pt_count(lxw_chart *self, uint16_t num_data_points)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", num_data_points);

    lxw_xml_empty_tag(self->file, "c:ptCount", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:v> element.
 */
STATIC void
_chart_write_v_num(lxw_chart *self, double number)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, number);

    lxw_xml_data_element(self->file, "c:v", data, NULL);
}

/*
 * Write the <c:v> element.
 */
STATIC void
_chart_write_v_str(lxw_chart *self, char *str)
{
    lxw_xml_data_element(self->file, "c:v", str, NULL);
}

/*
 * Write the <c:f> element.
 */
STATIC void
_chart_write_f(lxw_chart *self, char *formula)
{
    lxw_xml_data_element(self->file, "c:f", formula, NULL);
}

/*
 * Write the <c:pt> element.
 */
STATIC void
_chart_write_pt(lxw_chart *self, uint16_t index,
                lxw_series_data_point *data_point)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Ignore chart points that have no data. */
    if (data_point->no_data)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("idx", index);

    lxw_xml_start_tag(self->file, "c:pt", &attributes);

    if (data_point->is_string && data_point->string)
        _chart_write_v_str(self, data_point->string);
    else
        _chart_write_v_num(self, data_point->number);

    lxw_xml_end_tag(self->file, "c:pt");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:pt> element.
 */
STATIC void
_chart_write_num_pt(lxw_chart *self, uint16_t index,
                    lxw_series_data_point *data_point)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Ignore chart points that have no data. */
    if (data_point->no_data)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("idx", index);

    lxw_xml_start_tag(self->file, "c:pt", &attributes);

    _chart_write_v_num(self, data_point->number);

    lxw_xml_end_tag(self->file, "c:pt");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:formatCode> element.
 */
STATIC void
_chart_write_format_code(lxw_chart *self)
{
    lxw_xml_data_element(self->file, "c:formatCode", "General", NULL);
}

/*
 * Write the <c:numCache> element.
 */
STATIC void
_chart_write_num_cache(lxw_chart *self, lxw_series_range *range)
{
    lxw_series_data_point *data_point;
    uint16_t index = 0;

    lxw_xml_start_tag(self->file, "c:numCache", NULL);

    /* Write the c:formatCode element. */
    _chart_write_format_code(self);

    /* Write the c:ptCount element. */
    _chart_write_pt_count(self, range->num_data_points);

    STAILQ_FOREACH(data_point, range->data_cache, list_pointers) {
        /* Write the c:pt element. */
        _chart_write_num_pt(self, index, data_point);
        index++;
    }

    lxw_xml_end_tag(self->file, "c:numCache");
}

/*
 * Write the <c:strCache> element.
 */
STATIC void
_chart_write_str_cache(lxw_chart *self, lxw_series_range *range)
{
    lxw_series_data_point *data_point;
    uint16_t index = 0;

    lxw_xml_start_tag(self->file, "c:strCache", NULL);

    /* Write the c:ptCount element. */
    _chart_write_pt_count(self, range->num_data_points);

    STAILQ_FOREACH(data_point, range->data_cache, list_pointers) {
        /* Write the c:pt element. */
        _chart_write_pt(self, index, data_point);
        index++;
    }

    lxw_xml_end_tag(self->file, "c:strCache");
}

/*
 * Write the <c:numRef> element.
 */
STATIC void
_chart_write_num_ref(lxw_chart *self, lxw_series_range *range)
{
    lxw_xml_start_tag(self->file, "c:numRef", NULL);

    /* Write the c:f element. */
    _chart_write_f(self, range->formula);

    if (!STAILQ_EMPTY(range->data_cache)) {
        /* Write the c:numCache element. */
        _chart_write_num_cache(self, range);
    }

    lxw_xml_end_tag(self->file, "c:numRef");
}

/*
 * Write the <c:strRef> element.
 */
STATIC void
_chart_write_str_ref(lxw_chart *self, lxw_series_range *range)
{
    lxw_xml_start_tag(self->file, "c:strRef", NULL);

    /* Write the c:f element. */
    _chart_write_f(self, range->formula);

    if (!STAILQ_EMPTY(range->data_cache)) {
        /* Write the c:strCache element. */
        _chart_write_str_cache(self, range);
    }

    lxw_xml_end_tag(self->file, "c:strRef");
}

/*
 * Write the cached data elements.
 */
STATIC void
_chart_write_data_cache(lxw_chart *self, lxw_series_range *range,
                        uint8_t has_string_cache)
{
    if (has_string_cache) {
        /* Write the c:strRef element. */
        _chart_write_str_ref(self, range);
    }
    else {
        /* Write the c:numRef element. */
        _chart_write_num_ref(self, range);
    }
}

/*
 * Write the <c:tx> element with a simple value such as for series names.
 */
STATIC void
_chart_write_tx_value(lxw_chart *self, char *name)
{
    lxw_xml_start_tag(self->file, "c:tx", NULL);

    /* Write the c:v element. */
    _chart_write_v_str(self, name);

    lxw_xml_end_tag(self->file, "c:tx");
}

/*
 * Write the <c:tx> element with a simple value such as for series names.
 */
STATIC void
_chart_write_tx_formula(lxw_chart *self, lxw_chart_title *title)
{
    lxw_xml_start_tag(self->file, "c:tx", NULL);

    _chart_write_str_ref(self, title->range);

    lxw_xml_end_tag(self->file, "c:tx");
}

/*
 * Write the <c:txPr> element.
 */
STATIC void
_chart_write_tx_pr(lxw_chart *self, uint8_t is_horizontal,
                   lxw_chart_font *font)
{
    int32_t rotation = 0;

    if (font)
        rotation = font->rotation;

    lxw_xml_start_tag(self->file, "c:txPr", NULL);

    /* Write the a:bodyPr element. */
    _chart_write_a_body_pr(self, rotation, is_horizontal);

    /* Write the a:lstStyle element. */
    _chart_write_a_lst_style(self);

    /* Write the a:p element. */
    _chart_write_a_p_formula(self, font);

    lxw_xml_end_tag(self->file, "c:txPr");
}

/*
 * Write the <c:txPr> element for pie chart legends.
 */
STATIC void
_chart_write_tx_pr_pie(lxw_chart *self, uint8_t is_horizontal,
                       lxw_chart_font *font)
{
    int32_t rotation = 0;

    if (font)
        rotation = font->rotation;

    lxw_xml_start_tag(self->file, "c:txPr", NULL);

    /* Write the a:bodyPr element. */
    _chart_write_a_body_pr(self, rotation, is_horizontal);

    /* Write the a:lstStyle element. */
    _chart_write_a_lst_style(self);

    /* Write the a:p element. */
    _chart_write_a_p_pie(self, font);

    lxw_xml_end_tag(self->file, "c:txPr");
}

/*
 * Write the <c:txPr> element.
 */
STATIC void
_chart_write_axis_font(lxw_chart *self, lxw_chart_font *font)
{
    if (!font)
        return;

    lxw_xml_start_tag(self->file, "c:txPr", NULL);

    /* Write the a:bodyPr element. */
    _chart_write_a_body_pr(self, font->rotation, LXW_FALSE);

    /* Write the a:lstStyle element. */
    _chart_write_a_lst_style(self);

    lxw_xml_start_tag(self->file, "a:p", NULL);

    /* Write the a:pPr element. */
    _chart_write_a_p_pr_rich(self, font);

    /* Write the a:endParaRPr element. */
    _chart_write_a_end_para_rpr(self);

    lxw_xml_end_tag(self->file, "a:p");
    lxw_xml_end_tag(self->file, "c:txPr");
}

/*
 * Write the <c:rich> element.
 */
STATIC void
_chart_write_rich(lxw_chart *self, char *name, uint8_t is_horizontal,
                  lxw_chart_font *font)
{
    int32_t rotation = 0;

    if (font)
        rotation = font->rotation;

    lxw_xml_start_tag(self->file, "c:rich", NULL);

    /* Write the a:bodyPr element. */
    _chart_write_a_body_pr(self, rotation, is_horizontal);

    /* Write the a:lstStyle element. */
    _chart_write_a_lst_style(self);

    /* Write the a:p element. */
    _chart_write_a_p_rich(self, name, font);

    lxw_xml_end_tag(self->file, "c:rich");
}

/*
 * Write the <c:tx> element.
 */
STATIC void
_chart_write_tx_rich(lxw_chart *self, char *name, uint8_t is_horizontal,
                     lxw_chart_font *font)
{

    lxw_xml_start_tag(self->file, "c:tx", NULL);

    /* Write the c:rich element. */
    _chart_write_rich(self, name, is_horizontal, font);

    lxw_xml_end_tag(self->file, "c:tx");
}

/*
 * Write the <c:title> element for rich strings.
 */
STATIC void
_chart_write_title_rich(lxw_chart *self, lxw_chart_title *title)
{
    lxw_xml_start_tag(self->file, "c:title", NULL);

    /* Write the c:tx element. */
    _chart_write_tx_rich(self, title->name, title->is_horizontal,
                         title->font);

    /* Write the c:layout element. */
    _chart_write_layout(self);

    lxw_xml_end_tag(self->file, "c:title");
}

/*
 * Write the <c:title> element for a formula style title
 */
STATIC void
_chart_write_title_formula(lxw_chart *self, lxw_chart_title *title)
{
    lxw_xml_start_tag(self->file, "c:title", NULL);

    /* Write the c:tx element. */
    _chart_write_tx_formula(self, title);

    /* Write the c:layout element. */
    _chart_write_layout(self);

    /* Write the c:txPr element. */
    _chart_write_tx_pr(self, title->is_horizontal, title->font);

    lxw_xml_end_tag(self->file, "c:title");
}

/*
 * Write the <c:delete> element.
 */
STATIC void
_chart_write_delete(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:delete", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:autoTitleDeleted> element.
 */
STATIC void
_chart_write_auto_title_deleted(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:autoTitleDeleted", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:idx> element.
 */
STATIC void
_chart_write_idx(lxw_chart *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", index);

    lxw_xml_empty_tag(self->file, "c:idx", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:prstDash> element.
 */
STATIC void
_chart_write_a_prst_dash(lxw_chart *self, uint8_t dash_type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (dash_type == LXW_CHART_LINE_DASH_ROUND_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "sysDot");
    else if (dash_type == LXW_CHART_LINE_DASH_SQUARE_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "sysDash");
    else if (dash_type == LXW_CHART_LINE_DASH_DASH_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "dashDot");
    else if (dash_type == LXW_CHART_LINE_DASH_LONG_DASH)
        LXW_PUSH_ATTRIBUTES_STR("val", "lgDash");
    else if (dash_type == LXW_CHART_LINE_DASH_LONG_DASH_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "lgDashDot");
    else if (dash_type == LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "lgDashDotDot");
    else if (dash_type == LXW_CHART_LINE_DASH_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "dot");
    else if (dash_type == LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "sysDashDot");
    else if (dash_type == LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT)
        LXW_PUSH_ATTRIBUTES_STR("val", "sysDashDotDot");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "dash");

    lxw_xml_empty_tag(self->file, "a:prstDash", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:noFill> element.
 */
STATIC void
_chart_write_a_no_fill(lxw_chart *self)
{
    lxw_xml_empty_tag(self->file, "a:noFill", NULL);
}

/*
 * Write the <a:ln> element.
 */
STATIC void
_chart_write_a_ln(lxw_chart *self, lxw_chart_line *line)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    float width_flt;
    uint32_t width_int;

    LXW_INIT_ATTRIBUTES();

    /* Round width to nearest 0.25, like Excel. */
    width_flt = (float) (uint32_t) ((line->width + 0.125) * 4.0F) / 4.0F;

    /* Convert to internal units. */
    width_int = (uint32_t) (0.5 + (12700.0 * width_flt));

    if (width_int)
        LXW_PUSH_ATTRIBUTES_INT("w", width_int);

    lxw_xml_start_tag(self->file, "a:ln", &attributes);

    /* Write the line fill. */
    if (line->none) {
        /* Write the a:noFill element. */
        _chart_write_a_no_fill(self);
    }
    else if (line->has_color) {
        /* Write the a:solidFill element. */
        _chart_write_a_solid_fill(self, line->color, line->transparency);
    }

    /* Write the line/dash type. */
    if (line->dash_type) {
        /* Write the a:prstDash element. */
        _chart_write_a_prst_dash(self, line->dash_type);
    }

    lxw_xml_end_tag(self->file, "a:ln");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:fgClr> element.
 */
STATIC void
_chart_write_a_fg_clr(lxw_chart *self, lxw_color_t color)
{
    lxw_xml_start_tag(self->file, "a:fgClr", NULL);

    _chart_write_a_srgb_clr(self, color, LXW_FALSE);

    lxw_xml_end_tag(self->file, "a:fgClr");
}

/*
 * Write the <a:bgClr> element.
 */
STATIC void
_chart_write_a_bg_clr(lxw_chart *self, lxw_color_t color)
{
    lxw_xml_start_tag(self->file, "a:bgClr", NULL);

    _chart_write_a_srgb_clr(self, color, LXW_FALSE);

    lxw_xml_end_tag(self->file, "a:bgClr");
}

/*
 * Write the <a:pattFill> element.
 */
STATIC void
_chart_write_a_patt_fill(lxw_chart *self, lxw_chart_pattern *pattern)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (pattern->type == LXW_CHART_PATTERN_NONE)
        LXW_PUSH_ATTRIBUTES_STR("prst", "none");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_5)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct5");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_10)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct10");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_20)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct20");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_25)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct25");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_30)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct30");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_40)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct40");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_50)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct50");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_60)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct60");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_70)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct70");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_75)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct75");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_80)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct80");
    else if (pattern->type == LXW_CHART_PATTERN_PERCENT_90)
        LXW_PUSH_ATTRIBUTES_STR("prst", "pct90");
    else if (pattern->type == LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "ltDnDiag");
    else if (pattern->type == LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "ltUpDiag");
    else if (pattern->type == LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dkDnDiag");
    else if (pattern->type == LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dkUpDiag");
    else if (pattern->type == LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "wdDnDiag");
    else if (pattern->type == LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "wdUpDiag");
    else if (pattern->type == LXW_CHART_PATTERN_LIGHT_VERTICAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "ltVert");
    else if (pattern->type == LXW_CHART_PATTERN_LIGHT_HORIZONTAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "ltHorz");
    else if (pattern->type == LXW_CHART_PATTERN_NARROW_VERTICAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "narVert");
    else if (pattern->type == LXW_CHART_PATTERN_NARROW_HORIZONTAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "narHorz");
    else if (pattern->type == LXW_CHART_PATTERN_DARK_VERTICAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dkVert");
    else if (pattern->type == LXW_CHART_PATTERN_DARK_HORIZONTAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dkHorz");
    else if (pattern->type == LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dashDnDiag");
    else if (pattern->type == LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dashUpDiag");
    else if (pattern->type == LXW_CHART_PATTERN_DASHED_HORIZONTAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dashHorz");
    else if (pattern->type == LXW_CHART_PATTERN_DASHED_VERTICAL)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dashVert");
    else if (pattern->type == LXW_CHART_PATTERN_SMALL_CONFETTI)
        LXW_PUSH_ATTRIBUTES_STR("prst", "smConfetti");
    else if (pattern->type == LXW_CHART_PATTERN_LARGE_CONFETTI)
        LXW_PUSH_ATTRIBUTES_STR("prst", "lgConfetti");
    else if (pattern->type == LXW_CHART_PATTERN_ZIGZAG)
        LXW_PUSH_ATTRIBUTES_STR("prst", "zigZag");
    else if (pattern->type == LXW_CHART_PATTERN_WAVE)
        LXW_PUSH_ATTRIBUTES_STR("prst", "wave");
    else if (pattern->type == LXW_CHART_PATTERN_DIAGONAL_BRICK)
        LXW_PUSH_ATTRIBUTES_STR("prst", "diagBrick");
    else if (pattern->type == LXW_CHART_PATTERN_HORIZONTAL_BRICK)
        LXW_PUSH_ATTRIBUTES_STR("prst", "horzBrick");
    else if (pattern->type == LXW_CHART_PATTERN_WEAVE)
        LXW_PUSH_ATTRIBUTES_STR("prst", "weave");
    else if (pattern->type == LXW_CHART_PATTERN_PLAID)
        LXW_PUSH_ATTRIBUTES_STR("prst", "plaid");
    else if (pattern->type == LXW_CHART_PATTERN_DIVOT)
        LXW_PUSH_ATTRIBUTES_STR("prst", "divot");
    else if (pattern->type == LXW_CHART_PATTERN_DOTTED_GRID)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dotGrid");
    else if (pattern->type == LXW_CHART_PATTERN_DOTTED_DIAMOND)
        LXW_PUSH_ATTRIBUTES_STR("prst", "dotDmnd");
    else if (pattern->type == LXW_CHART_PATTERN_SHINGLE)
        LXW_PUSH_ATTRIBUTES_STR("prst", "shingle");
    else if (pattern->type == LXW_CHART_PATTERN_TRELLIS)
        LXW_PUSH_ATTRIBUTES_STR("prst", "trellis");
    else if (pattern->type == LXW_CHART_PATTERN_SPHERE)
        LXW_PUSH_ATTRIBUTES_STR("prst", "sphere");
    else if (pattern->type == LXW_CHART_PATTERN_SMALL_GRID)
        LXW_PUSH_ATTRIBUTES_STR("prst", "smGrid");
    else if (pattern->type == LXW_CHART_PATTERN_LARGE_GRID)
        LXW_PUSH_ATTRIBUTES_STR("prst", "lgGrid");
    else if (pattern->type == LXW_CHART_PATTERN_SMALL_CHECK)
        LXW_PUSH_ATTRIBUTES_STR("prst", "smCheck");
    else if (pattern->type == LXW_CHART_PATTERN_LARGE_CHECK)
        LXW_PUSH_ATTRIBUTES_STR("prst", "lgCheck");
    else if (pattern->type == LXW_CHART_PATTERN_OUTLINED_DIAMOND)
        LXW_PUSH_ATTRIBUTES_STR("prst", "openDmnd");
    else if (pattern->type == LXW_CHART_PATTERN_SOLID_DIAMOND)
        LXW_PUSH_ATTRIBUTES_STR("prst", "solidDmnd");
    else
        LXW_PUSH_ATTRIBUTES_STR("prst", "percent_50");

    lxw_xml_start_tag(self->file, "a:pattFill", &attributes);

    if (pattern->has_fg_color)
        _chart_write_a_fg_clr(self, pattern->fg_color);

    if (pattern->has_bg_color)
        _chart_write_a_bg_clr(self, pattern->bg_color);

    lxw_xml_end_tag(self->file, "a:pattFill");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:spPr> element.
 */
STATIC void
_chart_write_sp_pr(lxw_chart *self, lxw_chart_line *line,
                   lxw_chart_fill *fill, lxw_chart_pattern *pattern)
{
    if (!line && !fill && !pattern)
        return;

    lxw_xml_start_tag(self->file, "c:spPr", NULL);

    /* Write the series fill. Note: a pattern fill overrides a solid fill. */
    if (fill && !pattern) {
        if (fill->none) {
            /* Write the a:noFill element. */
            _chart_write_a_no_fill(self);
        }
        else {
            /* Write the a:solidFill element. */
            _chart_write_a_solid_fill(self, fill->color, fill->transparency);
        }
    }

    if (pattern) {
        /* Write the a:pattFill element. */
        _chart_write_a_patt_fill(self, pattern);
    }

    if (line) {
        /* Write the a:ln element. */
        _chart_write_a_ln(self, line);
    }

    lxw_xml_end_tag(self->file, "c:spPr");
}

/*
 * Write the <c:order> element.
 */
STATIC void
_chart_write_order(lxw_chart *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", index);

    lxw_xml_empty_tag(self->file, "c:order", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:axId> element.
 */
STATIC void
_chart_write_axis_id(lxw_chart *self, uint32_t axis_id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis_id);

    lxw_xml_empty_tag(self->file, "c:axId", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:axId> element.
 */
STATIC void
_chart_write_axis_ids(lxw_chart *self)
{
    if (!self->axis_id_1)
        _chart_add_axis_ids(self);

    _chart_write_axis_id(self, self->axis_id_1);
    _chart_write_axis_id(self, self->axis_id_2);
}

/*
 * Write the series name.
 */
STATIC void
_chart_write_series_name(lxw_chart *self, lxw_chart_series *series)
{
    if (series->title.name) {
        /* Write the c:tx element. */
        _chart_write_tx_value(self, series->title.name);
    }
    else if (series->title.range->formula) {
        /* Write the c:tx element. */
        _chart_write_tx_formula(self, &series->title);

    }
}

/*
 * Write the <c:majorTickMark> element.
 */
STATIC void
_chart_write_major_tick_mark(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->major_tick_mark)
        return;

    LXW_INIT_ATTRIBUTES();

    if (axis->major_tick_mark == LXW_CHART_AXIS_TICK_MARK_NONE)
        LXW_PUSH_ATTRIBUTES_STR("val", "none");
    else if (axis->major_tick_mark == LXW_CHART_AXIS_TICK_MARK_INSIDE)
        LXW_PUSH_ATTRIBUTES_STR("val", "in");
    else if (axis->major_tick_mark == LXW_CHART_AXIS_TICK_MARK_CROSSING)
        LXW_PUSH_ATTRIBUTES_STR("val", "cross");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "out");

    lxw_xml_empty_tag(self->file, "c:majorTickMark", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:minorTickMark> element.
 */
STATIC void
_chart_write_minor_tick_mark(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->minor_tick_mark)
        return;

    LXW_INIT_ATTRIBUTES();

    if (axis->minor_tick_mark == LXW_CHART_AXIS_TICK_MARK_NONE)
        LXW_PUSH_ATTRIBUTES_STR("val", "none");
    else if (axis->minor_tick_mark == LXW_CHART_AXIS_TICK_MARK_INSIDE)
        LXW_PUSH_ATTRIBUTES_STR("val", "in");
    else if (axis->minor_tick_mark == LXW_CHART_AXIS_TICK_MARK_CROSSING)
        LXW_PUSH_ATTRIBUTES_STR("val", "cross");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "out");

    lxw_xml_empty_tag(self->file, "c:minorTickMark", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:symbol> element.
 */
STATIC void
_chart_write_symbol(lxw_chart *self, uint8_t type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (type == LXW_CHART_MARKER_SQUARE)
        LXW_PUSH_ATTRIBUTES_STR("val", "square");
    else if (type == LXW_CHART_MARKER_DIAMOND)
        LXW_PUSH_ATTRIBUTES_STR("val", "diamond");
    else if (type == LXW_CHART_MARKER_TRIANGLE)
        LXW_PUSH_ATTRIBUTES_STR("val", "triangle");
    else if (type == LXW_CHART_MARKER_X)
        LXW_PUSH_ATTRIBUTES_STR("val", "x");
    else if (type == LXW_CHART_MARKER_STAR)
        LXW_PUSH_ATTRIBUTES_STR("val", "star");
    else if (type == LXW_CHART_MARKER_SHORT_DASH)
        LXW_PUSH_ATTRIBUTES_STR("val", "short_dash");
    else if (type == LXW_CHART_MARKER_LONG_DASH)
        LXW_PUSH_ATTRIBUTES_STR("val", "long_dash");
    else if (type == LXW_CHART_MARKER_CIRCLE)
        LXW_PUSH_ATTRIBUTES_STR("val", "circle");
    else if (type == LXW_CHART_MARKER_PLUS)
        LXW_PUSH_ATTRIBUTES_STR("val", "plus");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "none");

    lxw_xml_empty_tag(self->file, "c:symbol", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dPt> element.
 */
STATIC void
_chart_write_d_pt(lxw_chart *self, lxw_chart_point *point, uint16_t index)
{
    lxw_xml_start_tag(self->file, "c:dPt", NULL);

    /* Write the c:idx element. */
    _chart_write_idx(self, index);

    /* Scatter/Line charts have an additional marker for the point. */
    if (self->chart_group == LXW_CHART_SCATTER
        || self->chart_group == LXW_CHART_LINE)
        lxw_xml_start_tag(self->file, "c:marker", NULL);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, point->line, point->fill, point->pattern);

    if (self->chart_group == LXW_CHART_SCATTER
        || self->chart_group == LXW_CHART_LINE)
        lxw_xml_end_tag(self->file, "c:marker");

    lxw_xml_end_tag(self->file, "c:dPt");
}

/*
 * Write the <c:dPt> element.
 */
STATIC void
_chart_write_points(lxw_chart *self, lxw_chart_series *series)
{
    uint16_t index;

    for (index = 0; index < series->point_count; index++) {
        lxw_chart_point *point = &series->points[index];

        /* Ignore empty points. */
        if (!point->line && !point->fill && !point->pattern)
            continue;

        /* Write the c:dPt element. */
        _chart_write_d_pt(self, &series->points[index], index);
    }
}

/*
 * Write the <c:invertIfNegative> element.
 */
STATIC void
_chart_write_invert_if_negative(lxw_chart *self, lxw_chart_series *series)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!series->invert_if_negative)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:invertIfNegative", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showVal> element.
 */
STATIC void
_chart_write_show_val(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showVal", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showCatName> element.
 */
STATIC void
_chart_write_show_cat_name(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showCatName", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showSerName> element.
 */
STATIC void
_chart_write_show_ser_name(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showSerName", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showLeaderLines> element.
 */
STATIC void
_chart_write_show_leader_lines(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showLeaderLines", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dLblPos> element.
 */
STATIC void
_chart_write_d_lbl_pos(lxw_chart *self, uint8_t position)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (position == LXW_CHART_LABEL_POSITION_RIGHT)
        LXW_PUSH_ATTRIBUTES_STR("val", "r");
    else if (position == LXW_CHART_LABEL_POSITION_LEFT)
        LXW_PUSH_ATTRIBUTES_STR("val", "l");
    else if (position == LXW_CHART_LABEL_POSITION_ABOVE)
        LXW_PUSH_ATTRIBUTES_STR("val", "t");
    else if (position == LXW_CHART_LABEL_POSITION_BELOW)
        LXW_PUSH_ATTRIBUTES_STR("val", "b");
    else if (position == LXW_CHART_LABEL_POSITION_INSIDE_BASE)
        LXW_PUSH_ATTRIBUTES_STR("val", "inBase");
    else if (position == LXW_CHART_LABEL_POSITION_INSIDE_END)
        LXW_PUSH_ATTRIBUTES_STR("val", "inEnd");
    else if (position == LXW_CHART_LABEL_POSITION_OUTSIDE_END)
        LXW_PUSH_ATTRIBUTES_STR("val", "outEnd");
    else if (position == LXW_CHART_LABEL_POSITION_BEST_FIT)
        LXW_PUSH_ATTRIBUTES_STR("val", "bestFit");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "ctr");

    lxw_xml_empty_tag(self->file, "c:dLblPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:separator> element.
 */
STATIC void
_chart_write_separator(lxw_chart *self, uint8_t separator)
{
    if (separator == LXW_CHART_LABEL_SEPARATOR_SEMICOLON)
        lxw_xml_data_element(self->file, "c:separator", "; ", NULL);
    else if (separator == LXW_CHART_LABEL_SEPARATOR_PERIOD)
        lxw_xml_data_element(self->file, "c:separator", ". ", NULL);
    else if (separator == LXW_CHART_LABEL_SEPARATOR_NEWLINE)
        lxw_xml_data_element(self->file, "c:separator", "\n", NULL);
    else if (separator == LXW_CHART_LABEL_SEPARATOR_SPACE)
        lxw_xml_data_element(self->file, "c:separator", " ", NULL);
    else
        lxw_xml_data_element(self->file, "c:separator", ", ", NULL);
}

/*
 * Write the <c:showLegendKey> element.
 */
STATIC void
_chart_write_show_legend_key(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showLegendKey", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showPercent> element.
 */
STATIC void
_chart_write_show_percent(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showPercent", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:numFmt> element.
 */
STATIC void
_chart_write_label_num_fmt(lxw_chart *self, char *format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("formatCode", format);
    LXW_PUSH_ATTRIBUTES_STR("sourceLinked", "0");

    lxw_xml_empty_tag(self->file, "c:numFmt", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dLbls> element.
 */
STATIC void
_chart_write_d_lbls(lxw_chart *self, lxw_chart_series *series)
{
    if (!series->has_labels)
        return;

    lxw_xml_start_tag(self->file, "c:dLbls", NULL);

    /* Write the c:numFmt element. */
    if (series->label_num_format)
        _chart_write_label_num_fmt(self, series->label_num_format);

    if (series->label_font)
        _chart_write_tx_pr(self, LXW_FALSE, series->label_font);

    /* Write the c:dLblPos element. */
    if (series->label_position)
        _chart_write_d_lbl_pos(self, series->label_position);

    /* Write the c:showLegendKey element. */
    if (series->show_labels_legend)
        _chart_write_show_legend_key(self);

    /* Write the c:showVal element. */
    if (series->show_labels_value)
        _chart_write_show_val(self);

    /* Write the c:showCatName element. */
    if (series->show_labels_category)
        _chart_write_show_cat_name(self);

    /* Write the c:showSerName element. */
    if (series->show_labels_name)
        _chart_write_show_ser_name(self);

    /* Write the c:showPercent element. */
    if (series->show_labels_percent)
        _chart_write_show_percent(self);

    /* Write the c:separator element. */
    if (series->label_separator)
        _chart_write_separator(self, series->label_separator);

    /* Write the c:showLeaderLines element. */
    if (series->show_labels_leader)
        _chart_write_show_leader_lines(self);

    lxw_xml_end_tag(self->file, "c:dLbls");
}

/*
 * Write the <c:intercept> element.
 */
STATIC void
_chart_write_intercept(lxw_chart *self, double value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", value);

    lxw_xml_empty_tag(self->file, "c:intercept", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dispRSqr> element.
 */
STATIC void
_chart_write_disp_rsqr(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:dispRSqr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:trendlineLbl> element.
 */
STATIC void
_chart_write_trendline_lbl(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    lxw_xml_start_tag(self->file, "c:trendlineLbl", NULL);

    lxw_xml_empty_tag(self->file, "c:layout", NULL);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("formatCode", "General");
    LXW_PUSH_ATTRIBUTES_INT("sourceLinked", 0);

    lxw_xml_empty_tag(self->file, "c:numFmt", &attributes);

    lxw_xml_end_tag(self->file, "c:trendlineLbl");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dispEq> element.
 */
STATIC void
_chart_write_disp_eq(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:dispEq", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:period> element.
 */
STATIC void
_chart_write_period(lxw_chart *self, uint8_t value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", value);

    lxw_xml_empty_tag(self->file, "c:period", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:forward> element.
 */
STATIC void
_chart_write_forward(lxw_chart *self, double value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", value);

    lxw_xml_empty_tag(self->file, "c:forward", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:backward> element.
 */
STATIC void
_chart_write_backward(lxw_chart *self, double value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", value);

    lxw_xml_empty_tag(self->file, "c:backward", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:name> element.
 */
STATIC void
_chart_write_name(lxw_chart *self, char *name)
{
    lxw_xml_data_element(self->file, "c:name", name, NULL);
}

/*
 * Write the <c:trendlineType> element.
 */
STATIC void
_chart_write_trendline_type(lxw_chart *self, uint8_t type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (type == LXW_CHART_TRENDLINE_TYPE_LOG)
        LXW_PUSH_ATTRIBUTES_STR("val", "log");
    else if (type == LXW_CHART_TRENDLINE_TYPE_POLY)
        LXW_PUSH_ATTRIBUTES_STR("val", "poly");
    else if (type == LXW_CHART_TRENDLINE_TYPE_POWER)
        LXW_PUSH_ATTRIBUTES_STR("val", "power");
    else if (type == LXW_CHART_TRENDLINE_TYPE_EXP)
        LXW_PUSH_ATTRIBUTES_STR("val", "exp");
    else if (type == LXW_CHART_TRENDLINE_TYPE_AVERAGE)
        LXW_PUSH_ATTRIBUTES_STR("val", "movingAvg");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "linear");

    lxw_xml_empty_tag(self->file, "c:trendlineType", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:trendline> element.
 */
STATIC void
_chart_write_trendline(lxw_chart *self, lxw_chart_series *series)
{
    if (!series->has_trendline)
        return;

    lxw_xml_start_tag(self->file, "c:trendline", NULL);

    /* Write the c:name element. */
    if (series->trendline_name)
        _chart_write_name(self, series->trendline_name);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, series->trendline_line, NULL, NULL);

    /* Write the c:trendlineType element. */
    _chart_write_trendline_type(self, series->trendline_type);

    /* Write the c:order element. */
    if (series->trendline_type == LXW_CHART_TRENDLINE_TYPE_POLY
        && series->trendline_value >= 2) {

        _chart_write_order(self, series->trendline_value);
    }

    /* Write the c:period element. */
    if (series->trendline_type == LXW_CHART_TRENDLINE_TYPE_AVERAGE
        && series->trendline_value >= 2) {

        _chart_write_period(self, series->trendline_value);
    }

    if (series->has_trendline_forecast) {
        /* Write the c:forward element. */
        _chart_write_forward(self, series->trendline_forward);

        /* Write the c:backward element. */
        _chart_write_backward(self, series->trendline_backward);
    }

    /* Write the c:intercept element. */
    if (series->has_trendline_intercept)
        _chart_write_intercept(self, series->trendline_intercept);

    /* Write the c:dispRSqr element. */
    if (series->has_trendline_r_squared)
        _chart_write_disp_rsqr(self);

    if (series->has_trendline_equation) {
        /* Write the c:dispEq element. */
        _chart_write_disp_eq(self);

        /* Write the c:trendlineLbl element. */
        _chart_write_trendline_lbl(self);

    }

    lxw_xml_end_tag(self->file, "c:trendline");
}

/*
 * Write the <c:val> element.
 */
STATIC void
_chart_write_error_val(lxw_chart *self, double value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", value);

    lxw_xml_empty_tag(self->file, "c:val", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:noEndCap> element.
 */
STATIC void
_chart_write_no_end_cap(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:noEndCap", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:errValType> element.
 */
STATIC void
_chart_write_err_val_type(lxw_chart *self, uint8_t type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (type == LXW_CHART_ERROR_BAR_TYPE_FIXED)
        LXW_PUSH_ATTRIBUTES_STR("val", "fixedVal");
    else if (type == LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE)
        LXW_PUSH_ATTRIBUTES_STR("val", "percentage");
    else if (type == LXW_CHART_ERROR_BAR_TYPE_STD_DEV)
        LXW_PUSH_ATTRIBUTES_STR("val", "stdDev");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "stdErr");

    lxw_xml_empty_tag(self->file, "c:errValType", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:errBarType> element.
 */
STATIC void
_chart_write_err_bar_type(lxw_chart *self, uint8_t direction)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (direction == LXW_CHART_ERROR_BAR_DIR_PLUS)
        LXW_PUSH_ATTRIBUTES_STR("val", "plus");
    else if (direction == LXW_CHART_ERROR_BAR_DIR_MINUS)
        LXW_PUSH_ATTRIBUTES_STR("val", "minus");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "both");

    lxw_xml_empty_tag(self->file, "c:errBarType", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:errDir> element.
 */
STATIC void
_chart_write_err_dir(lxw_chart *self, uint8_t is_x)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (is_x)
        LXW_PUSH_ATTRIBUTES_STR("val", "x");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "y");

    lxw_xml_empty_tag(self->file, "c:errDir", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:errBars> element.
 */
STATIC void
_chart_write_err_bars(lxw_chart *self, lxw_series_error_bars *error_bars)
{
    if (!error_bars->is_set)
        return;

    lxw_xml_start_tag(self->file, "c:errBars", NULL);

    /* Write the c:errDir element, except for Column/Bar charts. */
    if (error_bars->chart_group != LXW_CHART_BAR
        && error_bars->chart_group != LXW_CHART_COLUMN) {

        _chart_write_err_dir(self, error_bars->is_x);
    }

    /* Write the c:errBarType element. */
    _chart_write_err_bar_type(self, error_bars->direction);

    /* Write the c:errValType element. */
    _chart_write_err_val_type(self, error_bars->type);

    /* Write the c:noEndCap element. */
    if (error_bars->endcap == LXW_CHART_ERROR_BAR_NO_CAP)
        _chart_write_no_end_cap(self);

    /* Write the c:val element. */
    if (error_bars->has_value)
        _chart_write_error_val(self, error_bars->value);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, error_bars->line, NULL, NULL);

    lxw_xml_end_tag(self->file, "c:errBars");
}

/*
 * Write the <c:errBars> element.
 */
STATIC void
_chart_write_error_bars(lxw_chart *self, lxw_chart_series *series)
{
    _chart_write_err_bars(self, series->x_error_bars);
    _chart_write_err_bars(self, series->y_error_bars);
}

/*
 * Write the <c:size> element.
 */
STATIC void
_chart_write_marker_size(lxw_chart *self, uint8_t size)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", size);

    lxw_xml_empty_tag(self->file, "c:size", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:marker> element.
 */
STATIC void
_chart_write_marker(lxw_chart *self, lxw_chart_marker *marker)
{
    /* If there isn't a user defined marker use the default, if this chart
     * type one. The default usually turns the marker off. */
    if (!marker)
        marker = self->default_marker;

    if (!marker)
        return;

    if (marker->type == LXW_CHART_MARKER_AUTOMATIC)
        return;

    lxw_xml_start_tag(self->file, "c:marker", NULL);

    /* Write the c:symbol element. */
    _chart_write_symbol(self, marker->type);

    /* Write the c:size element. */
    if (marker->size)
        _chart_write_marker_size(self, marker->size);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, marker->line, marker->fill, marker->pattern);

    lxw_xml_end_tag(self->file, "c:marker");
}

/*
 * Write the <c:marker> element.
 */
STATIC void
_chart_write_marker_value(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:marker", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:smooth> element.
 */
STATIC void
_chart_write_smooth(lxw_chart *self, uint8_t smooth)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!smooth)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:smooth", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:scatterStyle> element.
 */
STATIC void
_chart_write_scatter_style(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (self->type == LXW_CHART_SCATTER_SMOOTH
        || self->type == LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS)
        LXW_PUSH_ATTRIBUTES_STR("val", "smoothMarker");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "lineMarker");

    lxw_xml_empty_tag(self->file, "c:scatterStyle", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:cat> element.
 */
STATIC void
_chart_write_cat(lxw_chart *self, lxw_chart_series *series)
{
    uint8_t has_string_cache = series->categories->has_string_cache;

    /* Ignore <c:cat> elements for charts without category values. */
    if (!series->categories->formula)
        return;

    self->cat_has_num_fmt = !has_string_cache;

    lxw_xml_start_tag(self->file, "c:cat", NULL);

    /* Write the c:numRef element. */
    _chart_write_data_cache(self, series->categories, has_string_cache);

    lxw_xml_end_tag(self->file, "c:cat");
}

/*
 * Write the <c:xVal> element.
 */
STATIC void
_chart_write_x_val(lxw_chart *self, lxw_chart_series *series)
{
    uint8_t has_string_cache = series->categories->has_string_cache;

    lxw_xml_start_tag(self->file, "c:xVal", NULL);

    /* Write the data cache elements. */
    _chart_write_data_cache(self, series->categories, has_string_cache);

    lxw_xml_end_tag(self->file, "c:xVal");
}

/*
 * Write the <c:val> element.
 */
STATIC void
_chart_write_val(lxw_chart *self, lxw_chart_series *series)
{
    lxw_xml_start_tag(self->file, "c:val", NULL);

    /* Write the data cache elements. The string_cache is set to false since
     * this should always be a number series. */
    _chart_write_data_cache(self, series->values, LXW_FALSE);

    lxw_xml_end_tag(self->file, "c:val");
}

/*
 * Write the <c:yVal> element.
 */
STATIC void
_chart_write_y_val(lxw_chart *self, lxw_chart_series *series)
{
    lxw_xml_start_tag(self->file, "c:yVal", NULL);

    /* Write the data cache elements. The string_cache is set to false since
     * this should always be a number series. */
    _chart_write_data_cache(self, series->values, LXW_FALSE);

    lxw_xml_end_tag(self->file, "c:yVal");
}

/*
 * Write the <c:ser> element.
 */
STATIC void
_chart_write_ser(lxw_chart *self, lxw_chart_series *series)
{
    uint16_t index = self->series_index++;

    lxw_xml_start_tag(self->file, "c:ser", NULL);

    /* Write the c:idx element. */
    _chart_write_idx(self, index);

    /* Write the c:order element. */
    _chart_write_order(self, index);

    /* Write the series name. */
    _chart_write_series_name(self, series);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, series->line, series->fill, series->pattern);

    /* Write the c:marker element. */
    _chart_write_marker(self, series->marker);

    /* Write the c:invertIfNegative element. */
    _chart_write_invert_if_negative(self, series);

    /* Write the char points. */
    _chart_write_points(self, series);

    /* Write the c:dLbls element. */
    _chart_write_d_lbls(self, series);

    /* Write the c:trendline element. */
    _chart_write_trendline(self, series);

    /* Write the c:errBars element. */
    _chart_write_error_bars(self, series);

    /* Write the c:cat element. */
    _chart_write_cat(self, series);

    /* Write the c:val element. */
    _chart_write_val(self, series);

    /* Write the c:smooth element. */
    if (self->chart_group == LXW_CHART_SCATTER
        || self->chart_group == LXW_CHART_LINE)
        _chart_write_smooth(self, series->smooth);

    lxw_xml_end_tag(self->file, "c:ser");
}

/*
 * Write the <c:ser> element but with c:xVal/c:yVal instead of c:cat/c:val
 * elements.
 */
STATIC void
_chart_write_xval_ser(lxw_chart *self, lxw_chart_series *series)
{
    uint16_t index = self->series_index++;

    lxw_xml_start_tag(self->file, "c:ser", NULL);

    /* Write the c:idx element. */
    _chart_write_idx(self, index);

    /* Write the c:order element. */
    _chart_write_order(self, index);

    /* Write the series name. */
    _chart_write_series_name(self, series);

    /* Write the c:spPr element. */
    _chart_write_sp_pr(self, series->line, series->fill, series->pattern);

    /* Write the c:marker element. */
    _chart_write_marker(self, series->marker);

    /* Write the char points. */
    _chart_write_points(self, series);

    /* Write the c:dLbls element. */
    _chart_write_d_lbls(self, series);

    /* Write the c:trendline element. */
    _chart_write_trendline(self, series);

    /* Write the c:errBars element. */
    _chart_write_error_bars(self, series);

    /* Write the c:xVal element. */
    _chart_write_x_val(self, series);

    /* Write the yVal element. */
    _chart_write_y_val(self, series);

    /* Write the c:smooth element. */
    _chart_write_smooth(self, series->smooth);

    lxw_xml_end_tag(self->file, "c:ser");
}

/*
 * Write the <c:orientation> element.
 */
STATIC void
_chart_write_orientation(lxw_chart *self, uint8_t reverse)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    if (reverse)
        LXW_PUSH_ATTRIBUTES_STR("val", "maxMin");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "minMax");

    lxw_xml_empty_tag(self->file, "c:orientation", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:max> element.
 */
STATIC void
_chart_write_max(lxw_chart *self, double max)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", max);

    lxw_xml_empty_tag(self->file, "c:max", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:min> element.
 */
STATIC void
_chart_write_min(lxw_chart *self, double min)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", min);

    lxw_xml_empty_tag(self->file, "c:min", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:logBase> element.
 */
STATIC void
_chart_write_log_base(lxw_chart *self, uint16_t log_base)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!log_base)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", log_base);

    lxw_xml_empty_tag(self->file, "c:logBase", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:scaling> element.
 */
STATIC void
_chart_write_scaling(lxw_chart *self, uint8_t reverse,
                     uint8_t has_min, double min,
                     uint8_t has_max, double max, uint16_t log_base)
{
    lxw_xml_start_tag(self->file, "c:scaling", NULL);

    /* Write the c:logBase element. */
    _chart_write_log_base(self, log_base);

    /* Write the c:orientation element. */
    _chart_write_orientation(self, reverse);

    if (has_max) {
        /* Write the c:max element. */
        _chart_write_max(self, max);
    }

    if (has_min) {
        /* Write the c:min element. */
        _chart_write_min(self, min);
    }

    lxw_xml_end_tag(self->file, "c:scaling");
}

/*
 * Write the <c:axPos> element.
 */
STATIC void
_chart_write_axis_pos(lxw_chart *self, uint8_t position, uint8_t reverse)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    /* Reverse the axis direction if required. */
    position ^= reverse;

    if (position == LXW_CHART_AXIS_RIGHT)
        LXW_PUSH_ATTRIBUTES_STR("val", "r");
    else if (position == LXW_CHART_AXIS_LEFT)
        LXW_PUSH_ATTRIBUTES_STR("val", "l");
    else if (position == LXW_CHART_AXIS_TOP)
        LXW_PUSH_ATTRIBUTES_STR("val", "t");
    else if (position == LXW_CHART_AXIS_BOTTOM)
        LXW_PUSH_ATTRIBUTES_STR("val", "b");

    lxw_xml_empty_tag(self->file, "c:axPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:tickLblPos> element.
 */
STATIC void
_chart_write_tick_label_pos(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (axis->label_position == LXW_CHART_AXIS_LABEL_POSITION_HIGH)
        LXW_PUSH_ATTRIBUTES_STR("val", "high");
    else if (axis->label_position == LXW_CHART_AXIS_LABEL_POSITION_LOW)
        LXW_PUSH_ATTRIBUTES_STR("val", "low");
    else if (axis->label_position == LXW_CHART_AXIS_LABEL_POSITION_NONE)
        LXW_PUSH_ATTRIBUTES_STR("val", "none");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "nextTo");

    lxw_xml_empty_tag(self->file, "c:tickLblPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crossAx> element.
 */
STATIC void
_chart_write_cross_axis(lxw_chart *self, uint32_t axis_id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis_id);

    lxw_xml_empty_tag(self->file, "c:crossAx", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crosses> element.
 */
STATIC void
_chart_write_crosses(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (axis->crossing_max)
        LXW_PUSH_ATTRIBUTES_STR("val", "max");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "autoZero");

    lxw_xml_empty_tag(self->file, "c:crosses", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crossesAt> element.
 */
STATIC void
_chart_write_crosses_at(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_DBL("val", axis->crossing);

    lxw_xml_empty_tag(self->file, "c:crossesAt", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:auto> element.
 */
STATIC void
_chart_write_auto(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:auto", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lblAlgn> element.
 */
STATIC void
_chart_write_label_align(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "ctr");

    lxw_xml_empty_tag(self->file, "c:lblAlgn", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:tickLblSkip> element.
 */
STATIC void
_chart_write_tick_label_skip(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->interval_unit)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis->interval_unit);

    lxw_xml_empty_tag(self->file, "c:tickLblSkip", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:tickMarkSkip> element.
 */
STATIC void
_chart_write_tick_mark_skip(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->interval_tick)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis->interval_tick);

    lxw_xml_empty_tag(self->file, "c:tickMarkSkip", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:majorUnit> element.
 */
STATIC void
_chart_write_major_unit(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->has_major_unit)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", axis->major_unit);

    lxw_xml_empty_tag(self->file, "c:majorUnit", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:minorUnit> element.
 */
STATIC void
_chart_write_minor_unit(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->has_minor_unit)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("val", axis->minor_unit);

    lxw_xml_empty_tag(self->file, "c:minorUnit", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dispUnits> element.
 */
STATIC void
_chart_write_disp_units(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->display_units)
        return;

    LXW_INIT_ATTRIBUTES();

    lxw_xml_start_tag(self->file, "c:dispUnits", NULL);

    if (axis->display_units == LXW_CHART_AXIS_UNITS_HUNDREDS)
        LXW_PUSH_ATTRIBUTES_STR("val", "hundreds");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_THOUSANDS)
        LXW_PUSH_ATTRIBUTES_STR("val", "thousands");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_TEN_THOUSANDS)
        LXW_PUSH_ATTRIBUTES_STR("val", "tenThousands");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS)
        LXW_PUSH_ATTRIBUTES_STR("val", "hundredThousands");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_MILLIONS)
        LXW_PUSH_ATTRIBUTES_STR("val", "millions");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_TEN_MILLIONS)
        LXW_PUSH_ATTRIBUTES_STR("val", "tenMillions");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS)
        LXW_PUSH_ATTRIBUTES_STR("val", "hundredMillions");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_BILLIONS)
        LXW_PUSH_ATTRIBUTES_STR("val", "billions");
    else if (axis->display_units == LXW_CHART_AXIS_UNITS_TRILLIONS)
        LXW_PUSH_ATTRIBUTES_STR("val", "trillions");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "hundreds");

    lxw_xml_empty_tag(self->file, "c:builtInUnit", &attributes);

    if (axis->display_units_visible) {
        lxw_xml_start_tag(self->file, "c:dispUnitsLbl", NULL);
        lxw_xml_empty_tag(self->file, "c:layout", NULL);
        lxw_xml_end_tag(self->file, "c:dispUnitsLbl");
    }

    lxw_xml_end_tag(self->file, "c:dispUnits");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lblOffset> element.
 */
STATIC void
_chart_write_label_offset(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "100");

    lxw_xml_empty_tag(self->file, "c:lblOffset", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:majorGridlines> element.
 */
STATIC void
_chart_write_major_gridlines(lxw_chart *self, lxw_chart_axis *axis)
{
    if (!axis->major_gridlines.visible)
        return;

    if (axis->major_gridlines.line) {
        lxw_xml_start_tag(self->file, "c:majorGridlines", NULL);

        /* Write the c:spPr element for the axis line. */
        _chart_write_sp_pr(self, axis->major_gridlines.line, NULL, NULL);

        lxw_xml_end_tag(self->file, "c:majorGridlines");
    }
    else {
        lxw_xml_empty_tag(self->file, "c:majorGridlines", NULL);
    }
}

/*
 * Write the <c:minorGridlines> element.
 */
STATIC void
_chart_write_minor_gridlines(lxw_chart *self, lxw_chart_axis *axis)
{
    if (!axis->minor_gridlines.visible)
        return;

    if (axis->minor_gridlines.line) {
        lxw_xml_start_tag(self->file, "c:minorGridlines", NULL);

        /* Write the c:spPr element for the axis line. */
        _chart_write_sp_pr(self, axis->minor_gridlines.line, NULL, NULL);

        lxw_xml_end_tag(self->file, "c:minorGridlines");

    }
    else {
        lxw_xml_empty_tag(self->file, "c:minorGridlines", NULL);
    }
}

/*
 * Write the <c:numberFormat> element. Note: It is assumed that if a user
 * defined number format is supplied (i.e., non-default) then the sourceLinked
 * attribute is 0. The user can override this if required.
 */
STATIC void
_chart_write_number_format(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char *num_format;
    uint8_t source_linked = 1;

    /* Set the number format to the axis default if not set. */
    if (axis->num_format)
        num_format = axis->num_format;
    else
        num_format = axis->default_num_format;

    /* Check if a user defined number format has been set. */
    if (strcmp(num_format, axis->default_num_format))
        source_linked = 0;

    /* Allow override of sourceLinked. */
    if (axis->source_linked)
        source_linked = 1;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("formatCode", num_format);
    LXW_PUSH_ATTRIBUTES_INT("sourceLinked", source_linked);

    lxw_xml_empty_tag(self->file, "c:numFmt", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:numFmt> element. Special case handler for category axes which
 * don't always have a number format.
 */
STATIC void
_chart_write_cat_number_format(lxw_chart *self, lxw_chart_axis *axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char *num_format;
    uint8_t source_linked = 1;
    uint8_t default_format = LXW_TRUE;

    /* Set the number format to the axis default if not set. */
    if (axis->num_format)
        num_format = axis->num_format;
    else
        num_format = axis->default_num_format;

    /* Check if a user defined number format has been set. */
    if (strcmp(num_format, axis->default_num_format)) {
        source_linked = 0;
        default_format = LXW_FALSE;
    }

    /* Allow override of sourceLinked. */
    if (axis->source_linked)
        source_linked = 1;

    /* Skip if cat doesn't have a num format (unless it is non-default). */
    if (!self->cat_has_num_fmt && default_format)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("formatCode", num_format);
    LXW_PUSH_ATTRIBUTES_INT("sourceLinked", source_linked);

    lxw_xml_empty_tag(self->file, "c:numFmt", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crossBetween> element.
 */
STATIC void
_chart_write_cross_between(lxw_chart *self, uint8_t position)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!position)
        position = self->default_cross_between;

    LXW_INIT_ATTRIBUTES();

    if (position == LXW_CHART_AXIS_POSITION_ON_TICK)
        LXW_PUSH_ATTRIBUTES_STR("val", "midCat");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "between");

    lxw_xml_empty_tag(self->file, "c:crossBetween", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:overlay> element.
 */
STATIC void
_chart_write_overlay(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:overlay", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:legendPos> element.
 */
STATIC void
_chart_write_legend_pos(lxw_chart *self, char *position)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", position);

    lxw_xml_empty_tag(self->file, "c:legendPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:legendEntry> element.
 */
STATIC void
_chart_write_legend_entry(lxw_chart *self, uint16_t index)
{
    lxw_xml_start_tag(self->file, "c:legendEntry", NULL);

    /* Write the c:idx element. */
    _chart_write_idx(self, self->delete_series[index]);

    /* Write the c:delete element. */
    _chart_write_delete(self);

    lxw_xml_end_tag(self->file, "c:legendEntry");
}

/*
 * Write the <c:legend> element.
 */
STATIC void
_chart_write_legend(lxw_chart *self)
{
    uint8_t has_overlay = LXW_FALSE;
    uint16_t index;

    if (self->legend.position == LXW_CHART_LEGEND_NONE)
        return;

    lxw_xml_start_tag(self->file, "c:legend", NULL);

    /* Write the c:legendPos element. */
    switch (self->legend.position) {
        case LXW_CHART_LEGEND_LEFT:
            _chart_write_legend_pos(self, "l");
            break;
        case LXW_CHART_LEGEND_TOP:
            _chart_write_legend_pos(self, "t");
            break;
        case LXW_CHART_LEGEND_BOTTOM:
            _chart_write_legend_pos(self, "b");
            break;
        case LXW_CHART_LEGEND_OVERLAY_RIGHT:
            _chart_write_legend_pos(self, "r");
            has_overlay = LXW_TRUE;
            break;
        case LXW_CHART_LEGEND_OVERLAY_LEFT:
            _chart_write_legend_pos(self, "l");
            has_overlay = LXW_TRUE;
            break;
        default:
            _chart_write_legend_pos(self, "r");
    }

    /* Remove series labels from the legend. */
    for (index = 0; index < self->delete_series_count; index++) {
        /* Write the c:legendEntry element. */
        _chart_write_legend_entry(self, index);
    }

    /* Write the c:layout element. */
    _chart_write_layout(self);

    if (self->chart_group == LXW_CHART_PIE
        || self->chart_group == LXW_CHART_DOUGHNUT) {
        /* Write the c:overlay element. */
        if (has_overlay)
            _chart_write_overlay(self);

        /* Write the c:txPr element for Pie/Doughnut charts. */
        _chart_write_tx_pr_pie(self, LXW_FALSE, self->legend.font);
    }
    else {
        /* Write the c:txPr element for all other charts. */
        if (self->legend.font)
            _chart_write_tx_pr(self, LXW_FALSE, self->legend.font);

        /* Write the c:overlay element. */
        if (has_overlay)
            _chart_write_overlay(self);
    }

    lxw_xml_end_tag(self->file, "c:legend");
}

/*
 * Write the <c:plotVisOnly> element.
 */
STATIC void
_chart_write_plot_vis_only(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (self->show_hidden_data)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:plotVisOnly", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:headerFooter> element.
 */
STATIC void
_chart_write_header_footer(lxw_chart *self)
{
    lxw_xml_empty_tag(self->file, "c:headerFooter", NULL);
}

/*
 * Write the <c:pageMargins> element.
 */
STATIC void
_chart_write_page_margins(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("b", "0.75");
    LXW_PUSH_ATTRIBUTES_STR("l", "0.7");
    LXW_PUSH_ATTRIBUTES_STR("r", "0.7");
    LXW_PUSH_ATTRIBUTES_STR("t", "0.75");
    LXW_PUSH_ATTRIBUTES_STR("header", "0.3");
    LXW_PUSH_ATTRIBUTES_STR("footer", "0.3");

    lxw_xml_empty_tag(self->file, "c:pageMargins", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:pageSetup> element.
 */
STATIC void
_chart_write_page_setup(lxw_chart *self)
{
    lxw_xml_empty_tag(self->file, "c:pageSetup", NULL);
}

/*
 * Write the <c:printSettings> element.
 */
STATIC void
_chart_write_print_settings(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:printSettings", NULL);

    /* Write the c:headerFooter element. */
    _chart_write_header_footer(self);

    /* Write the c:pageMargins element. */
    _chart_write_page_margins(self);

    /* Write the c:pageSetup element. */
    _chart_write_page_setup(self);

    lxw_xml_end_tag(self->file, "c:printSettings");
}

/*
 * Write the <c:overlap> element.
 */
STATIC void
_chart_write_overlap(lxw_chart *self, int8_t overlap)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!overlap)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", overlap);

    lxw_xml_empty_tag(self->file, "c:overlap", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:gapWidth> element.
 */
STATIC void
_chart_write_gap_width(lxw_chart *self, uint16_t gap)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (gap == LXW_CHART_DEFAULT_GAP)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", gap);

    lxw_xml_empty_tag(self->file, "c:gapWidth", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dispBlanksAs> element.
 */
STATIC void
_chart_write_disp_blanks_as(lxw_chart *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (self->show_blanks_as != LXW_CHART_BLANKS_AS_ZERO
        && self->show_blanks_as != LXW_CHART_BLANKS_AS_CONNECTED)
        return;

    LXW_INIT_ATTRIBUTES();

    if (self->show_blanks_as == LXW_CHART_BLANKS_AS_ZERO)
        LXW_PUSH_ATTRIBUTES_STR("val", "zero");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "span");

    lxw_xml_empty_tag(self->file, "c:dispBlanksAs", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showHorzBorder> element.
 */
STATIC void
_chart_write_show_horz_border(lxw_chart *self, uint8_t value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!value)
        return;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showHorzBorder", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showVertBorder> element.
 */
STATIC void
_chart_write_show_vert_border(lxw_chart *self, uint8_t value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!value)
        return;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showVertBorder", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showOutline> element.
 */
STATIC void
_chart_write_show_outline(lxw_chart *self, uint8_t value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!value)
        return;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showOutline", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:showKeys> element.
 */
STATIC void
_chart_write_show_keys(lxw_chart *self, uint8_t value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!value)
        return;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag(self->file, "c:showKeys", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:dTable> element.
 */
STATIC void
_chart_write_d_table(lxw_chart *self)
{
    if (!self->has_table)
        return;

    lxw_xml_start_tag(self->file, "c:dTable", NULL);

    /* Write the c:showHorzBorder element. */
    _chart_write_show_horz_border(self, self->has_table_horizontal);

    /* Write the c:showVertBorder element. */
    _chart_write_show_vert_border(self, self->has_table_vertical);

    /* Write the c:showOutline element. */
    _chart_write_show_outline(self, self->has_table_outline);

    /* Write the c:showKeys element. */
    _chart_write_show_keys(self, self->has_table_legend_keys);

    /* Write the c:txPr element. */
    if (self->table_font)
        _chart_write_tx_pr(self, LXW_FALSE, self->table_font);

    lxw_xml_end_tag(self->file, "c:dTable");
}

/*
 * Write the <c:upBars> element.
 */
STATIC void
_chart_write_up_bars(lxw_chart *self, lxw_chart_line *line,
                     lxw_chart_fill *fill)
{
    if (line || fill) {
        lxw_xml_start_tag(self->file, "c:upBars", NULL);

        /* Write the c:spPr element. */
        _chart_write_sp_pr(self, line, fill, NULL);

        lxw_xml_end_tag(self->file, "c:upBars");
    }
    else {
        lxw_xml_empty_tag(self->file, "c:upBars", NULL);
    }
}

/*
 * Write the <c:downBars> element.
 */
STATIC void
_chart_write_down_bars(lxw_chart *self, lxw_chart_line *line,
                       lxw_chart_fill *fill)
{
    if (line || fill) {
        lxw_xml_start_tag(self->file, "c:downBars", NULL);

        /* Write the c:spPr element. */
        _chart_write_sp_pr(self, line, fill, NULL);

        lxw_xml_end_tag(self->file, "c:downBars");
    }
    else {
        lxw_xml_empty_tag(self->file, "c:downBars", NULL);
    }
}

/*
 * Write the <c:upDownBars> element.
 */
STATIC void
_chart_write_up_down_bars(lxw_chart *self)
{
    if (!self->has_up_down_bars)
        return;

    lxw_xml_start_tag(self->file, "c:upDownBars", NULL);

    /* Write the c:gapWidth element. */
    _chart_write_gap_width(self, 150);

    /* Write the c:upBars element. */
    _chart_write_up_bars(self, self->up_bar_line, self->up_bar_fill);

    /* Write the c:downBars element. */
    _chart_write_down_bars(self, self->down_bar_line, self->down_bar_fill);

    lxw_xml_end_tag(self->file, "c:upDownBars");
}

/*
 * Write the <c:dropLines> element.
 */
STATIC void
_chart_write_drop_lines(lxw_chart *self)
{
    if (!self->has_drop_lines)
        return;

    if (self->drop_lines_line) {
        lxw_xml_start_tag(self->file, "c:dropLines", NULL);

        _chart_write_sp_pr(self, self->drop_lines_line, NULL, NULL);

        lxw_xml_end_tag(self->file, "c:dropLines");
    }
    else {
        lxw_xml_empty_tag(self->file, "c:dropLines", NULL);
    }
}

/*
 * Write the <c:hiLowLines> element.
 */
STATIC void
_chart_write_hi_low_lines(lxw_chart *self)
{
    if (!self->has_high_low_lines)
        return;

    if (self->high_low_lines_line) {
        lxw_xml_start_tag(self->file, "c:hiLowLines", NULL);

        _chart_write_sp_pr(self, self->high_low_lines_line, NULL, NULL);

        lxw_xml_end_tag(self->file, "c:hiLowLines");
    }
    else {
        lxw_xml_empty_tag(self->file, "c:hiLowLines", NULL);
    }
}

/*
 * Write the <c:title> element.
 */
STATIC void
_chart_write_title(lxw_chart *self, lxw_chart_title *title)
{
    if (title->name) {
        /* Write the c:title element. */
        _chart_write_title_rich(self, title);
    }
    else if (title->range->formula) {
        /* Write the c:title element. */
        _chart_write_title_formula(self, title);
    }
}

/*
 * Write the <c:title> element.
 */
STATIC void
_chart_write_chart_title(lxw_chart *self)
{
    if (self->title.off) {
        /* Write the c:autoTitleDeleted element. */
        _chart_write_auto_title_deleted(self);
    }
    else {
        /* Write the c:title element. */
        _chart_write_title(self, &self->title);
    }
}

/*
 * Write the <c:catAx> element. Usually the X axis.
 */
STATIC void
_chart_write_cat_axis(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:catAx", NULL);

    _chart_write_axis_id(self, self->axis_id_1);

    /* Write the c:scaling element. Note we can't set max, min, or log base
     * for a Category axis in Excel.*/
    _chart_write_scaling(self,
                         self->x_axis->reverse,
                         LXW_FALSE, 0.0, LXW_FALSE, 0.0, 0);

    /* Write the c:delete element to hide axis. */
    if (self->x_axis->hidden)
        _chart_write_delete(self);

    /* Write the c:axPos element. */
    _chart_write_axis_pos(self, self->x_axis->axis_position,
                          self->y_axis->reverse);

    /* Write the c:majorGridlines element. */
    _chart_write_major_gridlines(self, self->x_axis);

    /* Write the c:minorGridlines element. */
    _chart_write_minor_gridlines(self, self->x_axis);

    /* Write the axis title elements. */
    self->x_axis->title.is_horizontal = self->has_horiz_cat_axis;
    _chart_write_title(self, &self->x_axis->title);

    /* Write the c:numFmt element. */
    _chart_write_cat_number_format(self, self->x_axis);

    /* Write the c:majorTickMark element. */
    _chart_write_major_tick_mark(self, self->x_axis);

    /* Write the c:minorTickMark element. */
    _chart_write_minor_tick_mark(self, self->x_axis);

    /* Write the c:tickLblPos element. */
    _chart_write_tick_label_pos(self, self->x_axis);

    /* Write the c:spPr element for the axis line. */
    _chart_write_sp_pr(self, self->x_axis->line, self->x_axis->fill,
                       self->x_axis->pattern);

    /* Write the axis font elements. */
    _chart_write_axis_font(self, self->x_axis->num_font);

    /* Write the c:crossAx element. */
    _chart_write_cross_axis(self, self->axis_id_2);

    /* Write the c:crosses element. */
    if (!self->y_axis->has_crossing || self->y_axis->crossing_max)
        _chart_write_crosses(self, self->y_axis);
    else
        _chart_write_crosses_at(self, self->y_axis);

    /* Write the c:auto element. */
    _chart_write_auto(self);

    /* Write the c:lblAlgn element. */
    _chart_write_label_align(self);

    /* Write the c:lblOffset element. */
    _chart_write_label_offset(self);

    /* Write the c:tickLblSkip element. */
    _chart_write_tick_label_skip(self, self->x_axis);

    /* Write the c:tickMarkSkip element. */
    _chart_write_tick_mark_skip(self, self->x_axis);

    lxw_xml_end_tag(self->file, "c:catAx");
}

/*
 * Write the <c:valAx> element.
 */
STATIC void
_chart_write_val_axis(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:valAx", NULL);

    _chart_write_axis_id(self, self->axis_id_2);

    /* Write the c:scaling element. */
    _chart_write_scaling(self,
                         self->y_axis->reverse,
                         self->y_axis->has_min, self->y_axis->min,
                         self->y_axis->has_max, self->y_axis->max,
                         self->y_axis->log_base);

    /* Write the c:delete element to hide axis. */
    if (self->y_axis->hidden)
        _chart_write_delete(self);

    /* Write the c:axPos element. */
    _chart_write_axis_pos(self, self->y_axis->axis_position,
                          self->x_axis->reverse);

    /* Write the c:majorGridlines element. */
    _chart_write_major_gridlines(self, self->y_axis);

    /* Write the c:minorGridlines element. */
    _chart_write_minor_gridlines(self, self->y_axis);

    /* Write the axis title elements. */
    self->y_axis->title.is_horizontal = self->has_horiz_val_axis;
    _chart_write_title(self, &self->y_axis->title);

    /* Write the c:numFmt element. */
    _chart_write_number_format(self, self->y_axis);

    /* Write the c:majorTickMark element. */
    _chart_write_major_tick_mark(self, self->y_axis);

    /* Write the c:minorTickMark element. */
    _chart_write_minor_tick_mark(self, self->y_axis);

    /* Write the c:tickLblPos element. */
    _chart_write_tick_label_pos(self, self->y_axis);

    /* Write the c:spPr element for the axis line. */
    _chart_write_sp_pr(self, self->y_axis->line, self->y_axis->fill,
                       self->y_axis->pattern);

    /* Write the axis font elements. */
    _chart_write_axis_font(self, self->y_axis->num_font);

    /* Write the c:crossAx element. */
    _chart_write_cross_axis(self, self->axis_id_1);

    /* Write the c:crosses element. */
    if (!self->x_axis->has_crossing || self->x_axis->crossing_max)
        _chart_write_crosses(self, self->x_axis);
    else
        _chart_write_crosses_at(self, self->x_axis);

    /* Write the c:crossBetween element. */
    _chart_write_cross_between(self, self->x_axis->position_axis);

    /* Write the c:majorUnit element. */
    _chart_write_major_unit(self, self->y_axis);

    /* Write the c:minorUnit element. */
    _chart_write_minor_unit(self, self->y_axis);

    /* Write the c:dispUnits element. */
    _chart_write_disp_units(self, self->y_axis);

    lxw_xml_end_tag(self->file, "c:valAx");
}

/*
 * Write the <c:valAx> element. This is for the second valAx in scatter plots.
 */
STATIC void
_chart_write_cat_val_axis(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:valAx", NULL);

    _chart_write_axis_id(self, self->axis_id_1);

    /* Write the c:scaling element. */
    _chart_write_scaling(self,
                         self->x_axis->reverse,
                         self->x_axis->has_min, self->x_axis->min,
                         self->x_axis->has_max, self->x_axis->max,
                         self->x_axis->log_base);

    /* Write the c:delete element to hide axis. */
    if (self->x_axis->hidden)
        _chart_write_delete(self);

    /* Write the c:axPos element. */
    _chart_write_axis_pos(self, self->x_axis->axis_position,
                          self->y_axis->reverse);

    /* Write the c:majorGridlines element. */
    _chart_write_major_gridlines(self, self->x_axis);

    /* Write the c:minorGridlines element. */
    _chart_write_minor_gridlines(self, self->x_axis);

    /* Write the axis title elements. */
    self->x_axis->title.is_horizontal = self->has_horiz_val_axis;
    _chart_write_title(self, &self->x_axis->title);

    /* Write the c:numFmt element. */
    _chart_write_number_format(self, self->x_axis);

    /* Write the c:majorTickMark element. */
    _chart_write_major_tick_mark(self, self->x_axis);

    /* Write the c:minorTickMark element. */
    _chart_write_minor_tick_mark(self, self->x_axis);

    /* Write the c:tickLblPos element. */
    _chart_write_tick_label_pos(self, self->x_axis);

    /* Write the c:spPr element for the axis line. */
    _chart_write_sp_pr(self, self->x_axis->line, self->x_axis->fill,
                       self->x_axis->pattern);

    /* Write the axis font elements. */
    _chart_write_axis_font(self, self->x_axis->num_font);

    /* Write the c:crossAx element. */
    _chart_write_cross_axis(self, self->axis_id_2);

    /* Write the c:crosses element. */
    if (!self->y_axis->has_crossing || self->y_axis->crossing_max)
        _chart_write_crosses(self, self->y_axis);
    else
        _chart_write_crosses_at(self, self->y_axis);

    /* Write the c:crossBetween element. */
    _chart_write_cross_between(self, self->y_axis->position_axis);

    /* Write the c:majorUnit element. */
    _chart_write_major_unit(self, self->x_axis);

    /* Write the c:minorUnit element. */
    _chart_write_minor_unit(self, self->x_axis);

    /* Write the c:dispUnits element. */
    _chart_write_disp_units(self, self->x_axis);

    lxw_xml_end_tag(self->file, "c:valAx");
}

/*
 * Write the <c:barDir> element.
 */
STATIC void
_chart_write_bar_dir(lxw_chart *self, char *type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", type);

    lxw_xml_empty_tag(self->file, "c:barDir", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write a area chart.
 */
STATIC void
_chart_write_area_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:areaChart", NULL);

    /* Write the c:grouping element. */
    _chart_write_grouping(self, self->grouping);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:dropLines element. */
    _chart_write_drop_lines(self);

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:areaChart");
}

/*
 * Write a bar chart.
 */
STATIC void
_chart_write_bar_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:barChart", NULL);

    /* Write the c:barDir element. */
    _chart_write_bar_dir(self, "bar");

    /* Write the c:grouping element. */
    _chart_write_grouping(self, self->grouping);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:gapWidth element. */
    _chart_write_gap_width(self, self->gap_y1);

    /* Write the c:overlap element. */
    _chart_write_overlap(self, self->overlap_y1);

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:barChart");
}

/*
 * Write a column chart.
 */
STATIC void
_chart_write_column_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:barChart", NULL);

    /* Write the c:barDir element. */
    _chart_write_bar_dir(self, "col");

    /* Write the c:grouping element. */
    _chart_write_grouping(self, self->grouping);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:gapWidth element. */
    _chart_write_gap_width(self, self->gap_y1);

    /* Write the c:overlap element. */
    _chart_write_overlap(self, self->overlap_y1);

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:barChart");
}

/*
 * Write a doughnut chart.
 */
STATIC void
_chart_write_doughnut_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:doughnutChart", NULL);

    /* Write the c:varyColors element. */
    _chart_write_vary_colors(self);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:firstSliceAng element. */
    _chart_write_first_slice_ang(self);

    /* Write the c:holeSize element. */
    _chart_write_hole_size(self);

    lxw_xml_end_tag(self->file, "c:doughnutChart");
}

/*
 * Write a line chart.
 */
STATIC void
_chart_write_line_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:lineChart", NULL);

    /* Write the c:grouping element. */
    _chart_write_grouping(self, self->grouping);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:dropLines element. */
    _chart_write_drop_lines(self);

    /* Write the c:hiLowLines element. */
    _chart_write_hi_low_lines(self);

    /* Write the c:upDownBars element. */
    _chart_write_up_down_bars(self);

    /* Write the c:marker element. */
    _chart_write_marker_value(self);

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:lineChart");
}

/*
 * Write a pie chart.
 */
STATIC void
_chart_write_pie_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:pieChart", NULL);

    /* Write the c:varyColors element. */
    _chart_write_vary_colors(self);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:firstSliceAng element. */
    _chart_write_first_slice_ang(self);

    lxw_xml_end_tag(self->file, "c:pieChart");
}

/*
 * Write a scatter chart.
 */
STATIC void
_chart_write_scatter_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:scatterChart", NULL);

    /* Write the c:scatterStyle element. */
    _chart_write_scatter_style(self);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {

        /* Add default scatter chart formatting to the series data unless
         * it has already been specified by the user.*/
        if (self->type == LXW_CHART_SCATTER) {
            if (!series->line) {
                lxw_chart_line line = {
                    0x000000,
                    LXW_TRUE,
                    2.25,
                    LXW_CHART_LINE_DASH_SOLID,
                    0,
                    LXW_FALSE
                };
                series->line = _chart_convert_line_args(&line);
            }
        }

        /* Write the c:ser element. */
        _chart_write_xval_ser(self, series);
    }

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:scatterChart");
}

/*
 * Write a radar chart.
 */
STATIC void
_chart_write_radar_chart(lxw_chart *self)
{
    lxw_chart_series *series;

    lxw_xml_start_tag(self->file, "c:radarChart", NULL);

    /* Write the c:radarStyle element. */
    _chart_write_radar_style(self);

    STAILQ_FOREACH(series, self->series_list, list_pointers) {
        /* Write the c:ser element. */
        _chart_write_ser(self, series);
    }

    /* Write the c:axId elements. */
    _chart_write_axis_ids(self);

    lxw_xml_end_tag(self->file, "c:radarChart");
}

/*
 * Reverse the opposite axis position if crossing position is "max".
 */
STATIC void
_chart_adjust_max_crossing(lxw_chart *self)
{
    if (self->x_axis->crossing_max)
        self->y_axis->axis_position ^= 1;

    if (self->y_axis->crossing_max)
        self->x_axis->axis_position ^= 1;
}

/*
 * Write the <c:plotArea> element.
 */
STATIC void
_chart_write_scatter_plot_area(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:plotArea", NULL);

    /* Write the c:layout element. */
    _chart_write_layout(self);

    /* Write subclass chart type elements for primary and secondary axes. */
    self->write_chart_type(self);

    /* Reverse the opposite axis position if crossing position is "max". */
    _chart_adjust_max_crossing(self);

    /* Write the c:catAx element. */
    _chart_write_cat_val_axis(self);

    self->has_horiz_val_axis = LXW_TRUE;

    /* Write the c:valAx element. */
    _chart_write_val_axis(self);

    /* Write the c:spPr element for the plotarea formatting. */
    _chart_write_sp_pr(self, self->plotarea_line, self->plotarea_fill,
                       self->plotarea_pattern);

    lxw_xml_end_tag(self->file, "c:plotArea");
}

/*
 * Write the <c:plotArea> element. Special handling for pie/doughnut.
 */
STATIC void
_chart_write_pie_plot_area(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:plotArea", NULL);

    /* Write the c:layout element. */
    _chart_write_layout(self);

    /* Write subclass chart type elements for primary and secondary axes. */
    self->write_chart_type(self);

    /* Write the c:spPr element for the plotarea formatting. */
    _chart_write_sp_pr(self, self->plotarea_line, self->plotarea_fill,
                       self->plotarea_pattern);

    lxw_xml_end_tag(self->file, "c:plotArea");
}

/*
 * Write the <c:plotArea> element.
 */
STATIC void
_chart_write_plot_area(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:plotArea", NULL);

    /* Write the c:layout element. */
    _chart_write_layout(self);

    /* Write subclass chart type elements for primary and secondary axes. */
    self->write_chart_type(self);

    /* Reverse the opposite axis position if crossing position is "max". */
    _chart_adjust_max_crossing(self);

    /* Write the c:catAx element. */
    _chart_write_cat_axis(self);

    /* Write the c:valAx element. */
    _chart_write_val_axis(self);

    /* Write the c:dTable element. */
    _chart_write_d_table(self);

    /* Write the c:spPr element for the plotarea formatting. */
    _chart_write_sp_pr(self, self->plotarea_line, self->plotarea_fill,
                       self->plotarea_pattern);

    lxw_xml_end_tag(self->file, "c:plotArea");
}

/*
 * Write the <c:chart> element.
 */
STATIC void
_chart_write_chart(lxw_chart *self)
{
    lxw_xml_start_tag(self->file, "c:chart", NULL);

    /* Write the c:title element. */
    _chart_write_chart_title(self);

    /* Write the c:plotArea element. */
    self->write_plot_area(self);

    /* Write the c:legend element. */
    _chart_write_legend(self);

    /* Write the c:plotVisOnly element. */
    _chart_write_plot_vis_only(self);

    /* Write the c:dispBlanksAs element. */
    _chart_write_disp_blanks_as(self);

    lxw_xml_end_tag(self->file, "c:chart");
}

/*
 * Initialize a area chart.
 */
STATIC void
_chart_initialize_area_chart(lxw_chart *self, uint8_t type)
{
    self->chart_group = LXW_CHART_AREA;
    self->grouping = LXW_GROUPING_STANDARD;
    self->default_cross_between = LXW_CHART_AXIS_POSITION_ON_TICK;
    self->x_axis->is_category = LXW_TRUE;
    self->default_label_position = LXW_CHART_LABEL_POSITION_CENTER;

    if (type == LXW_CHART_AREA_STACKED) {
        self->grouping = LXW_GROUPING_STACKED;
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    if (type == LXW_CHART_AREA_STACKED_PERCENT) {
        self->grouping = LXW_GROUPING_PERCENTSTACKED;
        _chart_axis_set_default_num_format(self->y_axis, "0%");
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_area_chart;
    self->write_plot_area = _chart_write_plot_area;
}

/*
 * Swap/reverse the bar chart axes prior to writing. It is the only chart
 * with the category axis in the vertical direction.
 */
STATIC void
_chart_swap_bar_axes(lxw_chart *self)
{
    lxw_chart_axis *tmp = self->x_axis;
    self->x_axis = self->y_axis;
    self->y_axis = tmp;
}

/*
 * Initialize a bar chart.
 */
STATIC void
_chart_initialize_bar_chart(lxw_chart *self, uint8_t type)
{
    /* Note: Bar chart category/value axis are reversed in comparison to
     *       other charts. Some of the defaults reflect this. */
    self->chart_group = LXW_CHART_BAR;
    self->x_axis->major_gridlines.visible = LXW_TRUE;
    self->y_axis->major_gridlines.visible = LXW_FALSE;
    self->y_axis->is_category = LXW_TRUE;
    self->x_axis->is_value = LXW_TRUE;
    self->has_horiz_cat_axis = LXW_TRUE;
    self->has_horiz_val_axis = LXW_FALSE;
    self->default_label_position = LXW_CHART_LABEL_POSITION_OUTSIDE_END;

    if (type == LXW_CHART_BAR_STACKED) {
        self->grouping = LXW_GROUPING_STACKED;
        self->has_overlap = LXW_TRUE;
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
        self->overlap_y1 = 100;

    }

    if (type == LXW_CHART_BAR_STACKED_PERCENT) {
        self->grouping = LXW_GROUPING_PERCENTSTACKED;
        _chart_axis_set_default_num_format(self->x_axis, "0%");
        self->has_overlap = LXW_TRUE;
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
        self->overlap_y1 = 100;
    }

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_bar_chart;
    self->write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a column chart.
 */
STATIC void
_chart_initialize_column_chart(lxw_chart *self, uint8_t type)
{
    self->chart_group = LXW_CHART_COLUMN;
    self->has_horiz_val_axis = LXW_FALSE;
    self->x_axis->is_category = LXW_TRUE;
    self->y_axis->is_value = LXW_TRUE;
    self->default_label_position = LXW_CHART_LABEL_POSITION_OUTSIDE_END;

    if (type == LXW_CHART_COLUMN_STACKED) {
        self->grouping = LXW_GROUPING_STACKED;
        self->has_overlap = LXW_TRUE;
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
        self->overlap_y1 = 100;
    }

    if (type == LXW_CHART_COLUMN_STACKED_PERCENT) {
        self->grouping = LXW_GROUPING_PERCENTSTACKED;
        _chart_axis_set_default_num_format(self->y_axis, "0%");
        self->has_overlap = LXW_TRUE;
        self->subtype = LXW_CHART_SUBTYPE_STACKED;
        self->overlap_y1 = 100;
    }

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_column_chart;
    self->write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a doughnut chart.
 */
STATIC void
_chart_initialize_doughnut_chart(lxw_chart *self)
{
    /* Initialize the function pointers for this chart type. */
    self->chart_group = LXW_CHART_DOUGHNUT;
    self->write_chart_type = _chart_write_doughnut_chart;
    self->write_plot_area = _chart_write_pie_plot_area;
    self->default_label_position = LXW_CHART_LABEL_POSITION_BEST_FIT;
}

/*
 * Initialize a line chart.
 */
STATIC void
_chart_initialize_line_chart(lxw_chart *self)
{
    self->chart_group = LXW_CHART_LINE;
    _chart_set_default_marker_type(self, LXW_CHART_MARKER_NONE);
    self->grouping = LXW_GROUPING_STANDARD;
    self->x_axis->is_category = LXW_TRUE;
    self->y_axis->is_value = LXW_TRUE;
    self->default_label_position = LXW_CHART_LABEL_POSITION_RIGHT;

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_line_chart;
    self->write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a pie chart.
 */
STATIC void
_chart_initialize_pie_chart(lxw_chart *self)
{
    /* Initialize the function pointers for this chart type. */
    self->chart_group = LXW_CHART_PIE;
    self->write_chart_type = _chart_write_pie_chart;
    self->write_plot_area = _chart_write_pie_plot_area;
    self->default_label_position = LXW_CHART_LABEL_POSITION_BEST_FIT;
}

/*
 * Initialize a scatter chart.
 */
STATIC void
_chart_initialize_scatter_chart(lxw_chart *self)
{
    self->chart_group = LXW_CHART_SCATTER;
    self->has_horiz_val_axis = LXW_FALSE;
    self->default_cross_between = LXW_CHART_AXIS_POSITION_ON_TICK;
    self->x_axis->is_value = LXW_TRUE;
    self->y_axis->is_value = LXW_TRUE;
    self->default_label_position = LXW_CHART_LABEL_POSITION_RIGHT;

    if (self->type == LXW_CHART_SCATTER_STRAIGHT
        || self->type == LXW_CHART_SCATTER_SMOOTH) {

        _chart_set_default_marker_type(self, LXW_CHART_MARKER_NONE);
    }

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_scatter_chart;
    self->write_plot_area = _chart_write_scatter_plot_area;
}

/*
 * Initialize a radar chart.
 */
STATIC void
_chart_initialize_radar_chart(lxw_chart *self, uint8_t type)
{
    if (type == LXW_CHART_RADAR)
        _chart_set_default_marker_type(self, LXW_CHART_MARKER_NONE);

    self->chart_group = LXW_CHART_RADAR;
    self->x_axis->major_gridlines.visible = LXW_TRUE;
    self->x_axis->is_category = LXW_TRUE;
    self->y_axis->is_value = LXW_TRUE;
    self->y_axis->major_tick_mark = LXW_CHART_AXIS_TICK_MARK_CROSSING;
    self->default_label_position = LXW_CHART_LABEL_POSITION_CENTER;

    /* Initialize the function pointers for this chart type. */
    self->write_chart_type = _chart_write_radar_chart;
    self->write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize the chart specific properties.
 */
STATIC void
_chart_initialize(lxw_chart *self, uint8_t type)
{
    switch (type) {

        case LXW_CHART_AREA:
        case LXW_CHART_AREA_STACKED:
        case LXW_CHART_AREA_STACKED_PERCENT:
            _chart_initialize_area_chart(self, type);
            break;

        case LXW_CHART_BAR:
        case LXW_CHART_BAR_STACKED:
        case LXW_CHART_BAR_STACKED_PERCENT:
            _chart_initialize_bar_chart(self, type);
            break;

        case LXW_CHART_COLUMN:
        case LXW_CHART_COLUMN_STACKED:
        case LXW_CHART_COLUMN_STACKED_PERCENT:
            _chart_initialize_column_chart(self, type);
            break;

        case LXW_CHART_DOUGHNUT:
            _chart_initialize_doughnut_chart(self);
            break;

        case LXW_CHART_LINE:
            _chart_initialize_line_chart(self);
            break;

        case LXW_CHART_PIE:
            _chart_initialize_pie_chart(self);
            break;

        case LXW_CHART_SCATTER:
        case LXW_CHART_SCATTER_STRAIGHT:
        case LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS:
        case LXW_CHART_SCATTER_SMOOTH:
        case LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS:
            _chart_initialize_scatter_chart(self);
            break;

        case LXW_CHART_RADAR:
        case LXW_CHART_RADAR_WITH_MARKERS:
        case LXW_CHART_RADAR_FILLED:
            _chart_initialize_radar_chart(self, type);
            break;

        default:
            LXW_WARN_FORMAT1("workbook_add_chart(): "
                             "unhandled chart type '%d'", type);
    }
}

/*
 * Assemble and write the XML file.
 */
void
lxw_chart_assemble_xml_file(lxw_chart *self)
{
    /* Reverse the X and Y axes for Bar charts. */
    if (self->type == LXW_CHART_BAR || self->type == LXW_CHART_BAR_STACKED
        || self->type == LXW_CHART_BAR_STACKED_PERCENT)
        _chart_swap_bar_axes(self);

    /* Write the XML declaration. */
    _chart_xml_declaration(self);

    /* Write the c:chartSpace element. */
    _chart_write_chart_space(self);

    /* Write the c:lang element. */
    _chart_write_lang(self);

    /* Write the c:style element. */
    _chart_write_style(self);

    /* Write the c:chart element. */
    _chart_write_chart(self);

    /* Write the c:spPr element for the chartarea formatting. */
    _chart_write_sp_pr(self, self->chartarea_line, self->chartarea_fill,
                       self->chartarea_pattern);

    /* Write the c:printSettings element. */
    _chart_write_print_settings(self);

    lxw_xml_end_tag(self->file, "c:chartSpace");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add data to a data cache in a range object, for testing only.
 */
lxw_error
lxw_chart_add_data_cache(lxw_series_range *range, uint8_t *data,
                         uint16_t rows, uint8_t cols, uint8_t col)
{
    struct lxw_series_data_point *data_point;
    uint16_t i;

    range->ignore_cache = LXW_TRUE;
    range->num_data_points = rows;

    /* Initialize the series range data cache. */
    for (i = 0; i < rows; i++) {
        data_point = calloc(1, sizeof(struct lxw_series_data_point));
        RETURN_ON_MEM_ERROR(data_point, LXW_ERROR_MEMORY_MALLOC_FAILED);
        STAILQ_INSERT_TAIL(range->data_cache, data_point, list_pointers);
        data_point->number = data[i * cols + col];
    }

    return LXW_NO_ERROR;
}

/*
 * Insert an image into the worksheet.
 */
lxw_chart_series *
chart_add_series(lxw_chart *self, const char *categories, const char *values)
{
    lxw_chart_series *series;

    /* Scatter charts require categories and values. */
    if (self->chart_group == LXW_CHART_SCATTER && values && !categories) {
        LXW_WARN("chart_add_series(): scatter charts must have "
                 "'categories' and 'values'");

        return NULL;
    }

    /* Create a new object to hold the series. */
    series = calloc(1, sizeof(lxw_chart_series));
    GOTO_LABEL_ON_MEM_ERROR(series, mem_error);

    series->categories = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->categories, mem_error);

    series->values = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->values, mem_error);

    series->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->title.range, mem_error);

    series->x_error_bars = calloc(1, sizeof(lxw_series_error_bars));
    GOTO_LABEL_ON_MEM_ERROR(series->x_error_bars, mem_error);

    series->y_error_bars = calloc(1, sizeof(lxw_series_error_bars));
    GOTO_LABEL_ON_MEM_ERROR(series->y_error_bars, mem_error);

    if (categories) {
        if (categories[0] == '=')
            series->categories->formula = lxw_strdup(categories + 1);
        else
            series->categories->formula = lxw_strdup(categories);
    }

    if (values) {
        if (values[0] == '=')
            series->values->formula = lxw_strdup(values + 1);
        else
            series->values->formula = lxw_strdup(values);
    }

    if (_chart_init_data_cache(series->categories) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(series->values) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(series->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (self->type == LXW_CHART_SCATTER_SMOOTH)
        series->smooth = LXW_TRUE;

    if (self->type == LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS)
        series->smooth = LXW_TRUE;

    series->y_error_bars->chart_group = self->chart_group;
    series->x_error_bars->chart_group = self->chart_group;
    series->x_error_bars->is_x = LXW_TRUE;

    series->default_label_position = self->default_label_position;

    STAILQ_INSERT_TAIL(self->series_list, series, list_pointers);

    return series;

mem_error:
    _chart_series_free(series);
    return NULL;
}

/*
 * Set on of the 48 built-in Excel chart styles.
 */
void
chart_set_style(lxw_chart *self, uint8_t style_id)
{
    /* The default style is 2. The range is 1 - 48 */
    if (style_id < 1 || style_id > 48)
        style_id = 2;

    self->style_id = style_id;
}

/*
 * Set a user defined name for a series.
 */
void
chart_series_set_name(lxw_chart_series *series, const char *name)
{
    if (!name)
        return;

    if (name[0] == '=')
        series->title.range->formula = lxw_strdup(name + 1);
    else
        series->title.name = lxw_strdup(name);
}

/*
 * Set an axis caption, with a range instead or a formula..
 */
void
chart_series_set_name_range(lxw_chart_series *series, const char *sheetname,
                            lxw_row_t row, lxw_col_t col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_name_range(): "
                 "sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    _chart_set_range(series->title.range, sheetname, row, col, row, col);
}

/*
 * Set the categories range for a series.
 */
void
chart_series_set_categories(lxw_chart_series *series, const char *sheetname,
                            lxw_row_t first_row, lxw_col_t first_col,
                            lxw_row_t last_row, lxw_col_t last_col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_categories(): "
                 "sheetname must be specified");
        return;
    }

    _chart_set_range(series->categories, sheetname,
                     first_row, first_col, last_row, last_col);
}

/*
 * Set the values range for a series.
 */
void
chart_series_set_values(lxw_chart_series *series, const char *sheetname,
                        lxw_row_t first_row, lxw_col_t first_col,
                        lxw_row_t last_row, lxw_col_t last_col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_values(): sheetname must be specified");
        return;
    }

    _chart_set_range(series->values, sheetname,
                     first_row, first_col, last_row, last_col);
}

/*
 * Set a line type for a series.
 */
void
chart_series_set_line(lxw_chart_series *series, lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(series->line);

    series->line = _chart_convert_line_args(line);
}

/*
 * Set a fill type for a series.
 */
void
chart_series_set_fill(lxw_chart_series *series, lxw_chart_fill *fill)
{
    if (!fill)
        return;

    /* Free any previously allocated resource. */
    free(series->fill);

    series->fill = _chart_convert_fill_args(fill);
}

/*
 * Invert the colors of a fill for a series.
 */
void
chart_series_set_invert_if_negative(lxw_chart_series *series)
{
    series->invert_if_negative = LXW_TRUE;
}

/*
 * Set a pattern type for a series.
 */
void
chart_series_set_pattern(lxw_chart_series *series, lxw_chart_pattern *pattern)
{
    if (!pattern)
        return;

    /* Free any previously allocated resource. */
    free(series->pattern);

    series->pattern = _chart_convert_pattern_args(pattern);
}

/*
 * Set a marker type for a series.
 */
void
chart_series_set_marker_type(lxw_chart_series *series, uint8_t type)
{
    if (!series->marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        series->marker = marker;
    }

    series->marker->type = type;
}

/*
 * Set a marker size for a series.
 */
void
chart_series_set_marker_size(lxw_chart_series *series, uint8_t size)
{
    if (!series->marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        series->marker = marker;
    }

    series->marker->size = size;
}

/*
 * Set a line type for a series marker.
 */
void
chart_series_set_marker_line(lxw_chart_series *series, lxw_chart_line *line)
{
    if (!line)
        return;

    if (!series->marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        series->marker = marker;
    }

    /* Free any previously allocated resource. */
    free(series->marker->line);

    series->marker->line = _chart_convert_line_args(line);
}

/*
 * Set a fill type for a series marker.
 */
void
chart_series_set_marker_fill(lxw_chart_series *series, lxw_chart_fill *fill)
{
    if (!fill)
        return;

    if (!series->marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        series->marker = marker;
    }

    /* Free any previously allocated resource. */
    free(series->marker->fill);

    series->marker->fill = _chart_convert_fill_args(fill);
}

/*
 * Set a pattern type for a series.
 */
void
chart_series_set_marker_pattern(lxw_chart_series *series,
                                lxw_chart_pattern *pattern)
{
    if (!pattern)
        return;

    if (!series->marker) {
        lxw_chart_marker *marker = calloc(1, sizeof(struct lxw_chart_marker));
        RETURN_VOID_ON_MEM_ERROR(marker);
        series->marker = marker;
    }

    /* Free any previously allocated resource. */
    free(series->marker->pattern);

    series->marker->pattern = _chart_convert_pattern_args(pattern);
}

/*
 * Store the horizontal page breaks on a worksheet.
 */
lxw_error
chart_series_set_points(lxw_chart_series *series, lxw_chart_point *points[])
{
    uint16_t i = 0;
    uint16_t point_count = 0;

    if (points == NULL)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    while (points[point_count])
        point_count++;

    if (point_count == 0)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* Free any existing resource. */
    _chart_free_points(series);

    series->points = calloc(point_count, sizeof(lxw_chart_point));
    RETURN_ON_MEM_ERROR(series->points, LXW_ERROR_MEMORY_MALLOC_FAILED);

    for (i = 0; i < point_count; i++) {
        lxw_chart_point *src_point = points[i];
        lxw_chart_point *dst_point = &series->points[i];

        dst_point->line = _chart_convert_line_args(src_point->line);
        dst_point->fill = _chart_convert_fill_args(src_point->fill);
        dst_point->pattern = _chart_convert_pattern_args(src_point->pattern);
    }

    series->point_count = point_count;

    return LXW_NO_ERROR;
}

/*
 * Set the smooth property for a line or scatter series.
 */
void
chart_series_set_smooth(lxw_chart_series *series, uint8_t smooth)
{
    series->smooth = smooth;
}

/*
 * Turn on default data labels for a series.
 */
void
chart_series_set_labels(lxw_chart_series *series)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_value = LXW_TRUE;
}

/*
 * Set the data labels options for a series.
 */
void
chart_series_set_labels_options(lxw_chart_series *series, uint8_t show_name,
                                uint8_t show_category, uint8_t show_value)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_name = show_name;
    series->show_labels_category = show_category;
    series->show_labels_value = show_value;
}

/*
 * Set the data labels separator for a series.
 */
void
chart_series_set_labels_separator(lxw_chart_series *series, uint8_t separator)
{
    series->has_labels = LXW_TRUE;
    series->label_separator = separator;
}

/*
 * Set the data labels position for a series.
 */
void
chart_series_set_labels_position(lxw_chart_series *series, uint8_t position)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_value = LXW_TRUE;

    if (position != series->default_label_position)
        series->label_position = position;
}

/*
 * Set the data labels position for a series.
 */
void
chart_series_set_labels_leader_line(lxw_chart_series *series)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_leader = LXW_TRUE;
}

/*
 * Turn on the data labels legend for a series.
 */
void
chart_series_set_labels_legend(lxw_chart_series *series)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_legend = LXW_TRUE;
}

/*
 * Turn on the data labels percentage for a series.
 */
void
chart_series_set_labels_percentage(lxw_chart_series *series)
{
    series->has_labels = LXW_TRUE;
    series->show_labels_percent = LXW_TRUE;
}

/*
 * Set an data labels number format.
 */
void
chart_series_set_labels_num_format(lxw_chart_series *series,
                                   const char *num_format)
{
    if (!num_format)
        return;

    /* Free any previously allocated resource. */
    free(series->label_num_format);

    series->label_num_format = lxw_strdup(num_format);
}

/*
 * Set an data labels font.
 */
void
chart_series_set_labels_font(lxw_chart_series *series, lxw_chart_font *font)
{
    if (!font)
        return;

    /* Free any previously allocated resource. */
    _chart_free_font(series->label_font);

    series->label_font = _chart_convert_font_args(font);
}

/*
 * Set the trendline for a chart series.
 */
void
chart_series_set_trendline(lxw_chart_series *series, uint8_t type,
                           uint8_t value)
{
    if (type == LXW_CHART_TRENDLINE_TYPE_POLY
        || type == LXW_CHART_TRENDLINE_TYPE_AVERAGE) {

        if (value < 2) {
            LXW_WARN("chart_series_set_trendline(): order/period value must "
                     "be >= 2 for Polynomial and Moving Average types");
            return;
        }

        series->trendline_value_type = type;
    }

    series->has_trendline = LXW_TRUE;
    series->trendline_type = type;
    series->trendline_value = value;
}

/*
 * Set the trendline forecast for a chart series.
 */
void
chart_series_set_trendline_forecast(lxw_chart_series *series, double forward,
                                    double backward)
{
    if (!series->has_trendline) {
        LXW_WARN("chart_series_set_trendline_forecast(): trendline type "
                 "must be set first using chart_series_set_trendline()");
        return;
    }

    if (series->trendline_type == LXW_CHART_TRENDLINE_TYPE_AVERAGE) {
        LXW_WARN("chart_series_set_trendline(): forecast isn't available "
                 "in Excel for a Moving Average trendline");
        return;
    }

    series->has_trendline_forecast = LXW_TRUE;
    series->trendline_forward = forward;
    series->trendline_backward = backward;
}

/*
 * Display the equation for a series trendline.
 */
void
chart_series_set_trendline_equation(lxw_chart_series *series)
{
    if (!series->has_trendline) {
        LXW_WARN("chart_series_set_trendline_equation(): trendline type "
                 "must be set first using chart_series_set_trendline()");
        return;
    }

    if (series->trendline_type == LXW_CHART_TRENDLINE_TYPE_AVERAGE) {
        LXW_WARN("chart_series_set_trendline_equation(): equation isn't "
                 "available in Excel for a Moving Average trendline");
        return;
    }

    series->has_trendline_equation = LXW_TRUE;
}

/*
 * Display the R squared value for a series trendline.
 */
void
chart_series_set_trendline_r_squared(lxw_chart_series *series)
{
    if (!series->has_trendline) {
        LXW_WARN("chart_series_set_trendline_r_squared(): trendline type "
                 "must be set first using chart_series_set_trendline()");
        return;
    }

    if (series->trendline_type == LXW_CHART_TRENDLINE_TYPE_AVERAGE) {
        LXW_WARN("chart_series_set_trendline_r_squared(): R squared isn't "
                 "available in Excel for a Moving Average trendline");
        return;
    }

    series->has_trendline_r_squared = LXW_TRUE;
}

/*
 * Set the trendline intercept for a chart series.
 */
void
chart_series_set_trendline_intercept(lxw_chart_series *series,
                                     double intercept)
{
    if (!series->has_trendline) {
        LXW_WARN("chart_series_set_trendline_intercept(): trendline type "
                 "must be set first using chart_series_set_trendline()");
        return;
    }

    if (series->trendline_type != LXW_CHART_TRENDLINE_TYPE_EXP
        && series->trendline_type != LXW_CHART_TRENDLINE_TYPE_LINEAR
        && series->trendline_type != LXW_CHART_TRENDLINE_TYPE_POLY) {

        LXW_WARN("chart_series_set_trendline_intercept(): intercept is only "
                 "available in Excel for Exponential, Linear and Polynomial "
                 "trendline types");
        return;
    }

    series->has_trendline_intercept = LXW_TRUE;
    series->trendline_intercept = intercept;
}

/*
 * Set a line type for a series trendline.
 */
void
chart_series_set_trendline_name(lxw_chart_series *series, const char *name)
{
    if (!name)
        return;

    /* Free any previously allocated resource. */
    free(series->trendline_name);

    series->trendline_name = lxw_strdup(name);
}

/*
 * Set a line type for a series trendline.
 */
void
chart_series_set_trendline_line(lxw_chart_series *series,
                                lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(series->trendline_line);

    series->trendline_line = _chart_convert_line_args(line);
}

/*
 * Set the X or Y error bars from a chart series.
 */
lxw_series_error_bars *
chart_series_get_error_bars(lxw_chart_series *series,
                            lxw_chart_error_bar_axis axis_type)
{
    if (!series)
        return NULL;

    if (axis_type == LXW_CHART_ERROR_BAR_AXIS_X)
        return series->x_error_bars;
    else if (axis_type == LXW_CHART_ERROR_BAR_AXIS_Y)
        return series->y_error_bars;
    else
        return NULL;
}

/*
 * Set the error bars and type for a chart series.
 */
void
chart_series_set_error_bars(lxw_series_error_bars *error_bars,
                            uint8_t type, double value)
{
    if (_chart_check_error_bars(error_bars, ""))
        return;

    error_bars->type = type;
    error_bars->value = value;
    error_bars->has_value = LXW_TRUE;
    error_bars->is_set = LXW_TRUE;

    if (type == LXW_CHART_ERROR_BAR_TYPE_STD_ERROR)
        error_bars->has_value = LXW_FALSE;
}

/*
 * Set the error bars direction for a chart series.
 */
void
chart_series_set_error_bars_direction(lxw_series_error_bars *error_bars,
                                      uint8_t direction)
{
    if (_chart_check_error_bars(error_bars, "_direction"))
        return;

    error_bars->direction = direction;
}

/*
 * Set the error bars end cap type for a chart series.
 */
void
chart_series_set_error_bars_endcap(lxw_series_error_bars *error_bars,
                                   uint8_t endcap)
{
    if (_chart_check_error_bars(error_bars, "_endcap"))
        return;

    error_bars->endcap = endcap;
}

/*
 * Set a line type for a series error bars.
 */
void
chart_series_set_error_bars_line(lxw_series_error_bars *error_bars,
                                 lxw_chart_line *line)
{
    if (_chart_check_error_bars(error_bars, "_line"))
        return;

    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(error_bars->line);

    error_bars->line = _chart_convert_line_args(line);
}

/*
 * Get an axis pointer from a chart.
 */
lxw_chart_axis *
chart_axis_get(lxw_chart *self, lxw_chart_axis_type axis_type)
{
    if (!self)
        return NULL;

    if (axis_type == LXW_CHART_AXIS_TYPE_X)
        return self->x_axis;
    else if (axis_type == LXW_CHART_AXIS_TYPE_Y)
        return self->y_axis;
    else
        return NULL;
}

/*
 * Set an axis caption.
 */
void
chart_axis_set_name(lxw_chart_axis *axis, const char *name)
{
    if (!name)
        return;

    if (name[0] == '=')
        axis->title.range->formula = lxw_strdup(name + 1);
    else
        axis->title.name = lxw_strdup(name);
}

/*
 * Set an axis caption, with a range instead or a formula.
 */
void
chart_axis_set_name_range(lxw_chart_axis *axis, const char *sheetname,
                          lxw_row_t row, lxw_col_t col)
{
    if (!sheetname) {
        LXW_WARN("chart_axis_set_name_range(): sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    _chart_set_range(axis->title.range, sheetname, row, col, row, col);
}

/*
 * Set an axis title/name font.
 */
void
chart_axis_set_name_font(lxw_chart_axis *axis, lxw_chart_font *font)
{
    if (!font)
        return;

    /* Free any previously allocated resource. */
    _chart_free_font(axis->title.font);

    axis->title.font = _chart_convert_font_args(font);
}

/*
 * Set an axis number font.
 */
void
chart_axis_set_num_font(lxw_chart_axis *axis, lxw_chart_font *font)
{
    if (!font)
        return;

    /* Free any previously allocated resource. */
    _chart_free_font(axis->num_font);

    axis->num_font = _chart_convert_font_args(font);
}

/*
 * Set an axis number format.
 */
void
chart_axis_set_num_format(lxw_chart_axis *axis, const char *num_format)
{
    if (!num_format)
        return;

    /* Free any previously allocated resource. */
    free(axis->num_format);

    axis->num_format = lxw_strdup(num_format);
}

/*
 * Set a line type for an axis.
 */
void
chart_axis_set_line(lxw_chart_axis *axis, lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(axis->line);

    axis->line = _chart_convert_line_args(line);
}

/*
 * Set a fill type for an axis.
 */
void
chart_axis_set_fill(lxw_chart_axis *axis, lxw_chart_fill *fill)
{
    if (!fill)
        return;

    /* Free any previously allocated resource. */
    free(axis->fill);

    axis->fill = _chart_convert_fill_args(fill);
}

/*
 * Set a pattern type for an axis.
 */
void
chart_axis_set_pattern(lxw_chart_axis *axis, lxw_chart_pattern *pattern)
{
    if (!pattern)
        return;

    /* Free any previously allocated resource. */
    free(axis->pattern);

    axis->pattern = _chart_convert_pattern_args(pattern);
}

/*
 * Reverse the direction of an axis.
 */
void
chart_axis_set_reverse(lxw_chart_axis *axis)
{
    axis->reverse = LXW_TRUE;
}

/*
 * Set the axis crossing position.
 */
void
chart_axis_set_crossing(lxw_chart_axis *axis, double value)
{
    axis->has_crossing = LXW_TRUE;
    axis->crossing = value;
}

/*
 * Set the axis crossing position as the max possible value.
 */
void
chart_axis_set_crossing_max(lxw_chart_axis *axis)
{
    axis->has_crossing = LXW_TRUE;
    axis->crossing_max = LXW_TRUE;
}

/*
 * Turn off/hide the axis.
 */
void
chart_axis_off(lxw_chart_axis *axis)
{
    axis->hidden = LXW_TRUE;
}

/*
 * Set the category axis position.
 */
void
chart_axis_set_position(lxw_chart_axis *axis, uint8_t position)
{
    LXW_WARN_CAT_AND_DATE_AXIS_ONLY("chart_axis_set_position");

    axis->position_axis = position;
}

/*
 * Set the axis label position.
 */
void
chart_axis_set_label_position(lxw_chart_axis *axis, uint8_t position)
{
    axis->label_position = position;
}

/*
 * Set the minimum value for an axis.
 */
void
chart_axis_set_min(lxw_chart_axis *axis, double min)
{
    LXW_WARN_VALUE_AND_DATE_AXIS_ONLY("chart_axis_set_min");

    axis->min = min;
    axis->has_min = LXW_TRUE;
}

/*
 * Set the maximum value for an axis.
 */
void
chart_axis_set_max(lxw_chart_axis *axis, double max)
{
    LXW_WARN_VALUE_AND_DATE_AXIS_ONLY("chart_axis_set_max");

    axis->max = max;
    axis->has_max = LXW_TRUE;
}

/*
 * Set the log base for an axis.
 */
void
chart_axis_set_log_base(lxw_chart_axis *axis, uint16_t log_base)
{
    LXW_WARN_VALUE_AXIS_ONLY("chart_axis_set_log_base");

    /* Excel log range is 2-1000. */
    if (log_base >= 2 && log_base <= 1000)
        axis->log_base = log_base;
}

/*
 * Set the major mark for an axis.
 */
void
chart_axis_set_major_tick_mark(lxw_chart_axis *axis, uint8_t type)
{
    axis->major_tick_mark = type;
}

/*
 * Set the minor mark for an axis.
 */
void
chart_axis_set_minor_tick_mark(lxw_chart_axis *axis, uint8_t type)
{
    axis->minor_tick_mark = type;
}

/*
 * Set interval unit for a category axis.
 */
void
chart_axis_set_interval_unit(lxw_chart_axis *axis, uint16_t unit)
{
    LXW_WARN_CAT_AND_DATE_AXIS_ONLY("chart_axis_set_major_unit");

    axis->interval_unit = unit;
}

/*
 * Set tick interval for a category axis.
 */
void
chart_axis_set_interval_tick(lxw_chart_axis *axis, uint16_t unit)
{
    LXW_WARN_CAT_AND_DATE_AXIS_ONLY("chart_axis_set_major_tick");

    axis->interval_tick = unit;
}

/*
 * Set major unit for a value axis.
 */
void
chart_axis_set_major_unit(lxw_chart_axis *axis, double unit)
{
    LXW_WARN_VALUE_AND_DATE_AXIS_ONLY("chart_axis_set_major_unit");

    axis->has_major_unit = LXW_TRUE;
    axis->major_unit = unit;
}

/*
 * Set minor unit for a value axis.
 */
void
chart_axis_set_minor_unit(lxw_chart_axis *axis, double unit)
{
    LXW_WARN_VALUE_AND_DATE_AXIS_ONLY("chart_axis_set_minor_unit");

    axis->has_minor_unit = LXW_TRUE;
    axis->minor_unit = unit;
}

/*
 * Set the display units for a value axis.
 */
void
chart_axis_set_display_units(lxw_chart_axis *axis, uint8_t units)
{
    LXW_WARN_VALUE_AXIS_ONLY("chart_axis_set_display_units");

    axis->display_units = units;
    axis->display_units_visible = LXW_TRUE;
}

/*
 * Turn on/off the display units for a value axis.
 */
void
chart_axis_set_display_units_visible(lxw_chart_axis *axis, uint8_t visible)
{
    LXW_WARN_VALUE_AXIS_ONLY("chart_axis_set_display_units");

    axis->display_units_visible = visible;
}

/*
 * Set the axis major gridlines on/off.
 */
void
chart_axis_major_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible)
{
    axis->major_gridlines.visible = visible;
}

/*
 * Set a line type for the major gridlines.
 */
void
chart_axis_major_gridlines_set_line(lxw_chart_axis *axis,
                                    lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(axis->major_gridlines.line);

    axis->major_gridlines.line = _chart_convert_line_args(line);

    /* If the gridline has a format it should also be visible. */
    if (axis->major_gridlines.line)
        axis->major_gridlines.visible = LXW_TRUE;
}

/*
 * Set the axis minor gridlines on/off.
 */
void
chart_axis_minor_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible)
{
    axis->minor_gridlines.visible = visible;
}

/*
 * Set a line type for the minor gridlines.
 */
void
chart_axis_minor_gridlines_set_line(lxw_chart_axis *axis,
                                    lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(axis->minor_gridlines.line);

    axis->minor_gridlines.line = _chart_convert_line_args(line);

    /* If the gridline has a format it should also be visible. */
    if (axis->minor_gridlines.line)
        axis->minor_gridlines.visible = LXW_TRUE;
}

/*
 * Set the chart title.
 */
void
chart_title_set_name(lxw_chart *self, const char *name)
{
    if (!name)
        return;

    if (name[0] == '=')
        self->title.range->formula = lxw_strdup(name + 1);
    else
        self->title.name = lxw_strdup(name);
}

/*
 * Set the chart title, with a range instead or a formula.
 */
void
chart_title_set_name_range(lxw_chart *self, const char *sheetname,
                           lxw_row_t row, lxw_col_t col)
{
    if (!sheetname) {
        LXW_WARN("chart_title_set_name_range(): sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    _chart_set_range(self->title.range, sheetname, row, col, row, col);
}

/*
 * Set the chart title font.
 */
void
chart_title_set_name_font(lxw_chart *self, lxw_chart_font *font)
{
    /* Free any previously allocated resource. */
    _chart_free_font(self->title.font);

    self->title.font = _chart_convert_font_args(font);
}

/*
 * Turn off the chart title.
 */
void
chart_title_off(lxw_chart *self)
{
    self->title.off = LXW_TRUE;
}

/*
 * Set the chart legend position.
 */
void
chart_legend_set_position(lxw_chart *self, uint8_t position)
{
    self->legend.position = position;
}

/*
 * Set the legend font.
 */
void
chart_legend_set_font(lxw_chart *self, lxw_chart_font *font)
{
    /* Free any previously allocated resource. */
    _chart_free_font(self->legend.font);

    self->legend.font = _chart_convert_font_args(font);
}

/*
 * Remove one or more series from the the legend.
 */
lxw_error
chart_legend_delete_series(lxw_chart *self, int16_t delete_series[])
{
    uint16_t count = 0;

    if (delete_series == NULL)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    while (delete_series[count] >= 0)
        count++;

    if (count == 0)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* The maximum number of series in a chart is 255. */
    if (count > 255)
        count = 255;

    self->delete_series = calloc(count, sizeof(int16_t));
    RETURN_ON_MEM_ERROR(self->delete_series, LXW_ERROR_MEMORY_MALLOC_FAILED);
    memcpy(self->delete_series, delete_series, count * sizeof(int16_t));
    self->delete_series_count = count;

    return LXW_NO_ERROR;
}

/*
 * Set a line type for the chartarea.
 */
void
chart_chartarea_set_line(lxw_chart *self, lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(self->chartarea_line);

    self->chartarea_line = _chart_convert_line_args(line);
}

/*
 * Set a fill type for the chartarea.
 */
void
chart_chartarea_set_fill(lxw_chart *self, lxw_chart_fill *fill)
{
    if (!fill)
        return;

    /* Free any previously allocated resource. */
    free(self->chartarea_fill);

    self->chartarea_fill = _chart_convert_fill_args(fill);
}

/*
 * Set a pattern type for the chartarea.
 */
void
chart_chartarea_set_pattern(lxw_chart *self, lxw_chart_pattern *pattern)
{
    if (!pattern)
        return;

    /* Free any previously allocated resource. */
    free(self->chartarea_pattern);

    self->chartarea_pattern = _chart_convert_pattern_args(pattern);
}

/*
 * Set a line type for the plotarea.
 */
void
chart_plotarea_set_line(lxw_chart *self, lxw_chart_line *line)
{
    if (!line)
        return;

    /* Free any previously allocated resource. */
    free(self->plotarea_line);

    self->plotarea_line = _chart_convert_line_args(line);
}

/*
 * Set a fill type for the plotarea.
 */
void
chart_plotarea_set_fill(lxw_chart *self, lxw_chart_fill *fill)
{
    if (!fill)
        return;

    /* Free any previously allocated resource. */
    free(self->plotarea_fill);

    self->plotarea_fill = _chart_convert_fill_args(fill);
}

/*
 * Set a pattern type for the plotarea.
 */
void
chart_plotarea_set_pattern(lxw_chart *self, lxw_chart_pattern *pattern)
{
    if (!pattern)
        return;

    /* Free any previously allocated resource. */
    free(self->plotarea_pattern);

    self->plotarea_pattern = _chart_convert_pattern_args(pattern);
}

/*
 * Turn on the chart data table.
 */
void
chart_set_table(lxw_chart *self)
{
    self->has_table = LXW_TRUE;
    self->has_table_horizontal = LXW_TRUE;
    self->has_table_vertical = LXW_TRUE;
    self->has_table_outline = LXW_TRUE;
    self->has_table_legend_keys = LXW_FALSE;
}

/*
 * Set the options for the chart data table grid.
 */
void
chart_set_table_grid(lxw_chart *self, uint8_t horizontal, uint8_t vertical,
                     uint8_t outline, uint8_t legend_keys)
{
    self->has_table = LXW_TRUE;
    self->has_table_horizontal = horizontal;
    self->has_table_vertical = vertical;
    self->has_table_outline = outline;
    self->has_table_legend_keys = legend_keys;
}

/*
 * Set the font for the chart data table grid.
 */
void
chart_set_table_font(lxw_chart *self, lxw_chart_font *font)
{
    self->has_table = LXW_TRUE;

    /* Free any previously allocated resource. */
    _chart_free_font(self->table_font);

    self->table_font = _chart_convert_font_args(font);
}

/*
 * Turn on up-down bars for the chart.
 */
void
chart_set_up_down_bars(lxw_chart *self)
{
    self->has_up_down_bars = LXW_TRUE;
}

/*
 * Turn on up-down bars for the chart, with formatting.
 */
void
chart_set_up_down_bars_format(lxw_chart *self, lxw_chart_line *up_bar_line,
                              lxw_chart_fill *up_bar_fill,
                              lxw_chart_line *down_bar_line,
                              lxw_chart_fill *down_bar_fill)
{
    self->has_up_down_bars = LXW_TRUE;

    /* Free any previously allocated resource. */
    free(self->up_bar_line);
    free(self->up_bar_fill);
    free(self->down_bar_line);
    free(self->down_bar_fill);

    self->up_bar_line = _chart_convert_line_args(up_bar_line);
    self->up_bar_fill = _chart_convert_fill_args(up_bar_fill);
    self->down_bar_line = _chart_convert_line_args(down_bar_line);
    self->down_bar_fill = _chart_convert_fill_args(down_bar_fill);
}

/*
 * Turn on drop lines for the chart.
 */
void
chart_set_drop_lines(lxw_chart *self, lxw_chart_line *line)
{
    /* Free any previously allocated resource. */
    free(self->drop_lines_line);

    self->has_drop_lines = LXW_TRUE;
    self->drop_lines_line = _chart_convert_line_args(line);
}

/*
 * Turn on high_low lines for the chart.
 */
void
chart_set_high_low_lines(lxw_chart *self, lxw_chart_line *line)
{
    /* Free any previously allocated resource. */
    free(self->high_low_lines_line);

    self->has_high_low_lines = LXW_TRUE;
    self->high_low_lines_line = _chart_convert_line_args(line);
}

/*
 * Set the Bar/Column overlap for all data series.
 */
void
chart_set_series_overlap(lxw_chart *self, int8_t overlap)
{
    if (overlap >= -100 && overlap <= 100)
        self->overlap_y1 = overlap;
    else
        LXW_WARN_FORMAT1("chart_set_series_overlap(): Chart series overlap "
                         "'%d' outside Excel range: -100 <= overlap <= 100",
                         overlap);
}

/*
 * Set the option for displaying blank data in a chart.
 */
void
chart_show_blanks_as(lxw_chart *self, uint8_t option)
{
    self->show_blanks_as = option;
}

/*
 * Display data on charts from hidden rows or columns.
 */
void
chart_show_hidden_data(lxw_chart *self)
{
    self->show_hidden_data = LXW_TRUE;
}

/*
 * Set the Bar/Column gap for all data series.
 */
void
chart_set_series_gap(lxw_chart *self, uint16_t gap)
{
    if (gap <= 500)
        self->gap_y1 = gap;
    else
        LXW_WARN_FORMAT1("chart_set_series_gap(): Chart series gap '%d' "
                         "outside Excel range: 0 <= gap <= 500", gap);
}

/*
 * Set the Pie/Doughnut chart rotation: the angle of the first slice.
 */
void
chart_set_rotation(lxw_chart *self, uint16_t rotation)
{
    if (rotation <= 360)
        self->rotation = rotation;
    else
        LXW_WARN_FORMAT1("chart_set_rotation(): Chart rotation '%d' outside "
                         "Excel range: 0 <= rotation <= 360", rotation);
}

/*
 * Set the Doughnut chart hole size.
 */
void
chart_set_hole_size(lxw_chart *self, uint8_t size)
{
    if (size >= 10 && size <= 90)
        self->hole_size = size;
    else
        LXW_WARN_FORMAT1("chart_set_hole_size(): Hole size '%d' outside "
                         "Excel range: 10 <= size <= 90", size);
}
