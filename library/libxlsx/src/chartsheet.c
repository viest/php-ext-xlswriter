/*****************************************************************************
 * chartsheet - A library for creating Excel XLSX chartsheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/chartsheet.h"
#include "lxlsx/utility.h"

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new chartsheet object.
 */
lxw_chartsheet *
lxw_chartsheet_new(lxw_worksheet_init_data *init_data)
{
    lxw_chartsheet *chartsheet = calloc(1, sizeof(lxw_chartsheet));
    GOTO_LABEL_ON_MEM_ERROR(chartsheet, mem_error);

    /* Use an embedded worksheet instance to write XML records that are
     * shared with worksheet.c. */
    chartsheet->worksheet = lxw_worksheet_new(NULL);
    GOTO_LABEL_ON_MEM_ERROR(chartsheet->worksheet, mem_error);

    if (init_data) {
        chartsheet->name = init_data->name;
        chartsheet->quoted_name = init_data->quoted_name;
        chartsheet->tmpdir = init_data->tmpdir;
        chartsheet->index = init_data->index;
        chartsheet->hidden = init_data->hidden;
        chartsheet->active_sheet = init_data->active_sheet;
        chartsheet->first_sheet = init_data->first_sheet;
    }

    chartsheet->worksheet->is_chartsheet = LXW_TRUE;
    chartsheet->worksheet->zoom_scale_normal = LXW_FALSE;
    chartsheet->worksheet->orientation = LXW_LANDSCAPE;

    return chartsheet;

mem_error:
    lxw_chartsheet_free(chartsheet);
    return NULL;
}

/*
 * Free a chartsheet object.
 */
void
lxw_chartsheet_free(lxw_chartsheet *chartsheet)
{
    if (!chartsheet)
        return;

    lxw_worksheet_free(chartsheet->worksheet);
    free((void *) chartsheet->name);
    free((void *) chartsheet->quoted_name);
    free(chartsheet);
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
_chartsheet_xml_declaration(lxw_chartsheet *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <chartsheet> element.
 */
STATIC void
_chartsheet_write_chartsheet(lxw_chartsheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] = "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main";
    char xmlns_r[] = "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxw_xml_start_tag(self->file, "chartsheet", &attributes);
    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetPr> element.
 */
STATIC void
_chartsheet_write_sheet_pr(lxw_chartsheet *self)
{
    lxw_worksheet_write_sheet_pr(self->worksheet);
}

/*
 * Write the <sheetViews> element.
 */
STATIC void
_chartsheet_write_sheet_views(lxw_chartsheet *self)
{
    lxw_worksheet_write_sheet_views(self->worksheet);
}

/*
 * Write the <pageMargins> element.
 */
STATIC void
_chartsheet_write_page_margins(lxw_chartsheet *self)
{
    lxw_worksheet_write_page_margins(self->worksheet);
}

/*
 * Write the <drawing> elements.
 */
STATIC void
_chartsheet_write_drawings(lxw_chartsheet *self)
{
    lxw_worksheet_write_drawings(self->worksheet);
}

/*
 * Write the <sheetProtection> element.
 */
STATIC void
_chartsheet_write_sheet_protection(lxw_chartsheet *self)
{
    lxw_worksheet_write_sheet_protection(self->worksheet, &self->protection);
}

/*
 * Write the <pageSetup> element.
 */
STATIC void
_chartsheet_write_page_setup(lxw_chartsheet *self)
{
    lxw_worksheet_write_page_setup(self->worksheet);
}

/*
 * Write the <headerFooter> element.
 */
STATIC void
_chartsheet_write_header_footer(lxw_chartsheet *self)
{
    lxw_worksheet_write_header_footer(self->worksheet);
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
lxw_chartsheet_assemble_xml_file(lxw_chartsheet *self)
{
    /* Set the embedded worksheet filehandle to the same as the chartsheet. */
    self->worksheet->file = self->file;

    /* Write the XML declaration. */
    _chartsheet_xml_declaration(self);

    /* Write the chartsheet element. */
    _chartsheet_write_chartsheet(self);

    /* Write the sheetPr element. */
    _chartsheet_write_sheet_pr(self);

    /* Write the sheetViews element. */
    _chartsheet_write_sheet_views(self);

    /* Write the sheetProtection element. */
    _chartsheet_write_sheet_protection(self);

    /* Write the pageMargins element. */
    _chartsheet_write_page_margins(self);

    /* Write the chartsheet page setup. */
    _chartsheet_write_page_setup(self);

    /* Write the headerFooter element. */
    _chartsheet_write_header_footer(self);

    /* Write the drawing element. */
    _chartsheet_write_drawings(self);

    lxw_xml_end_tag(self->file, "chartsheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Set a chartsheet chart, with options.
 */
lxw_error
chartsheet_set_chart_opt(lxw_chartsheet *self,
                         lxw_chart *chart, lxw_chart_options *user_options)
{
    lxw_object_properties *object_props;
    lxw_chart_series *series;

    if (!chart) {
        LXW_WARN("chartsheet_set_chart()/_opt(): chart must be non-NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the chart isn't being used more than once. */
    if (chart->in_use) {
        LXW_WARN("chartsheet_set_chart()/_opt(): the same chart object "
                 "cannot be set for a chartsheet more than once.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a data series. */
    if (STAILQ_EMPTY(chart->series_list)) {
        LXW_WARN("chartsheet_set_chart()/_opt(): chart must have a series.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a 'values' series. */
    STAILQ_FOREACH(series, chart->series_list, list_pointers) {
        if (!series->values->formula && !series->values->sheetname) {
            LXW_WARN("chartsheet_set_chart()/_opt(): chart must have a "
                     "'values' series.");

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }

    /* Create a new object to hold the chart image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    RETURN_ON_MEM_ERROR(object_props, LXW_ERROR_MEMORY_MALLOC_FAILED);

    if (user_options) {
        object_props->x_offset = user_options->x_offset;
        object_props->y_offset = user_options->y_offset;
        object_props->x_scale = user_options->x_scale;
        object_props->y_scale = user_options->y_scale;
    }

    object_props->width = 480;
    object_props->height = 288;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    /* Store chart references so they can be ordered in the workbook. */
    object_props->chart = chart;

    /* Store the chart data in the embedded worksheet. */
    STAILQ_INSERT_TAIL(self->worksheet->chart_data, object_props,
                       list_pointers);

    chart->in_use = LXW_TRUE;
    chart->is_chartsheet = LXW_TRUE;

    chart->is_protected = self->is_protected;

    self->chart = chart;

    return LXW_NO_ERROR;
}

/*
 * Set a chartsheet charts.
 */
lxw_error
chartsheet_set_chart(lxw_chartsheet *self, lxw_chart *chart)
{
    return chartsheet_set_chart_opt(self, chart, NULL);
}

/*
 * Set this chartsheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
void
chartsheet_select(lxw_chartsheet *self)
{
    self->selected = LXW_TRUE;

    /* Selected worksheet can't be hidden. */
    self->hidden = LXW_FALSE;
}

/*
 * Set this chartsheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
void
chartsheet_activate(lxw_chartsheet *self)
{
    self->worksheet->selected = LXW_TRUE;
    self->worksheet->active = LXW_TRUE;

    /* Active worksheet can't be hidden. */
    self->worksheet->hidden = LXW_FALSE;

    *self->active_sheet = self->index;
}

/*
 * Set this chartsheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
void
chartsheet_set_first_sheet(lxw_chartsheet *self)
{
    /* Active worksheet can't be hidden. */
    self->hidden = LXW_FALSE;

    *self->first_sheet = self->index;
}

/*
 * Hide this chartsheet.
 */
void
chartsheet_hide(lxw_chartsheet *self)
{
    self->hidden = LXW_TRUE;

    /* A hidden worksheet shouldn't be active or selected. */
    self->selected = LXW_FALSE;

    /* If this is active_sheet or first_sheet reset the workbook value. */
    if (*self->first_sheet == self->index)
        *self->first_sheet = 0;

    if (*self->active_sheet == self->index)
        *self->active_sheet = 0;
}

/*
 * Set the color of the chartsheet tab.
 */
void
chartsheet_set_tab_color(lxw_chartsheet *self, lxw_color_t color)
{
    self->worksheet->tab_color = color;
}

/*
 * Set the chartsheet protection flags to prevent modification of chartsheet
 * objects.
 */
void
chartsheet_protect(lxw_chartsheet *self, const char *password,
                   lxw_protection *options)
{
    struct lxw_protection_obj *protect = &self->protection;

    /* Copy any user parameters to the internal structure. */
    if (options) {
        protect->objects = options->no_objects;
        protect->no_content = options->no_content;
    }
    else {
        protect->objects = LXW_FALSE;
        protect->no_content = LXW_FALSE;
    }

    if (password) {
        uint16_t hash = lxw_hash_password(password);
        lxw_snprintf(protect->hash, 5, "%X", hash);
    }
    else {
        if (protect->objects && protect->no_content)
            return;
    }

    protect->no_sheet = LXW_TRUE;
    protect->scenarios = LXW_TRUE;
    protect->is_configured = LXW_TRUE;

    if (self->chart)
        self->chart->is_protected = LXW_TRUE;
    else
        self->is_protected = LXW_TRUE;
}

/*
 * Set the chartsheet zoom factor.
 */
void
chartsheet_set_zoom(lxw_chartsheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400) {
        LXW_WARN("chartsheet_set_zoom(): "
                 "Zoom factor scale outside range: 10 <= zoom <= 400.");
        return;
    }

    self->worksheet->zoom = scale;
}

/*
 * Set the page orientation as portrait.
 */
void
chartsheet_set_portrait(lxw_chartsheet *self)
{
    worksheet_set_portrait(self->worksheet);
}

/*
 * Set the page orientation as landscape.
 */
void
chartsheet_set_landscape(lxw_chartsheet *self)
{
    worksheet_set_landscape(self->worksheet);
}

/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
void
chartsheet_set_paper(lxw_chartsheet *self, uint8_t paper_size)
{
    worksheet_set_paper(self->worksheet, paper_size);
}

/*
 * Set all the page margins in inches.
 */
void
chartsheet_set_margins(lxw_chartsheet *self, double left, double right,
                       double top, double bottom)
{
    worksheet_set_margins(self->worksheet, left, right, top, bottom);
}

/*
 * Set the page header caption and options.
 */
lxw_error
chartsheet_set_header_opt(lxw_chartsheet *self, const char *string,
                          lxw_header_footer_options *options)
{
    return worksheet_set_header_opt(self->worksheet, string, options);
}

/*
 * Set the page footer caption and options.
 */
lxw_error
chartsheet_set_footer_opt(lxw_chartsheet *self, const char *string,
                          lxw_header_footer_options *options)
{
    return worksheet_set_footer_opt(self->worksheet, string, options);
}

/*
 * Set the page header caption.
 */
lxw_error
chartsheet_set_header(lxw_chartsheet *self, const char *string)
{
    return chartsheet_set_header_opt(self, string, NULL);
}

/*
 * Set the page footer caption.
 */
lxw_error
chartsheet_set_footer(lxw_chartsheet *self, const char *string)
{
    return chartsheet_set_footer_opt(self, string, NULL);
}
