/*****************************************************************************
 * chartsheet - A library for creating Excel XLSX chartsheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/chartsheet.h"
#include "libxlsx/utility.h"

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
lxlsx_chartsheet *
lxlsx_chartsheet_new(lxlsx_worksheet_init_data *init_data)
{
    lxlsx_chartsheet *chartsheet = calloc(1, sizeof(lxlsx_chartsheet));
    GOTO_LABEL_ON_MEM_ERROR(chartsheet, mem_error);

    /* Use an embedded worksheet instance to write XML records that are
     * shared with worksheet.c. */
    chartsheet->worksheet = lxlsx_worksheet_new(NULL);
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

    chartsheet->worksheet->is_chartsheet = LXLSX_TRUE;
    chartsheet->worksheet->zoom_scale_normal = LXLSX_FALSE;
    chartsheet->worksheet->orientation = LXLSX_LANDSCAPE;

    return chartsheet;

mem_error:
    lxlsx_chartsheet_free(chartsheet);
    return NULL;
}

/*
 * Free a chartsheet object.
 */
void
lxlsx_chartsheet_free(lxlsx_chartsheet *chartsheet)
{
    if (!chartsheet)
        return;

    lxlsx_worksheet_free(chartsheet->worksheet);
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
_chartsheet_xml_declaration(lxlsx_chartsheet *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <chartsheet> element.
 */
STATIC void
_chartsheet_write_chartsheet(lxlsx_chartsheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main";
    char xmlns_r[] = "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxlsx_xml_start_tag(self->file, "chartsheet", &attributes);
    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetPr> element.
 */
STATIC void
_chartsheet_write_sheet_pr(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_sheet_pr(self->worksheet);
}

/*
 * Write the <sheetViews> element.
 */
STATIC void
_chartsheet_write_sheet_views(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_sheet_views(self->worksheet);
}

/*
 * Write the <pageMargins> element.
 */
STATIC void
_chartsheet_write_page_margins(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_page_margins(self->worksheet);
}

/*
 * Write the <drawing> elements.
 */
STATIC void
_chartsheet_write_drawings(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_drawings(self->worksheet);
}

/*
 * Write the <sheetProtection> element.
 */
STATIC void
_chartsheet_write_sheet_protection(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_sheet_protection(self->worksheet, &self->protection);
}

/*
 * Write the <pageSetup> element.
 */
STATIC void
_chartsheet_write_page_setup(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_page_setup(self->worksheet);
}

/*
 * Write the <headerFooter> element.
 */
STATIC void
_chartsheet_write_header_footer(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_write_header_footer(self->worksheet);
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
lxlsx_chartsheet_assemble_xml_file(lxlsx_chartsheet *self)
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

    lxlsx_xml_end_tag(self->file, "chartsheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Set a chartsheet chart, with options.
 */
lxlsx_error
lxlsx_chartsheet_set_chart_opt(lxlsx_chartsheet *self,
                         lxlsx_chart *chart, lxlsx_chart_options *user_options)
{
    lxlsx_object_properties *object_props;
    lxlsx_chart_series *series;

    if (!chart) {
        LXLSX_WARN("lxlsx_chartsheet_set_chart()/_opt(): chart must be non-NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the chart isn't being used more than once. */
    if (chart->in_use) {
        LXLSX_WARN("lxlsx_chartsheet_set_chart()/_opt(): the same chart object "
                 "cannot be set for a chartsheet more than once.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a data series. */
    if (STAILQ_EMPTY(chart->series_list)) {
        LXLSX_WARN("lxlsx_chartsheet_set_chart()/_opt(): chart must have a series.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a 'values' series. */
    STAILQ_FOREACH(series, chart->series_list, list_pointers) {
        if (!series->values->formula && !series->values->sheetname) {
            LXLSX_WARN("lxlsx_chartsheet_set_chart()/_opt(): chart must have a "
                     "'values' series.");

            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }
    }

    /* Create a new object to hold the chart image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    RETURN_ON_MEM_ERROR(object_props, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

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
    STAILQ_INSERT_TAIL(self->worksheet->lxlsx_chart_data, object_props,
                       list_pointers);

    chart->in_use = LXLSX_TRUE;
    chart->is_chartsheet = LXLSX_TRUE;

    chart->is_protected = self->is_protected;

    self->chart = chart;

    return LXLSX_NO_ERROR;
}

/*
 * Set a chartsheet charts.
 */
lxlsx_error
lxlsx_chartsheet_set_chart(lxlsx_chartsheet *self, lxlsx_chart *chart)
{
    return lxlsx_chartsheet_set_chart_opt(self, chart, NULL);
}

/*
 * Set this chartsheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
void
lxlsx_chartsheet_select(lxlsx_chartsheet *self)
{
    self->selected = LXLSX_TRUE;

    /* Selected worksheet can't be hidden. */
    self->hidden = LXLSX_FALSE;
}

/*
 * Set this chartsheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
void
lxlsx_chartsheet_activate(lxlsx_chartsheet *self)
{
    self->worksheet->selected = LXLSX_TRUE;
    self->worksheet->active = LXLSX_TRUE;

    /* Active worksheet can't be hidden. */
    self->worksheet->hidden = LXLSX_FALSE;

    *self->active_sheet = self->index;
}

/*
 * Set this chartsheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
void
lxlsx_chartsheet_set_first_sheet(lxlsx_chartsheet *self)
{
    /* Active worksheet can't be hidden. */
    self->hidden = LXLSX_FALSE;

    *self->first_sheet = self->index;
}

/*
 * Hide this chartsheet.
 */
void
lxlsx_chartsheet_hide(lxlsx_chartsheet *self)
{
    self->hidden = LXLSX_TRUE;

    /* A hidden worksheet shouldn't be active or selected. */
    self->selected = LXLSX_FALSE;

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
lxlsx_chartsheet_set_tab_color(lxlsx_chartsheet *self, lxlsx_color_t color)
{
    self->worksheet->tab_color = color;
}

/*
 * Set the chartsheet protection flags to prevent modification of chartsheet
 * objects.
 */
void
lxlsx_chartsheet_protect(lxlsx_chartsheet *self, const char *password,
                   lxlsx_protection *options)
{
    struct lxlsx_protection_obj *protect = &self->protection;

    /* Copy any user parameters to the internal structure. */
    if (options) {
        protect->objects = options->no_objects;
        protect->no_content = options->no_content;
    }
    else {
        protect->objects = LXLSX_FALSE;
        protect->no_content = LXLSX_FALSE;
    }

    if (password) {
        uint16_t hash = lxlsx_hash_password(password);
        lxlsx_snprintf(protect->hash, 5, "%X", hash);
    }
    else {
        if (protect->objects && protect->no_content)
            return;
    }

    protect->no_sheet = LXLSX_TRUE;
    protect->scenarios = LXLSX_TRUE;
    protect->is_configured = LXLSX_TRUE;

    if (self->chart)
        self->chart->is_protected = LXLSX_TRUE;
    else
        self->is_protected = LXLSX_TRUE;
}

/*
 * Set the chartsheet zoom factor.
 */
void
lxlsx_chartsheet_set_zoom(lxlsx_chartsheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400) {
        LXLSX_WARN("lxlsx_chartsheet_set_zoom(): "
                 "Zoom factor scale outside range: 10 <= zoom <= 400.");
        return;
    }

    self->worksheet->zoom = scale;
}

/*
 * Set the page orientation as portrait.
 */
void
lxlsx_chartsheet_set_portrait(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_set_portrait(self->worksheet);
}

/*
 * Set the page orientation as landscape.
 */
void
lxlsx_chartsheet_set_landscape(lxlsx_chartsheet *self)
{
    lxlsx_worksheet_set_landscape(self->worksheet);
}

/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
void
lxlsx_chartsheet_set_paper(lxlsx_chartsheet *self, uint8_t paper_size)
{
    lxlsx_worksheet_set_paper(self->worksheet, paper_size);
}

/*
 * Set all the page margins in inches.
 */
void
lxlsx_chartsheet_set_margins(lxlsx_chartsheet *self, double left, double right,
                       double top, double bottom)
{
    lxlsx_worksheet_set_margins(self->worksheet, left, right, top, bottom);
}

/*
 * Set the page header caption and options.
 */
lxlsx_error
lxlsx_chartsheet_set_header_opt(lxlsx_chartsheet *self, const char *string,
                          lxlsx_header_footer_options *options)
{
    return lxlsx_worksheet_set_header_opt(self->worksheet, string, options);
}

/*
 * Set the page footer caption and options.
 */
lxlsx_error
lxlsx_chartsheet_set_footer_opt(lxlsx_chartsheet *self, const char *string,
                          lxlsx_header_footer_options *options)
{
    return lxlsx_worksheet_set_footer_opt(self->worksheet, string, options);
}

/*
 * Set the page header caption.
 */
lxlsx_error
lxlsx_chartsheet_set_header(lxlsx_chartsheet *self, const char *string)
{
    return lxlsx_chartsheet_set_header_opt(self, string, NULL);
}

/*
 * Set the page footer caption.
 */
lxlsx_error
lxlsx_chartsheet_set_footer(lxlsx_chartsheet *self, const char *string)
{
    return lxlsx_chartsheet_set_footer_opt(self, string, NULL);
}
