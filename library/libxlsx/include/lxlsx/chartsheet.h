/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * chartsheet - A libxlsxwriter library for creating Excel XLSX chartsheet files.
 *
 */

/**
 * @page chartsheet_page The Chartsheet object
 *
 * The Chartsheet object represents an Excel chartsheet, which is a type of
 * worksheet that only contains a chart. The Chartsheet object handles
 * operations such as adding a chart and setting the page layout.
 *
 * See @ref chartsheet.h for full details of the functionality.
 *
 * @file chartsheet.h
 *
 * @brief Functions related to adding data and formatting to a chartsheet.
 *
 * The Chartsheet object represents an Excel chartsheet. It handles operations
 * such as adding a chart and setting the page layout.
 *
 * A Chartsheet object isn't created directly. Instead a chartsheet is created
 * by calling the workbook_add_chartsheet() function from a Workbook object. A
 * chartsheet object functions as a worksheet and not as a chart. In order to
 * have it display data a #lxw_chart object must be created and added to the
 * chartsheet:
 *
 * @code
 *     #include "lxlsx.h"
 *
 *     int main() {
 *
 *         lxw_workbook   *workbook   = new_workbook("chartsheet.xlsx");
 *         lxw_worksheet  *worksheet  = workbook_add_worksheet(workbook, NULL);
 *         lxw_chartsheet *chartsheet = workbook_add_chartsheet(workbook, NULL);
 *         lxw_chart      *chart      = workbook_add_chart(workbook, LXW_CHART_BAR);
 *
 *         //... Set up the chart.
 *
 *         // Add the chart to the chartsheet.
 *         return workbook_close(workbook);
 *
 *     }
 * @endcode
 *
 * @image html chartsheet.png
 *
 * The data for the chartsheet chart must be contained on a separate
 * worksheet. That is why it is always created in conjunction with at least
 * one data worksheet, as shown above.
 */

#ifndef __LXW_CHARTSHEET_H__
#define __LXW_CHARTSHEET_H__

#include <stdint.h>

#include "common.h"
#include "worksheet.h"
#include "drawing.h"
#include "utility.h"

/**
 * @brief Struct to represent an Excel chartsheet.
 *
 * The members of the lxw_chartsheet struct aren't modified directly. Instead
 * the chartsheet properties are set by calling the functions shown in
 * chartsheet.h.
 */
typedef struct lxw_chartsheet {

    FILE *file;
    lxw_worksheet *worksheet;
    lxw_chart *chart;

    struct lxw_protection_obj protection;
    uint8_t is_protected;

    const char *name;
    const char *quoted_name;
    const char *tmpdir;
    uint16_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    uint16_t rel_count;

    STAILQ_ENTRY (lxw_chartsheet) list_pointers;

} lxw_chartsheet;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Insert a chart object into a chartsheet.
 *
 * @param chartsheet   Pointer to a lxw_chartsheet instance to be updated.
 * @param chart        A #lxw_chart object created via workbook_add_chart().
 *
 * @return A #lxw_error code.
 *
 * The `%chartsheet_set_chart()` function can be used to insert a chart into a
 * chartsheet. The chart object must be created first using the
 * `workbook_add_chart()` function and configured using the @ref chart.h
 * functions.
 *
 * @code
 *     // Create the chartsheet.
 *     lxw_chartsheet *chartsheet = workbook_add_chartsheet(workbook, NULL);
 *
 *     // Create a chart object.
 *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_LINE);
 *
 *     // Add a data series to the chart.
 *     chart_add_series(chart, NULL, "=Sheet1!$A$1:$A$6");
 *
 *     // Insert the chart into the chartsheet.
 *     chartsheet_set_chart(chartsheet, chart);
 * @endcode
 *
 * @image html chartsheet2.png
 *
 * **Note:**
 *
 * A chart may only be inserted once into a chartsheet or a worksheet. If
 * several similar charts are required then each one must be created
 * separately.
 *
 */
lxw_error chartsheet_set_chart(lxw_chartsheet *chartsheet, lxw_chart *chart);

/* Not currently required since scale options aren't useful in a chartsheet. */
lxw_error chartsheet_set_chart_opt(lxw_chartsheet *chartsheet,
                                   lxw_chart *chart,
                                   lxw_chart_options *user_options);

/**
 * @brief Make a chartsheet the active, i.e., visible chartsheet.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * The `%chartsheet_activate()` function is used to specify which chartsheet
 * is initially visible in a multi-sheet workbook:
 *
 * @code
 *     lxw_worksheet  *worksheet1  = workbook_add_worksheet(workbook, NULL);
 *     lxw_chartsheet *chartsheet1 = workbook_add_chartsheet(workbook, NULL);
 *     lxw_chartsheet *chartsheet2 = workbook_add_chartsheet(workbook, NULL);
 *     lxw_chartsheet *chartsheet3 = workbook_add_chartsheet(workbook, NULL);
 *
 *     chartsheet_activate(chartsheet3);
 * @endcode
 *
 * @image html chartsheet_activate.png
 *
 * More than one chartsheet can be selected via the `chartsheet_select()`
 * function, see below, however only one chartsheet can be active.
 *
 * The default active chartsheet is the first chartsheet.
 *
 * See also `worksheet_activate()`.
 *
 */
void chartsheet_activate(lxw_chartsheet *chartsheet);

/**
 * @brief Set a chartsheet tab as selected.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * The `%chartsheet_select()` function is used to indicate that a chartsheet
 * is selected in a multi-sheet workbook:
 *
 * @code
 *     chartsheet_activate(chartsheet1);
 *     chartsheet_select(chartsheet2);
 *     chartsheet_select(chartsheet3);
 *
 * @endcode
 *
 * A selected chartsheet has its tab highlighted. Selecting chartsheets is a
 * way of grouping them together so that, for example, several chartsheets
 * could be printed in one go. A chartsheet that has been activated via the
 * `chartsheet_activate()` function will also appear as selected.
 *
 * See also `worksheet_select()`.
 *
 */
void chartsheet_select(lxw_chartsheet *chartsheet);

/**
 * @brief Hide the current chartsheet.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * The `%chartsheet_hide()` function is used to hide a chartsheet:
 *
 * @code
 *     chartsheet_hide(chartsheet2);
 * @endcode
 *
 * You may wish to hide a chartsheet in order to avoid confusing a user with
 * intermediate data or calculations.
 *
 * @image html hide_sheet.png
 *
 * A hidden chartsheet can not be activated or selected so this function is
 * mutually exclusive with the `chartsheet_activate()` and
 * `chartsheet_select()` functions. In addition, since the first chartsheet
 * will default to being the active chartsheet, you cannot hide the first
 * chartsheet without activating another sheet:
 *
 * @code
 *     chartsheet_activate(chartsheet2);
 *     chartsheet_hide(chartsheet1);
 * @endcode
 *
 * See also `worksheet_hide()`.
 *
 */
void chartsheet_hide(lxw_chartsheet *chartsheet);

/**
 * @brief Set current chartsheet as the first visible sheet tab.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * The `chartsheet_activate()` function determines which chartsheet is
 * initially selected.  However, if there are a large number of chartsheets the
 * selected chartsheet may not appear on the screen. To avoid this you can
 * select the leftmost visible chartsheet tab using
 * `%chartsheet_set_first_sheet()`:
 *
 * @code
 *     chartsheet_set_first_sheet(chartsheet19); // First visible chartsheet tab.
 *     chartsheet_activate(chartsheet20);        // First visible chartsheet.
 * @endcode
 *
 * This function is not required very often. The default value is the first
 * chartsheet.
 *
 * See also `worksheet_set_first_sheet()`.
 *
 */
void chartsheet_set_first_sheet(lxw_chartsheet *chartsheet);

/**
 * @brief Set the color of the chartsheet tab.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param color      The tab color.
 *
 * The `%chartsheet_set_tab_color()` function is used to change the color of
 * the chartsheet tab:
 *
 * @code
 *      chartsheet_set_tab_color(chartsheet1, LXW_COLOR_RED);
 *      chartsheet_set_tab_color(chartsheet2, LXW_COLOR_GREEN);
 *      chartsheet_set_tab_color(chartsheet3, 0xFF9900); // Orange.
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * See also `worksheet_set_tab_color()`.
 */
void chartsheet_set_tab_color(lxw_chartsheet *chartsheet, lxw_color_t color);

/**
 * @brief Protect elements of a chartsheet from modification.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param password   A chartsheet password.
 * @param options    Chartsheet elements to protect.
 *
 * The `%chartsheet_protect()` function protects chartsheet elements from
 * modification:
 *
 * @code
 *     chartsheet_protect(chartsheet, "Some Password", options);
 * @endcode
 *
 * The `password` and lxw_protection pointer are both optional:
 *
 * @code
 *     chartsheet_protect(chartsheet2, NULL,       my_options);
 *     chartsheet_protect(chartsheet3, "password", NULL);
 *     chartsheet_protect(chartsheet4, "password", my_options);
 * @endcode
 *
 * Passing a `NULL` password is the same as turning on protection without a
 * password. Passing a `NULL` password and `NULL` options had no effect on
 * chartsheets.
 *
 * You can specify which chartsheet elements you wish to protect by passing a
 * lxw_protection pointer in the `options` argument. In Excel chartsheets only
 * have two protection options:
 *
 *     no_content
 *     no_objects
 *
 * All parameters are off by default. Individual elements can be protected as
 * follows:
 *
 * @code
 *     lxw_protection options = {
 *         .no_content  = 1,
 *         .no_objects  = 1,
 *     };
 *
 *     chartsheet_protect(chartsheet, NULL, &options);
 *
 * @endcode
 *
 * See also worksheet_protect().
 *
 * **Note:** Sheet level passwords in Excel offer **very** weak
 * protection. They don't encrypt your data and are very easy to
 * deactivate. Full workbook encryption is not supported by `libxlsxwriter`
 * since it requires a completely different file format.
 */
void chartsheet_protect(lxw_chartsheet *chartsheet, const char *password,
                        lxw_protection *options);

/**
 * @brief Set the chartsheet zoom factor.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param scale      Chartsheet zoom factor.
 *
 * Set the chartsheet zoom factor in the range `10 <= zoom <= 400`:
 *
 * @code
 *     chartsheet_set_zoom(chartsheet, 75);
 * @endcode
 *
 * The default zoom factor is 100. It isn't possible to set the zoom to
 * "Selection" because it is calculated by Excel at run-time.
 *
 * See also `worksheet_set_zoom()`.
 */
void chartsheet_set_zoom(lxw_chartsheet *chartsheet, uint16_t scale);

/**
 * @brief Set the page orientation as landscape.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * This function is used to set the orientation of a chartsheet's printed page
 * to landscape. The default chartsheet orientation is landscape, so this
 * function isn't generally required:
 *
 * @code
 *     chartsheet_set_landscape(chartsheet);
 * @endcode
 */
void chartsheet_set_landscape(lxw_chartsheet *chartsheet);

/**
 * @brief Set the page orientation as portrait.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 *
 * This function is used to set the orientation of a chartsheet's printed page
 * to portrait:
 *
 * @code
 *     chartsheet_set_portrait(chartsheet);
 * @endcode
 */
void chartsheet_set_portrait(lxw_chartsheet *chartsheet);

/**
 * @brief Set the paper type for printing.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param paper_type The Excel paper format type.
 *
 * This function is used to set the paper format for the printed output of a
 * chartsheet:
 *
 * @code
 *     chartsheet_set_paper(chartsheet1, 1);  // US Letter
 *     chartsheet_set_paper(chartsheet2, 9);  // A4
 * @endcode
 *
 * If you do not specify a paper type the chartsheet will print using the
 * printer's default paper style.
 *
 * See `worksheet_set_paper()` for a full list of available paper sizes.
 */
void chartsheet_set_paper(lxw_chartsheet *chartsheet, uint8_t paper_type);

/**
 * @brief Set the chartsheet margins for the printed page.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param left       Left margin in inches.   Excel default is 0.7.
 * @param right      Right margin in inches.  Excel default is 0.7.
 * @param top        Top margin in inches.    Excel default is 0.75.
 * @param bottom     Bottom margin in inches. Excel default is 0.75.
 *
 * The `%chartsheet_set_margins()` function is used to set the margins of the
 * chartsheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    chartsheet_set_margins(chartsheet, 1.3, 1.2, -1, -1);
 * @endcode
 *
 */
void chartsheet_set_margins(lxw_chartsheet *chartsheet, double left,
                            double right, double top, double bottom);

/**
 * @brief Set the printed page header caption.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param string     The header string.
 *
 * @return A #lxw_error code.
 *
 * Headers and footers are generated using a string which is a combination of
 * plain text and control characters
 *
 * @code
 *     chartsheet_set_header(chartsheet, "&LHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    | Hello                                                         |
 *     //    |                                                               |
 *
 *
 *     chartsheet_set_header(chartsheet, "&CHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                          Hello                                |
 *     //    |                                                               |
 *
 *
 *     chartsheet_set_header(chartsheet, "&RHello");
 *
 *     //     ---------------------------------------------------------------
 *     //    |                                                               |
 *     //    |                                                         Hello |
 *     //    |                                                               |
 *
 *
 * @endcode
 *
 * See `worksheet_set_header()` for a full explanation of the syntax of
 * Excel's header formatting and control characters.
 *
 */
lxw_error chartsheet_set_header(lxw_chartsheet *chartsheet,
                                const char *string);

/**
 * @brief Set the printed page footer caption.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param string     The footer string.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as chartsheet_set_header().
 *
 */
lxw_error chartsheet_set_footer(lxw_chartsheet *chartsheet,
                                const char *string);

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param string     The header string.
 * @param options    Header options.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as chartsheet_set_header() with an
 * additional parameter to specify options for the header.
 *
 * Currently, the only available option is the header margin:
 *
 * @code
 *
 *    lxw_header_footer_options header_options = { 0.2 };
 *
 *    chartsheet_set_header_opt(chartsheet, "Some text", &header_options);
 *
 * @endcode
 *
 */
lxw_error chartsheet_set_header_opt(lxw_chartsheet *chartsheet,
                                    const char *string,
                                    lxw_header_footer_options *options);

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param chartsheet Pointer to a lxw_chartsheet instance to be updated.
 * @param string     The footer string.
 * @param options    Footer options.
 *
 * @return A #lxw_error code.
 *
 * The syntax of this function is the same as chartsheet_set_header_opt().
 *
 */
lxw_error chartsheet_set_footer_opt(lxw_chartsheet *chartsheet,
                                    const char *string,
                                    lxw_header_footer_options *options);

lxw_chartsheet *lxw_chartsheet_new(lxw_worksheet_init_data *init_data);
void lxw_chartsheet_free(lxw_chartsheet *chartsheet);
void lxw_chartsheet_assemble_xml_file(lxw_chartsheet *chartsheet);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _chartsheet_xml_declaration(lxw_chartsheet *chartsheet);
STATIC void _chartsheet_write_sheet_protection(lxw_chartsheet *chartsheet);
#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_CHARTSHEET_H__ */
