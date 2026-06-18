/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 */

/**
 * @page lxlsx_workbook_page The Workbook object
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * See @ref workbook.h for full details of the functionality.
 *
 * @file workbook.h
 *
 * @brief Functions related to creating an Excel xlsx workbook.
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
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
 * @image html workbook01.png
 *
 */
#ifndef __LXLSX_WORKBOOK_H__
#define __LXLSX_WORKBOOK_H__

#include <stdint.h>
#include <stdio.h>
#include <errno.h>

#include "worksheet.h"
#include "chartsheet.h"
#include "chart.h"
#include "shared_strings.h"
#include "hash_table.h"
#include "common.h"
#include "edit.h"

#define LXLSX_DEFINED_NAME_LENGTH 128

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxlsx_worksheet_names, lxlsx_worksheet_name);
RB_HEAD(lxlsx_chartsheet_names, lxlsx_chartsheet_name);
RB_HEAD(lxlsx_image_md5s, lxlsx_image_md5);

/* Define the queue.h structs for the workbook lists. */
STAILQ_HEAD(lxlsx_sheets, lxlsx_sheet);
STAILQ_HEAD(lxlsx_worksheets, lxlsx_worksheet);
STAILQ_HEAD(lxlsx_chartsheets, lxlsx_chartsheet);
STAILQ_HEAD(lxlsx_charts, lxlsx_chart);
TAILQ_HEAD(lxlsx_defined_names, lxlsx_defined_name);

/* Struct to hold the 2 sheet types. */
typedef struct lxlsx_sheet {
    uint8_t is_chartsheet;

    union {
        lxlsx_worksheet *worksheet;
        lxlsx_chartsheet *chartsheet;
    } u;

    STAILQ_ENTRY (lxlsx_sheet) list_pointers;
} lxlsx_sheet;

/* Struct to represent a worksheet name/pointer pair. */
typedef struct lxlsx_worksheet_name {
    const char *name;
    lxlsx_worksheet *worksheet;

    RB_ENTRY (lxlsx_worksheet_name) tree_pointers;
} lxlsx_worksheet_name;

/* Struct to represent a chartsheet name/pointer pair. */
typedef struct lxlsx_chartsheet_name {
    const char *name;
    lxlsx_chartsheet *chartsheet;

    RB_ENTRY (lxlsx_chartsheet_name) tree_pointers;
} lxlsx_chartsheet_name;

/* Struct to represent an image MD5/ID pair. */
typedef struct lxlsx_image_md5 {
    uint32_t id;
    char *md5;

    RB_ENTRY (lxlsx_image_md5) tree_pointers;
} lxlsx_image_md5;

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXLSX_RB_GENERATE_WORKSHEET_NAMES(name, type, field, cmp)  \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)          \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)          \
    RB_GENERATE_INSERT(name, type, field, cmp, static)           \
    RB_GENERATE_REMOVE(name, type, field, static)                \
    RB_GENERATE_FIND(name, type, field, cmp, static)             \
    RB_GENERATE_NEXT(name, type, field, static)                  \
    RB_GENERATE_MINMAX(name, type, field, static)                \
    /* Add unused struct to allow adding a semicolon */          \
    struct lxlsx_rb_generate_worksheet_names{int unused;}

#define LXLSX_RB_GENERATE_CHARTSHEET_NAMES(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)          \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)          \
    RB_GENERATE_INSERT(name, type, field, cmp, static)           \
    RB_GENERATE_REMOVE(name, type, field, static)                \
    RB_GENERATE_FIND(name, type, field, cmp, static)             \
    RB_GENERATE_NEXT(name, type, field, static)                  \
    RB_GENERATE_MINMAX(name, type, field, static)                \
    /* Add unused struct to allow adding a semicolon */          \
    struct lxlsx_rb_generate_charsheet_names{int unused;}

#define LXLSX_RB_GENERATE_IMAGE_MD5S(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)          \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)          \
    RB_GENERATE_INSERT(name, type, field, cmp, static)           \
    RB_GENERATE_REMOVE(name, type, field, static)                \
    RB_GENERATE_FIND(name, type, field, cmp, static)             \
    RB_GENERATE_NEXT(name, type, field, static)                  \
    RB_GENERATE_MINMAX(name, type, field, static)                \
    /* Add unused struct to allow adding a semicolon */          \
    struct lxlsx_rb_generate_image_md5s{int unused;}

/**
 * @brief Macro to loop over all the worksheets in a workbook.
 *
 * This macro allows you to loop over all the worksheets that have been
 * added to a workbook. You must provide a lxlsx_worksheet pointer and
 * a pointer to the lxlsx_workbook:
 *
 * @code
 *    lxlsx_workbook  *workbook = lxlsx_workbook_new("test.xlsx");
 *
 *    lxlsx_worksheet *worksheet; // Generic worksheet pointer.
 *
 *    // Worksheet objects used in the program.
 *    lxlsx_worksheet *worksheet1 = lxlsx_workbook_add_worksheet(workbook, NULL);
 *    lxlsx_worksheet *worksheet2 = lxlsx_workbook_add_worksheet(workbook, NULL);
 *    lxlsx_worksheet *worksheet3 = lxlsx_workbook_add_worksheet(workbook, NULL);
 *
 *    // Iterate over the 3 worksheets and perform the same operation on each.
 *    LXLSX_FOREACH_WORKSHEET(worksheet, workbook) {
 *        lxlsx_worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *    }
 * @endcode
 */
#define LXLSX_FOREACH_WORKSHEET(worksheet, workbook) \
    STAILQ_FOREACH((worksheet), (workbook)->worksheets, list_pointers)

/* Struct to represent a defined name. */
typedef struct lxlsx_defined_name {
    int16_t index;
    uint8_t hidden;
    char name[LXLSX_DEFINED_NAME_LENGTH];
    char lxlsx_app_name[LXLSX_DEFINED_NAME_LENGTH];
    char formula[LXLSX_DEFINED_NAME_LENGTH];
    char normalised_name[LXLSX_DEFINED_NAME_LENGTH];
    char normalised_sheetname[LXLSX_DEFINED_NAME_LENGTH];

    /* List pointers for queue.h. */
    TAILQ_ENTRY (lxlsx_defined_name) list_pointers;
} lxlsx_defined_name;

/**
 * Workbook document properties. Set any unused fields to NULL or 0.
 */
typedef struct lxlsx_doc_properties {
    /** The title of the Excel Document. */
    const char *title;

    /** The subject of the Excel Document. */
    const char *subject;

    /** The author of the Excel Document. */
    const char *author;

    /** The manager field of the Excel Document. */
    const char *manager;

    /** The company field of the Excel Document. */
    const char *company;

    /** The category of the Excel Document. */
    const char *category;

    /** The keywords of the Excel Document. */
    const char *keywords;

    /** The comment field of the Excel Document. */
    const char *comments;

    /** The status of the Excel Document. */
    const char *status;

    /** The hyperlink base URL of the Excel Document. */
    const char *hyperlink_base;

    /** The file creation date/time shown in Excel. This defaults to the
     * current time and date if set to 0. If you wish to create files that are
     * binary equivalent (for the same input data) then you should set this
     * creation date/time to a known value. */
    time_t created;

} lxlsx_doc_properties;

/**
 * @brief Workbook options.
 *
 * Optional parameters when creating a new Workbook object via
 * lxlsx_workbook_new_opt().
 *
 * The following properties are supported:
 *
 * - `constant_memory`: This option reduces the amount of data stored in
 *   memory so that large files can be written efficiently. This option is off
 *   by default. See the notes below for limitations when this mode is on.
 *
 * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior to
 *   assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tmpdir` option.
 *
 * - `use_zip64`: Make the zip library use ZIP64 extensions when writing very
 *   large xlsx files to allow the zip container, or individual XML files
 *   within it, to be greater than 4 GB. See [ZIP64 on Wikipedia][zip64_wiki]
 *   for more information. This option is off by default.
 *
 *   [zip64_wiki]: https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64

 * - `output_buffer`: Output to a buffer instead of a file. The buffer must be
 *   freed manually by calling free(). This option can only be used if filename
 *   is NULL.
 *
 * - `output_buffer_size`: Used with output_buffer to get the size of the
 *   created buffer. This option can only be used if filename is NULL.
 *
 * @note In `constant_memory` mode each row of in-memory data is written to
 * disk and then freed when a new row is started via one of the
 * `lxlsx_worksheet_write_*()` functions. Therefore, once this option is active data
 * should be written in sequential row by row order. For this reason
 * `lxlsx_worksheet_merge_range()` and some other row based functionality doesn't
 * work in this mode. See @ref ww_mem_constant for more details.
 *
 * @note Also, in `constant_memory` mode the library uses temp file storage
 * for worksheet data. This can lead to an issue on OSes that map the `/tmp`
 * directory into memory since it is possible to consume the "system" memory
 * even though the "process" memory remains constant. In these cases you
 * should use an alternative temp file location by using the `tmpdir` option
 * shown above. See @ref ww_mem_temp for more details.
 */
typedef struct lxlsx_workbook_options {
    /** Optimize the workbook to use constant memory for worksheets. */
    uint8_t constant_memory;

    /** Directory to use for the temporary files created by libxlsxwriter. */
    const char *tmpdir;

    /** Allow ZIP64 extensions when creating the xlsx file zip container. */
    uint8_t use_zip64;

    /** Output buffer to use instead of writing to a file */
    const char **output_buffer;

    /** Used with output_buffer to get the size of the created buffer */
    size_t *output_buffer_size;
} lxlsx_workbook_options;

/**
 * @brief Struct to represent an Excel workbook.
 *
 * The members of the lxlsx_workbook struct aren't modified directly. Instead
 * the workbook properties are set by calling the functions shown in
 * workbook.h.
 */
typedef struct lxlsx_workbook {

    FILE *file;
    struct lxlsx_sheets *sheets;
    struct lxlsx_worksheets *worksheets;
    struct lxlsx_chartsheets *chartsheets;
    struct lxlsx_worksheet_names *lxlsx_worksheet_names;
    struct lxlsx_chartsheet_names *lxlsx_chartsheet_names;
    struct lxlsx_image_md5s *image_md5s;
    struct lxlsx_image_md5s *embedded_image_md5s;
    struct lxlsx_image_md5s *header_image_md5s;
    struct lxlsx_image_md5s *background_md5s;
    struct lxlsx_charts *charts;
    struct lxlsx_charts *ordered_charts;
    struct lxlsx_formats *formats;
    struct lxlsx_defined_names *defined_names;
    lxlsx_sst *sst;
    lxlsx_doc_properties *properties;
    struct lxlsx_custom_properties *lxlsx_custom_properties;

    char *filename;
    lxlsx_workbook_options options;
    lxlsx_edit_session *edit_session;
    uint8_t is_edit;

    uint16_t num_sheets;
    uint16_t num_worksheets;
    uint16_t num_chartsheets;
    uint16_t first_sheet;
    uint16_t active_sheet;
    uint16_t num_xf_formats;
    uint16_t num_dxf_formats;
    uint16_t num_format_count;
    uint16_t lxlsx_drawing_count;
    uint16_t comment_count;
    uint32_t num_embedded_images;
    uint16_t window_width;
    uint16_t window_height;

    uint16_t font_count;
    uint16_t border_count;
    uint16_t fill_count;
    uint8_t optimize;
    uint16_t max_url_length;
    uint8_t read_only;

    uint8_t has_png;
    uint8_t has_jpeg;
    uint8_t has_bmp;
    uint8_t has_gif;
    uint8_t has_vml;
    uint8_t has_comments;
    uint8_t has_metadata;
    uint8_t has_embedded_images;
    uint8_t has_dynamic_functions;
    uint8_t has_embedded_image_descriptions;

    lxlsx_hash_table *used_xf_formats;
    lxlsx_hash_table *used_dxf_formats;

    char *vba_project;
    char *vba_project_signature;
    char *vba_codename;

    uint8_t use_1904_epoch;

    lxlsx_format *default_url_format;

} lxlsx_workbook;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Create a new workbook object.
 *
 * @param filename The name of the new Excel file to create.
 *
 * @return A lxlsx_workbook instance.
 *
 * The `%lxlsx_workbook_new()` constructor is used to create a new Excel workbook
 * with a given filename:
 *
 * @code
 *     lxlsx_workbook *workbook  = lxlsx_workbook_new("filename.xlsx");
 * @endcode
 *
 * When specifying a filename it is recommended that you use an `.xlsx`
 * extension or Excel will generate a warning when opening the file.
 *
 */
lxlsx_workbook *lxlsx_workbook_new(const char *filename);

/**
 * @brief Create a new workbook object, and set the workbook options.
 *
 * @param filename The name of the new Excel file to create.
 * @param options  Workbook options.
 *
 * @return A lxlsx_workbook instance.
 *
 * This function is the same as the `lxlsx_workbook_new()` constructor but allows
 * additional options to be set.
 *
 * @code
 *    lxlsx_workbook_options options = {.constant_memory = LXLSX_TRUE,
 *                                    .tmpdir = "C:\\Temp",
 *                                    .use_zip64 = LXLSX_FALSE,
 *                                    .output_buffer = NULL,
 *                                    .output_buffer_size = NULL};
 *
 *    lxlsx_workbook  *workbook  = lxlsx_workbook_new_opt("filename.xlsx", &options);
 * @endcode
 *
 * The options that can be set via #lxlsx_workbook_options are:
 *
 * - `constant_memory`: This option reduces the amount of data stored in
 *   memory so that large files can be written efficiently. This option is off
 *   by default. See the note below for limitations when this mode is on.
 *
 * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior to
 *   assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tmpdir` option.
 *
 * - `use_zip64`: Make the zip library use ZIP64 extensions when writing very
 *   large xlsx files to allow the zip container, or individual XML files
 *   within it, to be greater than 4 GB. See [ZIP64 on Wikipedia][zip64_wiki]
 *   for more information. This option is off by default.
 *
 *   [zip64_wiki]: https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64
 *
 * - `output_buffer`: Output to a memory buffer instead of a file. The buffer
 *   must be freed manually by calling `free()`. This option can only be used if
 *   filename is NULL.
 *
 * - `output_buffer_size`: Used with output_buffer to get the size of the
 *   created buffer. This option can only be used if filename is `NULL`.
 *
 * @note In `constant_memory` mode each row of in-memory data is written to
 * disk and then freed when a new row is started via one of the
 * `lxlsx_worksheet_write_*()` functions. Therefore, once this option is active data
 * should be written in sequential row by row order. For this reason
 * `lxlsx_worksheet_merge_range()` and some other row based functionality doesn't
 * work in this mode. See @ref ww_mem_constant for more details.
 *
 * @note Also, in `constant_memory` mode the library uses temp file storage
 * for worksheet data. This can lead to an issue on OSes that map the `/tmp`
 * directory into memory since it is possible to consume the "system" memory
 * even though the "process" memory remains constant. In these cases you
 * should use an alternative temp file location by using the `tmpdir` option
 * shown above. See @ref ww_mem_temp for more details.
 */
lxlsx_workbook *lxlsx_workbook_new_opt(const char *filename,
                               lxlsx_workbook_options *options);

/**
 * @brief Open an existing workbook for in-place style editing.
 *
 * The returned workbook uses the same worksheet write functions as a newly
 * created workbook. Unsupported operations return
 * #LXLSX_ERROR_FEATURE_NOT_SUPPORTED.
 */
lxlsx_workbook *lxlsx_workbook_open(const char *filename);

/**
 * @brief Save an opened workbook without freeing it.
 *
 * Workbooks opened with lxlsx_workbook_open() are saved by applying their
 * edits to the source package snapshot. Newly created workbooks should still
 * be finalized with lxlsx_workbook_close().
 */
lxlsx_error lxlsx_workbook_save_as(lxlsx_workbook *workbook,
                                   const char *path);

uint8_t lxlsx_workbook_is_edit(lxlsx_workbook *workbook);

/**
 * @brief Add a new worksheet to a workbook.
 *
 * @param workbook  Pointer to a lxlsx_workbook instance.
 * @param sheetname Optional worksheet name, defaults to Sheet1, etc.
 *
 * @return A lxlsx_worksheet object.
 *
 * The `%lxlsx_workbook_add_worksheet()` function adds a new worksheet to a workbook.
 *
 * At least one worksheet should be added to a new workbook: The @ref
 * worksheet.h "Worksheet" object is used to write data and configure a
 * worksheet in the workbook.
 *
 * The `sheetname` parameter is optional. If it is `NULL` the default
 * Excel convention will be followed, i.e. Sheet1, Sheet2, etc.:
 *
 * @code
 *     worksheet = lxlsx_workbook_add_worksheet(workbook, NULL  );     // Sheet1
 *     worksheet = lxlsx_workbook_add_worksheet(workbook, "Foglio2");  // Foglio2
 *     worksheet = lxlsx_workbook_add_worksheet(workbook, "Data");     // Data
 *     worksheet = lxlsx_workbook_add_worksheet(workbook, NULL  );     // Sheet4
 *
 * @endcode
 *
 * @image html workbook02.png
 *
 * The worksheet name must be a valid Excel worksheet name, i.e:
 *
 * - The name cannot be blank.
 * - The name is less than or equal to 31 UTF-8 characters.
 * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
 * - The name doesn't start or end with an apostrophe.
 * - The name isn't already in use. (Case insensitive).
 *
 * If any of these errors are encountered the function will return NULL.
 * You can check for valid name using the `lxlsx_workbook_validate_sheet_name()`
 * function.
 *
 * @note You should also avoid using the worksheet name "History" (case
 * insensitive) which is reserved in English language versions of
 * Excel. Non-English versions may have restrictions on the equivalent word.
 */
lxlsx_worksheet *lxlsx_workbook_add_worksheet(lxlsx_workbook *workbook,
                                      const char *sheetname);

/**
 * @brief Add a new chartsheet to a workbook.
 *
 * @param workbook  Pointer to a lxlsx_workbook instance.
 * @param sheetname Optional chartsheet name, defaults to Chart1, etc.
 *
 * @return A lxlsx_chartsheet object.
 *
 * The `%lxlsx_workbook_add_chartsheet()` function adds a new chartsheet to a
 * workbook. The @ref chartsheet.h "Chartsheet" object is like a worksheet
 * except it displays a chart instead of cell data.
 *
 * @image html chartsheet.png
 *
 * The `sheetname` parameter is optional. If it is `NULL` the default
 * Excel convention will be followed, i.e. Chart1, Chart2, etc.:
 *
 * @code
 *     chartsheet = lxlsx_workbook_add_chartsheet(workbook, NULL  );     // Chart1
 *     chartsheet = lxlsx_workbook_add_chartsheet(workbook, "My Chart"); // My Chart
 *     chartsheet = lxlsx_workbook_add_chartsheet(workbook, NULL  );     // Chart3
 *
 * @endcode
 *
 * The chartsheet name must be a valid Excel worksheet name, i.e.:
 *
 * - The name cannot be blank.
 * - The name is less than or equal to 31 UTF-8 characters.
 * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
 * - The name doesn't start or end with an apostrophe.
 * - The name isn't already in use. (Case insensitive).
 *
 * If any of these errors are encountered the function will return NULL.
 * You can check for valid name using the `lxlsx_workbook_validate_sheet_name()`
 * function.
 *
 * @note You should also avoid using the worksheet name "History" (case
 * insensitive) which is reserved in English language versions of
 * Excel. Non-English versions may have restrictions on the equivalent word.
 *
 * At least one worksheet should be added to a new workbook when creating a
 * chartsheet in order to provide data for the chart. The @ref worksheet.h
 * "Worksheet" object is used to write data and configure a worksheet in the
 * workbook.
 */
lxlsx_chartsheet *lxlsx_workbook_add_chartsheet(lxlsx_workbook *workbook,
                                        const char *sheetname);

/**
 * @brief Create a new @ref format.h "Format" object to formats cells in
 *        worksheets.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 *
 * @return A lxlsx_format instance.
 *
 * The `lxlsx_workbook_add_format()` function can be used to create new @ref
 * format.h "Format" objects which are used to apply formatting to a cell.
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
 * @endcode
 *
 * See @ref format.h "the Format object" and @ref working_with_formats
 * sections for more details about Format properties and how to set them.
 *
 */
lxlsx_format *lxlsx_workbook_add_format(lxlsx_workbook *workbook);

/**
 * @brief Create a new chart to be added to a worksheet:
 *
 * @param workbook   Pointer to a lxlsx_workbook instance.
 * @param lxlsx_chart_type The type of chart to be created. See #lxlsx_chart_type.
 *
 * @return A lxlsx_chart object.
 *
 * The `%lxlsx_workbook_add_chart()` function creates a new chart object that can
 * be added to a worksheet:
 *
 * @code
 *     // Create a chart object.
 *     lxlsx_chart *chart = lxlsx_workbook_add_chart(workbook, LXLSX_CHART_COLUMN);
 *
 *     // Add data series to the chart.
 *     lxlsx_chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
 *     lxlsx_chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5");
 *     lxlsx_chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5");
 *
 *     // Insert the chart into the worksheet
 *     lxlsx_worksheet_insert_chart(worksheet, CELL("B7"), chart);
 * @endcode
 *
 * The available chart types are defined in #lxlsx_chart_type. The types of
 * charts that are supported are:
 *
 * | Chart type                               | Description                            |
 * | :--------------------------------------- | :------------------------------------  |
 * | #LXLSX_CHART_AREA                          | Area chart.                            |
 * | #LXLSX_CHART_AREA_STACKED                  | Area chart - stacked.                  |
 * | #LXLSX_CHART_AREA_STACKED_PERCENT          | Area chart - percentage stacked.       |
 * | #LXLSX_CHART_BAR                           | Bar chart.                             |
 * | #LXLSX_CHART_BAR_STACKED                   | Bar chart - stacked.                   |
 * | #LXLSX_CHART_BAR_STACKED_PERCENT           | Bar chart - percentage stacked.        |
 * | #LXLSX_CHART_COLUMN                        | Column chart.                          |
 * | #LXLSX_CHART_COLUMN_STACKED                | Column chart - stacked.                |
 * | #LXLSX_CHART_COLUMN_STACKED_PERCENT        | Column chart - percentage stacked.     |
 * | #LXLSX_CHART_DOUGHNUT                      | Doughnut chart.                        |
 * | #LXLSX_CHART_LINE                          | Line chart.                            |
 * | #LXLSX_CHART_LINE_STACKED                  | Line chart - stacked.                  |
 * | #LXLSX_CHART_LINE_STACKED_PERCENT          | Line chart - percentage stacked.       |
 * | #LXLSX_CHART_PIE                           | Pie chart.                             |
 * | #LXLSX_CHART_SCATTER                       | Scatter chart.                         |
 * | #LXLSX_CHART_SCATTER_STRAIGHT              | Scatter chart - straight.              |
 * | #LXLSX_CHART_SCATTER_STRAIGHT_WITH_MARKERS | Scatter chart - straight with markers. |
 * | #LXLSX_CHART_SCATTER_SMOOTH                | Scatter chart - smooth.                |
 * | #LXLSX_CHART_SCATTER_SMOOTH_WITH_MARKERS   | Scatter chart - smooth with markers.   |
 * | #LXLSX_CHART_RADAR                         | Radar chart.                           |
 * | #LXLSX_CHART_RADAR_WITH_MARKERS            | Radar chart - with markers.            |
 * | #LXLSX_CHART_RADAR_FILLED                  | Radar chart - filled.                  |
 *
 *
 *
 * See @ref chart.h for details.
 */
lxlsx_chart *lxlsx_workbook_add_chart(lxlsx_workbook *workbook, uint8_t lxlsx_chart_type);

/**
 * @brief Close the Workbook object and write the XLSX file.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 *
 * @return A #lxlsx_error.
 *
 * The `%lxlsx_workbook_close()` function closes a Workbook object, writes the Excel
 * file to disk, frees any memory allocated internally to the Workbook and
 * frees the object itself.
 *
 * @code
 *     lxlsx_workbook_close(workbook);
 * @endcode
 *
 * The `%lxlsx_workbook_close()` function returns any #lxlsx_error error codes
 * encountered when creating the Excel file. The error code can be returned
 * from the program main or the calling function:
 *
 * @code
 *     return lxlsx_workbook_close(workbook);
 * @endcode
 *
 */
lxlsx_error lxlsx_workbook_close(lxlsx_workbook *workbook);

/**
 * @brief Set the document properties such as Title, Author etc.
 *
 * @param workbook   Pointer to a lxlsx_workbook instance.
 * @param properties Document properties to set.
 *
 * @return A #lxlsx_error.
 *
 * The `%lxlsx_workbook_set_properties` function can be used to set the document
 * properties of the Excel file created by `libxlsxwriter`. These properties
 * are visible when you use the `Office Button -> Prepare -> Properties`
 * option in Excel and are also available to external applications that read
 * or index windows files.
 *
 * The properties that can be set are:
 *
 * - `title`
 * - `subject`
 * - `author`
 * - `manager`
 * - `company`
 * - `category`
 * - `keywords`
 * - `comments`
 * - `hyperlink_base`
 * - `created`
 *
 * The properties are specified via a `lxlsx_doc_properties` struct. All the
 * fields are all optional. An example of how to create and pass the
 * properties is:
 *
 * @code
 *     // Create a properties structure and set some of the fields.
 *     lxlsx_doc_properties properties = {
 *         .title    = "This is an example spreadsheet",
 *         .subject  = "With document properties",
 *         .author   = "John McNamara",
 *         .manager  = "Dr. Heinz Doofenshmirtz",
 *         .company  = "of Wolves",
 *         .category = "Example spreadsheets",
 *         .keywords = "Sample, Example, Properties",
 *         .comments = "Created with libxlsxwriter",
 *         .status   = "Quo",
 *     };
 *
 *     // Set the properties in the workbook.
 *     lxlsx_workbook_set_properties(workbook, &properties);
 * @endcode
 *
 * @image html doc_properties.png
 *
 * The `created` parameter sets the file creation date/time shown in
 * Excel. This defaults to the current time and date if set to 0. If you wish
 * to create files that are binary equivalent (for the same input data) then
 * you should set this creation date/time to a known value using a `time_t`
 * value.
 *
 */
lxlsx_error lxlsx_workbook_set_properties(lxlsx_workbook *workbook,
                                  lxlsx_doc_properties *properties);

/**
 * @brief Set a custom document text property.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxlsx_error.
 *
 * The `%lxlsx_workbook_set_custom_property_string()` function can be used to set one
 * or more custom document text properties not covered by the standard
 * properties in the `lxlsx_workbook_set_properties()` function above.
 *
 *  For example:
 *
 * @code
 *     lxlsx_workbook_set_custom_property_string(workbook, "Checked by", "Eve");
 * @endcode
 *
 * @image html lxlsx_custom_properties.png
 *
 * There are 4 `lxlsx_workbook_set_custom_property_string_*()` functions for each
 * of the custom property types supported by Excel:
 *
 * - text/string: `lxlsx_workbook_set_custom_property_string()`
 * - number:      `lxlsx_workbook_set_custom_property_number()`
 * - datetime:    `lxlsx_workbook_set_custom_property_datetime()`
 * - boolean:     `lxlsx_workbook_set_custom_property_boolean()`
 *
 * **Note**: the name and value parameters are limited to 255 characters
 * by Excel.
 *
 */
lxlsx_error lxlsx_workbook_set_custom_property_string(lxlsx_workbook *workbook,
                                              const char *name,
                                              const char *value);
/**
 * @brief Set a custom document number property.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxlsx_error.
 *
 * Set a custom document number property.
 * See `lxlsx_workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     lxlsx_workbook_set_custom_property_number(workbook, "Document number", 12345);
 * @endcode
 */
lxlsx_error lxlsx_workbook_set_custom_property_number(lxlsx_workbook *workbook,
                                              const char *name, double value);

/* Undocumented since the user can use lxlsx_workbook_set_custom_property_number().
 * Only implemented for file format completeness and testing.
 */
lxlsx_error lxlsx_workbook_set_custom_property_integer(lxlsx_workbook *workbook,
                                               const char *name,
                                               int32_t value);

/**
 * @brief Set a custom document boolean property.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     The name of the custom property.
 * @param value    The value of the custom property.
 *
 * @return A #lxlsx_error.
 *
 * Set a custom document boolean property.
 * See `lxlsx_workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     lxlsx_workbook_set_custom_property_boolean(workbook, "Has Review", 1);
 * @endcode
 */
lxlsx_error lxlsx_workbook_set_custom_property_boolean(lxlsx_workbook *workbook,
                                               const char *name,
                                               uint8_t value);
/**
 * @brief Set a custom document date or time property.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     The name of the custom property.
 * @param datetime The value of the custom property.
 *
 * @return A #lxlsx_error.
 *
 * Set a custom date or time number property.
 * See `lxlsx_workbook_set_custom_property_string()` above for details.
 *
 * @code
 *     lxlsx_datetime datetime  = {2016, 12, 1,  11, 55, 0.0};
 *
 *     lxlsx_workbook_set_custom_property_datetime(workbook, "Date completed", &datetime);
 * @endcode
 */
lxlsx_error lxlsx_workbook_set_custom_property_datetime(lxlsx_workbook *workbook,
                                                const char *name,
                                                lxlsx_datetime *datetime);

/**
 * @brief Create a defined name in the workbook to use as a variable.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     The defined name.
 * @param formula  The cell or range that the defined name refers to.
 *
 * @return A #lxlsx_error.
 *
 * This function is used to defined a name that can be used to represent a
 * value, a single cell or a range of cells in a workbook: These defined names
 * can then be used in formulas:
 *
 * @code
 *     lxlsx_workbook_define_name(workbook, "Exchange_rate", "=0.96");
 *     lxlsx_worksheet_write_formula(worksheet, 2, 1, "=Exchange_rate", NULL);
 *
 * @endcode
 *
 * @image html defined_name.png
 *
 * As in Excel a name defined like this is "global" to the workbook and can be
 * referred to from any worksheet:
 *
 * @code
 *     // Global workbook name.
 *     lxlsx_workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10");
 * @endcode
 *
 * It is also possible to define a local/worksheet name by prefixing it with
 * the sheet name using the syntax `'sheetname!definedname'`:
 *
 * @code
 *     // Local worksheet name.
 *     lxlsx_workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * If the sheet name contains spaces or special characters you must follow the
 * Excel convention and enclose it in single quotes:
 *
 * @code
 *     lxlsx_workbook_define_name(workbook, "'New Data'!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * The rules for names in Excel are explained in the
 * [Microsoft Office documentation](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64).
 *
 */
lxlsx_error lxlsx_workbook_define_name(lxlsx_workbook *workbook, const char *name,
                               const char *formula);

/**
 * @brief Get the default URL format used with `lxlsx_worksheet_write_url()`.
 *
 * @param  workbook Pointer to a lxlsx_workbook instance.
 * @return A lxlsx_format instance that has hyperlink properties set.
 *
 * This function returns a lxlsx_format instance that is used for the default
 * blue underline hyperlink in the `lxlsx_worksheet_write_url()` function when a
 * format isn't specified:
 *
 * @code
 *     lxlsx_format *url_format = lxlsx_workbook_get_default_url_format(workbook);
 * @endcode
 *
 * The format is the hyperlink style defined by Excel for the default theme.
 * This format is only ever required when overwriting a string URL with
 * data of a different type. See the example below.
 */
lxlsx_format *lxlsx_workbook_get_default_url_format(lxlsx_workbook *workbook);

/**
 * @brief Get a worksheet object from its name.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     Worksheet name.
 *
 * @return A lxlsx_worksheet object.
 *
 * This function returns a lxlsx_worksheet object reference based on its name:
 *
 * @code
 *     worksheet = lxlsx_workbook_get_worksheet_by_name(workbook, "Sheet1");
 * @endcode
 *
 */
lxlsx_worksheet *lxlsx_workbook_get_worksheet_by_name(lxlsx_workbook *workbook,
                                              const char *name);

/**
 * @brief Get a chartsheet object from its name.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     chartsheet name.
 *
 * @return A lxlsx_chartsheet object.
 *
 * This function returns a lxlsx_chartsheet object reference based on its name:
 *
 * @code
 *     chartsheet = lxlsx_workbook_get_chartsheet_by_name(workbook, "Chart1");
 * @endcode
 *
 */
lxlsx_chartsheet *lxlsx_workbook_get_chartsheet_by_name(lxlsx_workbook *workbook,
                                                const char *name);

/**
 * @brief Validate a worksheet or chartsheet name.
 *
 * @param workbook  Pointer to a lxlsx_workbook instance.
 * @param sheetname Sheet name to validate.
 *
 * @return A #lxlsx_error.
 *
 * This function is used to validate a worksheet or chartsheet name according
 * to the rules used by Excel:
 *
 * - The name cannot be blank.
 * - The name is less than or equal to 31 UTF-8 characters.
 * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
 * - The name doesn't start or end with an apostrophe.
 * - The name isn't already in use. (Case insensitive, see the note below).
 *
 * @code
 *     lxlsx_error err = lxlsx_workbook_validate_sheet_name(workbook, "Foglio");
 * @endcode
 *
 * This function is called by `lxlsx_workbook_add_worksheet()` and
 * `lxlsx_workbook_add_chartsheet()` but it can be explicitly called by the user
 * beforehand to ensure that the sheet name is valid.
 *
 * @note You should also avoid using the worksheet name "History" (case
 * insensitive) which is reserved in English language versions of
 * Excel. Non-English versions may have restrictions on the equivalent word.
 *
 * @note This function does an ASCII lowercase string comparison to determine
 * if the sheet name is already in use. It doesn't take UTF-8 characters into
 * account. Thus it would flag "Café" and "café" as a duplicate (just like
 * Excel) but it wouldn't catch "CAFÉ". If you need a full UTF-8 case
 * insensitive check you should use a third party library to implement it.
 *
 */
lxlsx_error lxlsx_workbook_validate_sheet_name(lxlsx_workbook *workbook,
                                       const char *sheetname);

/**
 * @brief Add a vbaProject binary to the Excel workbook.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param filename The path/filename of the vbaProject.bin file.
 *
 * The `%lxlsx_workbook_add_vba_project()` function can be used to add macros or
 * functions to a workbook using a binary VBA project file that has been
 * extracted from an existing Excel xlsm file:
 *
 * @code
 *     lxlsx_workbook_add_vba_project(workbook, "vbaProject.bin");
 * @endcode
 *
 * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
 * have to be `vbaProject.bin`. Any suitable path/name for an existing VBA bin
 * file will do.
 *
 * Once you add a VBA project had been add to an libxlsxwriter workbook you
 * should ensure that the file extension is `.xlsm` to prevent Excel from
 * giving a warning when it opens the file:
 *
 * @code
 *     lxlsx_workbook *workbook = new_workbook("macro.xlsm");
 * @endcode
 *
 * See also @ref working_with_macros
 *
 * @return A #lxlsx_error.
 */
lxlsx_error lxlsx_workbook_add_vba_project(lxlsx_workbook *workbook,
                                   const char *filename);

/**
 * @brief Add a vbaProject binary and a vbaProjectSignature binary to the Excel
 * workbook.
 *
 * @param workbook    Pointer to a lxlsx_workbook instance.
 * @param vba_project The path/filename of the vbaProject.bin file.
 * @param signature   The path/filename of the vbaProjectSignature.bin file.
 *
 * The `%lxlsx_workbook_add_signed_vba_project()` function can be used to add digitally
 * signed macros or functions to a workbook. The function adds a binary VBA project
 * file and a binary VBA project signature file that have been extracted from an
 * existing Excel xlsm file with digitally signed macros:
 *
 * @code
 *     lxlsx_workbook_add_signed_vba_project(workbook, "vbaProject.bin", "vbaProjectSignature.bin");
 * @endcode
 *
 * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
 * have to be `vbaProject.bin`. Any suitable path/name for an existing VBA bin
 * file will do. The same applies for `vbaProjectSignature.bin`.
 *
 * See also @ref working_with_macros
 *
 * @return A #lxlsx_error.
 */
lxlsx_error lxlsx_workbook_add_signed_vba_project(lxlsx_workbook *workbook,
                                          const char *vba_project,
                                          const char *signature);

/**
 * @brief Set the VBA name for the workbook.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param name     Name of the workbook used by VBA.
 *
 * The `lxlsx_workbook_set_vba_name()` function can be used to set the VBA name for
 * the workbook. This is sometimes required when a vbaProject macro included
 * via `lxlsx_workbook_add_vba_project()` refers to the workbook by a name other
 * than `ThisWorkbook`.
 *
 * @code
 *     lxlsx_workbook_set_vba_name(workbook, "MyWorkbook");
 * @endcode
 *
 * If an Excel VBA name for the workbook isn't specified then libxlsxwriter
 * will use `ThisWorkbook`.
 *
 * See also @ref working_with_macros
 *
 * @return A #lxlsx_error.
 */
lxlsx_error lxlsx_workbook_set_vba_name(lxlsx_workbook *workbook, const char *name);

/**
 * @brief Add a recommendation to open the file in "read-only" mode.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 *
 * This function can be used to set the Excel "Read-only Recommended" option
 * that is available when saving a file. This presents the user of the file
 * with an option to open it in "read-only" mode. This means that any changes
 * to the file can't be saved back to the same file and must be saved to a new
 * file. It can be set as follows:
 *
 * @code
 *     lxlsx_workbook_read_only_recommended(workbook);
 * @endcode
 *
 * Which will raise a dialog like the following when opening the file:
 *
 * @image html read_only.png
 */
void lxlsx_workbook_read_only_recommended(lxlsx_workbook *workbook);

/**
 * @brief Set the workbook to use the 1904 epoch.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 *
 * The `%lxlsx_workbook_use_1904_epoch()` function can be used to set the workbook to
 * use the 1904 epoch instead of the default 1900 epoch.
 *
 * Excel supports two date epochs. The first based on 1900-01-01 is the default
 * for all Windows versions of Excel and for recent versions of Excel for macOS.
 * Older versions of Excel for macOS used a 1904-01-01 epoch. The 1904 epoch can
 * be set for compatibility with older versions of Excel or to work around the
 * Excel limitation of not being able to handle negative times.
 *
 * This function should be called before `lxlsx_worksheet_add_worksheet()`.
 *
 * @code
 *     lxlsx_workbook_use_1904_epoch(workbook);
 * @endcode
 *
 */
void lxlsx_workbook_use_1904_epoch(lxlsx_workbook *workbook);

/**
 * @brief Set the size of a workbook window.
 *
 * @param workbook Pointer to a lxlsx_workbook instance.
 * @param width    Width of the window in pixels.
 * @param height   Height of the window in pixels.
 *
 * Set the size of a workbook window. This is generally only useful on macOS
 * since Microsoft Windows uses the window size from the last time an Excel file
 * was opened/saved. The default size is 1073 x 644 pixels.
 *
 * The resulting pixel sizes may not exactly match the target screen and
 * resolution since it is based on the original Excel for Windows sizes. Some
 * trial and error may be required to get an exact size.
 */
void lxlsx_workbook_set_size(lxlsx_workbook *workbook,
                       uint16_t width, uint16_t height);

void lxlsx_workbook_free(lxlsx_workbook *workbook);
void lxlsx_workbook_assemble_xml_file(lxlsx_workbook *workbook);
void lxlsx_workbook_set_default_xf_indices(lxlsx_workbook *workbook);
void lxlsx_workbook_unset_default_url_format(lxlsx_workbook *workbook);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _workbook_xml_declaration(lxlsx_workbook *self);
STATIC void _write_workbook(lxlsx_workbook *self);
STATIC void _write_file_version(lxlsx_workbook *self);
STATIC void _write_workbook_pr(lxlsx_workbook *self);
STATIC void _write_book_views(lxlsx_workbook *self);
STATIC void _write_workbook_view(lxlsx_workbook *self);
STATIC void _write_sheet(lxlsx_workbook *self,
                         const char *name, uint32_t sheet_id, uint8_t hidden);
STATIC void _write_sheets(lxlsx_workbook *self);
STATIC void _write_calc_pr(lxlsx_workbook *self);

STATIC void _write_defined_name(lxlsx_workbook *self,
                                lxlsx_defined_name *define_name);
STATIC void _write_defined_names(lxlsx_workbook *self);

STATIC lxlsx_error _store_defined_name(lxlsx_workbook *self, const char *name,
                                     const char *lxlsx_app_name,
                                     const char *formula, int16_t index,
                                     uint8_t hidden);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_WORKBOOK_H__ */


/* XLSX workbook read API. */

#ifndef LXLSX_READER_WORKBOOK_H
#define LXLSX_READER_WORKBOOK_H

#include "libxlsx/common.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/styles.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_workbook lxlsx_reader_workbook;

typedef struct {
    lxlsx_reader_sst_mode sst_mode;
} lxlsx_reader_open_options;

lxlsx_reader_error lxlsx_reader_workbook_open       (const char *filename, lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_fd    (int fd,                lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_memory(const void *data, size_t len, lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_ex    (const char *filename,
                                   const lxlsx_reader_open_options *opts,
                                   lxlsx_reader_workbook **out);
void      lxlsx_reader_workbook_close      (lxlsx_reader_workbook *wb);

size_t            lxlsx_reader_workbook_sheet_count(const lxlsx_reader_workbook *wb);
const char       *lxlsx_reader_workbook_sheet_name (const lxlsx_reader_workbook *wb, size_t index);

typedef enum {
    LXLSX_READER_SHEET_VISIBLE      = 0,
    LXLSX_READER_SHEET_HIDDEN       = 1,
    LXLSX_READER_SHEET_VERY_HIDDEN  = 2
} lxlsx_reader_sheet_visibility;

lxlsx_reader_sheet_visibility lxlsx_reader_workbook_sheet_visibility(const lxlsx_reader_workbook *wb, size_t index);

/* Defined names (a.k.a. named ranges). scope_sheet_index is -1 for
 * workbook-scope; otherwise the 0-based sheet index it's bound to. */
typedef struct {
    const char *name;
    const char *formula;
    int         scope_sheet_index;
    int         hidden;
} lxlsx_reader_defined_name;

size_t lxlsx_reader_workbook_defined_name_count(const lxlsx_reader_workbook *wb);
int    lxlsx_reader_workbook_defined_name_get  (const lxlsx_reader_workbook *wb, size_t idx,
                                       lxlsx_reader_defined_name *out);

lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_name (lxlsx_reader_workbook   *wb,
                                              const char     *name,
                                              uint32_t        flags,
                                              lxlsx_reader_worksheet **out);
lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_index(lxlsx_reader_workbook   *wb,
                                              size_t          index,
                                              uint32_t        flags,
                                              lxlsx_reader_worksheet **out);

int               lxlsx_reader_workbook_uses_1904_dates(const lxlsx_reader_workbook *wb);
const lxlsx_reader_styles *lxlsx_reader_workbook_get_styles     (const lxlsx_reader_workbook *wb);

#ifdef __cplusplus
}
#endif

#endif
