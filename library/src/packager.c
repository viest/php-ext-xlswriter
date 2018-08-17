/*****************************************************************************
 * packager - A library for creating Excel XLSX packager files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/packager.h"
#include "xlsxwriter/hash_table.h"
#include "xlsxwriter/utility.h"

STATIC lxw_error _add_file_to_zip(lxw_packager *self, FILE * file,
                                  const char *filename);

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/
/* Avoid non MSVC definition of _WIN32 in MinGW. */

#ifdef __MINGW32__
#undef _WIN32
#endif

#ifdef _WIN32

/* Silence Windows warning with duplicate symbol for SLIST_ENTRY in local
 * queue.h and widows.h. */
#undef SLIST_ENTRY

#include <windows.h>

#ifdef USE_SYSTEM_MINIZIP
#include "minizip/iowin32.h"
#else
#include "../third_party/minizip/iowin32.h"
#endif

zipFile
_open_zipfile_win32(const char *filename)
{
    int n;
    zlib_filefunc64_def filefunc;

    wchar_t wide_filename[_MAX_PATH + 1] = L"";

    /* Build a UTF-16 filename for Win32. */
    n = MultiByteToWideChar(CP_UTF8, 0, filename, (int) strlen(filename),
                            wide_filename, _MAX_PATH);

    if (n == 0) {
        LXW_ERROR("MultiByteToWideChar error");
        return NULL;
    }

    /* Use the native Win32 file handling functions with minizip. */
    fill_win32_filefunc64W(&filefunc);

    return zipOpen2_64(wide_filename, 0, NULL, &filefunc);
}

#endif

/*
 * Create a new packager object.
 */
lxw_packager *
lxw_packager_new(const char *filename, char *tmpdir)
{
    lxw_packager *packager = calloc(1, sizeof(lxw_packager));
    GOTO_LABEL_ON_MEM_ERROR(packager, mem_error);

    packager->buffer = calloc(1, LXW_ZIP_BUFFER_SIZE);
    GOTO_LABEL_ON_MEM_ERROR(packager->buffer, mem_error);

    packager->filename = lxw_strdup(filename);
    packager->tmpdir = tmpdir;
    GOTO_LABEL_ON_MEM_ERROR(packager->filename, mem_error);

    packager->buffer_size = LXW_ZIP_BUFFER_SIZE;

    /* Initialize the zip_fileinfo struct to Jan 1 1980 like Excel. */
    packager->zipfile_info.tmz_date.tm_sec = 0;
    packager->zipfile_info.tmz_date.tm_min = 0;
    packager->zipfile_info.tmz_date.tm_hour = 0;
    packager->zipfile_info.tmz_date.tm_mday = 1;
    packager->zipfile_info.tmz_date.tm_mon = 0;
    packager->zipfile_info.tmz_date.tm_year = 1980;
    packager->zipfile_info.dosDate = 0;
    packager->zipfile_info.internal_fa = 0;
    packager->zipfile_info.external_fa = 0;

    /* Create a zip container for the xlsx file. */
#ifdef _WIN32
    packager->zipfile = _open_zipfile_win32(packager->filename);
#else
    packager->zipfile = zipOpen(packager->filename, 0);
#endif

    if (packager->zipfile == NULL)
        goto mem_error;

    return packager;

mem_error:
    lxw_packager_free(packager);
    return NULL;
}

/*
 * Free a packager object.
 */
void
lxw_packager_free(lxw_packager *packager)
{
    if (!packager)
        return;

    free(packager->buffer);
    free(packager->filename);
    free(packager);
}

/*****************************************************************************
 *
 * File assembly functions.
 *
 ****************************************************************************/
/*
 * Write the workbook.xml file.
 */
STATIC lxw_error
_write_workbook_file(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_error err;

    workbook->file = lxw_tmpfile(self->tmpdir);
    if (!workbook->file)
        return LXW_ERROR_CREATING_TMPFILE;

    lxw_workbook_assemble_xml_file(workbook);

    err = _add_file_to_zip(self, workbook->file, "xl/workbook.xml");
    RETURN_ON_ERROR(err);

    fclose(workbook->file);

    return LXW_NO_ERROR;
}

/*
 * Write the worksheet files.
 */
STATIC lxw_error
_write_worksheet_files(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err;

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {
        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/worksheets/sheet%d.xml", index++);

        if (worksheet->optimize_row)
            lxw_worksheet_write_single_row(worksheet);

        worksheet->file = lxw_tmpfile(self->tmpdir);
        if (!worksheet->file)
            return LXW_ERROR_CREATING_TMPFILE;

        lxw_worksheet_assemble_xml_file(worksheet);

        err = _add_file_to_zip(self, worksheet->file, sheetname);
        RETURN_ON_ERROR(err);

        fclose(worksheet->file);
    }

    return LXW_NO_ERROR;
}

/*
 * Write the /xl/media/image?.xml files.
 */
STATIC lxw_error
_write_image_files(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    lxw_image_options *image;
    lxw_error err;
    FILE *image_stream;

    char filename[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {

        if (STAILQ_EMPTY(worksheet->image_data))
            continue;

        STAILQ_FOREACH(image, worksheet->image_data, list_pointers) {

            lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                         "xl/media/image%d.%s", index++, image->extension);

            /* Check that the image file exists and can be opened. */
            image_stream = fopen(image->filename, "rb");
            if (!image_stream) {
                LXW_WARN_FORMAT1("Error adding image to xlsx file: file "
                                 "doesn't exist or can't be opened: %s.",
                                 image->filename);
                return LXW_ERROR_CREATING_TMPFILE;
            }

            err = _add_file_to_zip(self, image_stream, filename);
            fclose(image_stream);

            RETURN_ON_ERROR(err);
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Write the chart files.
 */
STATIC lxw_error
_write_chart_files(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_chart *chart;
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err;

    STAILQ_FOREACH(chart, workbook->ordered_charts, ordered_list_pointers) {

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/charts/chart%d.xml", index++);

        chart->file = lxw_tmpfile(self->tmpdir);
        if (!chart->file)
            return LXW_ERROR_CREATING_TMPFILE;

        lxw_chart_assemble_xml_file(chart);

        err = _add_file_to_zip(self, chart->file, sheetname);
        RETURN_ON_ERROR(err);

        self->chart_count++;

        fclose(chart->file);
    }

    return LXW_NO_ERROR;
}

/*
 * Write the drawing files.
 */
STATIC lxw_error
_write_drawing_files(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    lxw_drawing *drawing;
    char filename[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err;

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {
        drawing = worksheet->drawing;

        if (drawing) {
            lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                         "xl/drawings/drawing%d.xml", index++);

            drawing->file = lxw_tmpfile(self->tmpdir);
            if (!drawing->file)
                return LXW_ERROR_CREATING_TMPFILE;

            lxw_drawing_assemble_xml_file(drawing);
            err = _add_file_to_zip(self, drawing->file, filename);
            RETURN_ON_ERROR(err);

            fclose(drawing->file);

            self->drawing_count++;
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Write the sharedStrings.xml file.
 */
STATIC lxw_error
_write_shared_strings_file(lxw_packager *self)
{
    lxw_sst *sst = self->workbook->sst;
    lxw_error err;

    /* Skip the sharedStrings file if there are no shared strings. */
    if (!sst->string_count)
        return LXW_NO_ERROR;

    sst->file = lxw_tmpfile(self->tmpdir);
    if (!sst->file)
        return LXW_ERROR_CREATING_TMPFILE;

    lxw_sst_assemble_xml_file(sst);

    err = _add_file_to_zip(self, sst->file, "xl/sharedStrings.xml");
    RETURN_ON_ERROR(err);

    fclose(sst->file);

    return LXW_NO_ERROR;
}

/*
 * Write the app.xml file.
 */
STATIC lxw_error
_write_app_file(lxw_packager *self)
{
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    lxw_defined_name *defined_name;
    lxw_app *app;
    uint16_t named_range_count = 0;
    char *autofilter;
    char *has_range;
    char number[LXW_ATTR_32] = { 0 };
    lxw_error err = LXW_NO_ERROR;

    app = lxw_app_new();
    if (!app) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    app->file = lxw_tmpfile(self->tmpdir);
    if (!app->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxw_snprintf(number, LXW_ATTR_32, "%d", self->workbook->num_sheets);

    lxw_app_add_heading_pair(app, "Worksheets", number);

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {
        lxw_app_add_part_name(app, worksheet->name);
    }

    /* Add the Named Ranges parts. */
    TAILQ_FOREACH(defined_name, workbook->defined_names, list_pointers) {

        has_range = strchr(defined_name->formula, '!');
        autofilter = strstr(defined_name->app_name, "_FilterDatabase");

        /* Only store defined names with ranges (except for autofilters). */
        if (has_range && !autofilter) {
            lxw_app_add_part_name(app, defined_name->app_name);
            named_range_count++;
        }
    }

    /* Add the Named Range heading pairs. */
    if (named_range_count) {
        lxw_snprintf(number, LXW_ATTR_32, "%d", named_range_count);
        lxw_app_add_heading_pair(app, "Named Ranges", number);
    }

    /* Set the app/doc properties. */
    app->properties = workbook->properties;

    lxw_app_assemble_xml_file(app);

    err = _add_file_to_zip(self, app->file, "docProps/app.xml");

    fclose(app->file);

mem_error:
    lxw_app_free(app);

    return err;
}

/*
 * Write the core.xml file.
 */
STATIC lxw_error
_write_core_file(lxw_packager *self)
{
    lxw_error err = LXW_NO_ERROR;
    lxw_core *core = lxw_core_new();

    if (!core) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    core->file = lxw_tmpfile(self->tmpdir);
    if (!core->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    core->properties = self->workbook->properties;

    lxw_core_assemble_xml_file(core);

    err = _add_file_to_zip(self, core->file, "docProps/core.xml");

    fclose(core->file);

mem_error:
    lxw_core_free(core);

    return err;
}

/*
 * Write the custom.xml file.
 */
STATIC lxw_error
_write_custom_file(lxw_packager *self)
{
    lxw_custom *custom;
    lxw_error err = LXW_NO_ERROR;

    if (STAILQ_EMPTY(self->workbook->custom_properties))
        return LXW_NO_ERROR;

    custom = lxw_custom_new();
    if (!custom) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    custom->file = lxw_tmpfile(self->tmpdir);
    if (!custom->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    custom->custom_properties = self->workbook->custom_properties;

    lxw_custom_assemble_xml_file(custom);

    err = _add_file_to_zip(self, custom->file, "docProps/custom.xml");

    fclose(custom->file);

mem_error:
    lxw_custom_free(custom);
    return err;
}

/*
 * Write the theme.xml file.
 */
STATIC lxw_error
_write_theme_file(lxw_packager *self)
{
    lxw_error err = LXW_NO_ERROR;
    lxw_theme *theme = lxw_theme_new();

    if (!theme) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    theme->file = lxw_tmpfile(self->tmpdir);
    if (!theme->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxw_theme_assemble_xml_file(theme);

    err = _add_file_to_zip(self, theme->file, "xl/theme/theme1.xml");

    fclose(theme->file);

mem_error:
    lxw_theme_free(theme);

    return err;
}

/*
 * Write the styles.xml file.
 */
STATIC lxw_error
_write_styles_file(lxw_packager *self)
{
    lxw_styles *styles = lxw_styles_new();
    lxw_hash_element *hash_element;
    lxw_error err = LXW_NO_ERROR;

    if (!styles) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    /* Copy the unique and in-use formats from the workbook to the styles
     * xf_format list. */
    LXW_FOREACH_ORDERED(hash_element, self->workbook->used_xf_formats) {
        lxw_format *workbook_format = (lxw_format *) hash_element->value;
        lxw_format *style_format = lxw_format_new();

        if (!style_format) {
            err = LXW_ERROR_MEMORY_MALLOC_FAILED;
            goto mem_error;
        }

        memcpy(style_format, workbook_format, sizeof(lxw_format));
        STAILQ_INSERT_TAIL(styles->xf_formats, style_format, list_pointers);
    }

    styles->font_count = self->workbook->font_count;
    styles->border_count = self->workbook->border_count;
    styles->fill_count = self->workbook->fill_count;
    styles->num_format_count = self->workbook->num_format_count;
    styles->xf_count = self->workbook->used_xf_formats->unique_count;

    styles->file = lxw_tmpfile(self->tmpdir);
    if (!styles->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxw_styles_assemble_xml_file(styles);

    err = _add_file_to_zip(self, styles->file, "xl/styles.xml");

    fclose(styles->file);

mem_error:
    lxw_styles_free(styles);

    return err;
}

/*
 * Write the ContentTypes.xml file.
 */
STATIC lxw_error
_write_content_types_file(lxw_packager *self)
{
    lxw_content_types *content_types = lxw_content_types_new();
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    char filename[LXW_MAX_ATTRIBUTE_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err = LXW_NO_ERROR;

    if (!content_types) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    content_types->file = lxw_tmpfile(self->tmpdir);
    if (!content_types->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    if (workbook->has_png)
        lxw_ct_add_default(content_types, "png", "image/png");

    if (workbook->has_jpeg)
        lxw_ct_add_default(content_types, "jpeg", "image/jpeg");

    if (workbook->has_bmp)
        lxw_ct_add_default(content_types, "bmp", "image/bmp");

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {
        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "/xl/worksheets/sheet%d.xml", index++);
        lxw_ct_add_worksheet_name(content_types, filename);
    }

    for (index = 1; index <= self->chart_count; index++) {
        lxw_snprintf(filename, LXW_FILENAME_LENGTH, "/xl/charts/chart%d.xml",
                     index);
        lxw_ct_add_chart_name(content_types, filename);
    }

    for (index = 1; index <= self->drawing_count; index++) {
        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "/xl/drawings/drawing%d.xml", index);
        lxw_ct_add_drawing_name(content_types, filename);
    }

    if (workbook->sst->string_count)
        lxw_ct_add_shared_strings(content_types);

    if (!STAILQ_EMPTY(self->workbook->custom_properties))
        lxw_ct_add_custom_properties(content_types);

    lxw_content_types_assemble_xml_file(content_types);

    err = _add_file_to_zip(self, content_types->file, "[Content_Types].xml");

    fclose(content_types->file);

mem_error:
    lxw_content_types_free(content_types);

    return err;
}

/*
 * Write the workbook .rels xml file.
 */
STATIC lxw_error
_write_workbook_rels_file(lxw_packager *self)
{
    lxw_relationships *rels = lxw_relationships_new();
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err = LXW_NO_ERROR;

    if (!rels) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rels->file = lxw_tmpfile(self->tmpdir);
    if (!rels->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {
        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH, "worksheets/sheet%d.xml",
                     index++);
        lxw_add_document_relationship(rels, "/worksheet", sheetname);
    }

    lxw_add_document_relationship(rels, "/theme", "theme/theme1.xml");
    lxw_add_document_relationship(rels, "/styles", "styles.xml");

    if (workbook->sst->string_count)
        lxw_add_document_relationship(rels, "/sharedStrings",
                                      "sharedStrings.xml");

    lxw_relationships_assemble_xml_file(rels);

    err = _add_file_to_zip(self, rels->file, "xl/_rels/workbook.xml.rels");

    fclose(rels->file);

mem_error:
    lxw_free_relationships(rels);

    return err;
}

/*
 * Write the worksheet .rels files for worksheets that contain links to
 * external data such as hyperlinks or drawings.
 */
STATIC lxw_error
_write_worksheet_rels_file(lxw_packager *self)
{
    lxw_relationships *rels;
    lxw_rel_tuple *rel;
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 0;
    lxw_error err;

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {

        index++;

        if (STAILQ_EMPTY(worksheet->external_hyperlinks) &&
            STAILQ_EMPTY(worksheet->external_drawing_links))
            continue;

        rels = lxw_relationships_new();

        rels->file = lxw_tmpfile(self->tmpdir);
        if (!rels->file) {
            lxw_free_relationships(rels);
            return LXW_ERROR_CREATING_TMPFILE;
        }

        STAILQ_FOREACH(rel, worksheet->external_hyperlinks, list_pointers) {
            lxw_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        STAILQ_FOREACH(rel, worksheet->external_drawing_links, list_pointers) {
            lxw_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/worksheets/_rels/sheet%d.xml.rels", index);

        lxw_relationships_assemble_xml_file(rels);

        err = _add_file_to_zip(self, rels->file, sheetname);

        fclose(rels->file);
        lxw_free_relationships(rels);

        RETURN_ON_ERROR(err);
    }

    return LXW_NO_ERROR;
}

/*
 * Write the drawing .rels files for worksheets that contain charts or
 * drawings.
 */
STATIC lxw_error
_write_drawing_rels_file(lxw_packager *self)
{
    lxw_relationships *rels;
    lxw_rel_tuple *rel;
    lxw_workbook *workbook = self->workbook;
    lxw_worksheet *worksheet;
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    lxw_error err;

    STAILQ_FOREACH(worksheet, workbook->worksheets, list_pointers) {

        if (STAILQ_EMPTY(worksheet->drawing_links))
            continue;

        rels = lxw_relationships_new();

        rels->file = lxw_tmpfile(self->tmpdir);
        if (!rels->file) {
            lxw_free_relationships(rels);
            return LXW_ERROR_CREATING_TMPFILE;
        }

        STAILQ_FOREACH(rel, worksheet->drawing_links, list_pointers) {
            lxw_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        }

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/drawings/_rels/drawing%d.xml.rels", index++);

        lxw_relationships_assemble_xml_file(rels);

        err = _add_file_to_zip(self, rels->file, sheetname);

        fclose(rels->file);
        lxw_free_relationships(rels);

        RETURN_ON_ERROR(err);
    }

    return LXW_NO_ERROR;
}

/*
 * Write the _rels/.rels xml file.
 */
STATIC lxw_error
_write_root_rels_file(lxw_packager *self)
{
    lxw_relationships *rels = lxw_relationships_new();
    lxw_error err = LXW_NO_ERROR;

    if (!rels) {
        err = LXW_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rels->file = lxw_tmpfile(self->tmpdir);
    if (!rels->file) {
        err = LXW_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxw_add_document_relationship(rels, "/officeDocument", "xl/workbook.xml");

    lxw_add_package_relationship(rels,
                                 "/metadata/core-properties",
                                 "docProps/core.xml");

    lxw_add_document_relationship(rels,
                                  "/extended-properties", "docProps/app.xml");

    if (!STAILQ_EMPTY(self->workbook->custom_properties))
        lxw_add_document_relationship(rels,
                                      "/custom-properties",
                                      "docProps/custom.xml");

    lxw_relationships_assemble_xml_file(rels);

    err = _add_file_to_zip(self, rels->file, "_rels/.rels");

    fclose(rels->file);

mem_error:
    lxw_free_relationships(rels);

    return err;
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

STATIC lxw_error
_add_file_to_zip(lxw_packager *self, FILE * file, const char *filename)
{
    int16_t error = ZIP_OK;
    size_t size_read;

    error = zipOpenNewFileInZip4_64(self->zipfile,
                                    filename,
                                    &self->zipfile_info,
                                    NULL, 0, NULL, 0, NULL,
                                    Z_DEFLATED, Z_DEFAULT_COMPRESSION, 0,
                                    -MAX_WBITS, DEF_MEM_LEVEL,
                                    Z_DEFAULT_STRATEGY, NULL, 0, 0, 0, 0);

    if (error != ZIP_OK) {
        LXW_ERROR("Error adding member to zipfile");
        RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
    }

    fflush(file);
    rewind(file);

    size_read = fread(self->buffer, 1, self->buffer_size, file);

    while (size_read) {

        if (size_read < self->buffer_size) {
            if (feof(file) == 0) {
                LXW_ERROR("Error reading member file data");
                RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
            }
        }

        error = zipWriteInFileInZip(self->zipfile,
                                    self->buffer, (unsigned int) size_read);

        if (error < 0) {
            LXW_ERROR("Error in writing member in the zipfile");
            RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
        }

        size_read = fread(self->buffer, 1, self->buffer_size, file);
    }

    if (error < 0) {
        RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
    }
    else {
        error = zipCloseFileInZip(self->zipfile);
        if (error != ZIP_OK) {
            LXW_ERROR("Error in closing member in the zipfile");
            RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Write the xml files that make up the XLXS OPC package.
 */
lxw_error
lxw_create_package(lxw_packager *self)
{
    lxw_error error;
    int8_t zip_error;

    error = _write_worksheet_files(self);
    RETURN_ON_ERROR(error);

    error = _write_workbook_file(self);
    RETURN_ON_ERROR(error);

    error = _write_chart_files(self);
    RETURN_ON_ERROR(error);

    error = _write_drawing_files(self);
    RETURN_ON_ERROR(error);

    error = _write_shared_strings_file(self);
    RETURN_ON_ERROR(error);

    error = _write_app_file(self);
    RETURN_ON_ERROR(error);

    error = _write_core_file(self);
    RETURN_ON_ERROR(error);

    error = _write_custom_file(self);
    RETURN_ON_ERROR(error);

    error = _write_theme_file(self);
    RETURN_ON_ERROR(error);

    error = _write_styles_file(self);
    RETURN_ON_ERROR(error);

    error = _write_content_types_file(self);
    RETURN_ON_ERROR(error);

    error = _write_workbook_rels_file(self);
    RETURN_ON_ERROR(error);

    error = _write_worksheet_rels_file(self);
    RETURN_ON_ERROR(error);

    error = _write_drawing_rels_file(self);
    RETURN_ON_ERROR(error);

    error = _write_image_files(self);
    RETURN_ON_ERROR(error);

    error = _write_root_rels_file(self);
    RETURN_ON_ERROR(error);

    zip_error = zipClose(self->zipfile, NULL);
    if (zip_error) {
        RETURN_ON_ZIP_ERROR(zip_error, LXW_ERROR_ZIP_CLOSE);
    }

    return LXW_NO_ERROR;
}
