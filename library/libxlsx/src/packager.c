/*****************************************************************************
 * packager - A library for assembling xml files into an Excel XLSX file.
 *
 * A class for writing the Excel XLSX Packager file.
 *
 * This module is used in conjunction with libxlsxwriter to create an
 * Excel XLSX container file.
 *
 * From Wikipedia: The Open Packaging Conventions (OPC) is a
 * container-file technology initially created by Microsoft to store
 * a combination of XML and non-XML files that together form a single
 * entity such as an Open XML Paper Specification (OpenXPS)
 * document. http://en.wikipedia.org/wiki/Open_Packaging_Conventions.
 *
 * At its simplest an Excel XLSX file contains the following elements::
 *
 *      ____ [Content_Types].xml
 *     |
 *     |____ docProps
 *     | |____ app.xml
 *     | |____ core.xml
 *     |
 *     |____ xl
 *     | |____ workbook.xml
 *     | |____ worksheets
 *     | | |____ sheet1.xml
 *     | |
 *     | |____ styles.xml
 *     | |
 *     | |____ theme
 *     | | |____ theme1.xml
 *     | |
 *     | |_____rels
 *     |   |____ workbook.xml.rels
 *     |
 *     |_____rels
 *       |____ .rels
 *
 * The Packager class coordinates the classes that represent the
 * elements of the package and writes them into the XLSX file.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include <zlib.h>
#include "libxlsx/xmlwriter.h"
#include "libxlsx/packager.h"
#include "libxlsx/hash_table.h"
#include "libxlsx/utility.h"

STATIC lxlsx_error _add_file_to_zip(lxlsx_packager *self, FILE *file,
                                  const char *filename);

STATIC lxlsx_error _add_buffer_to_zip(lxlsx_packager *self, const char *buffer,
                                    size_t buffer_size, const char *filename);

STATIC lxlsx_error _add_to_zip(lxlsx_packager *self, FILE *file,
                             char **buffer, size_t *buffer_size,
                             const char *filename);

STATIC lxlsx_error _write_vml_drawing_rels_file(lxlsx_packager *self,
                                              lxlsx_worksheet *worksheet,
                                              uint32_t index);

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
 * queue.h and windows.h. */
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
        LXLSX_ERROR("MultiByteToWideChar error");
        return NULL;
    }

    /* Use the native Win32 file handling functions with minizip. */
    fill_win32_filefunc64W(&filefunc);

    return zipOpen2_64(wide_filename, 0, NULL, &filefunc);
}

#endif

STATIC voidpf ZCALLBACK
_fopen_memstream(voidpf opaque, const char *filename, int mode)
{
    lxlsx_packager *packager = (lxlsx_packager *) opaque;
    (void) filename;
    (void) mode;
    return lxlsx_get_filehandle(&packager->output_buffer,
                              &packager->output_buffer_size,
                              packager->tmpdir);
}

STATIC int ZCALLBACK
_fclose_memstream(voidpf opaque, voidpf stream)
{
    lxlsx_packager *packager = (lxlsx_packager *) opaque;
    FILE *file = (FILE *) stream;
    long size;

    /* Ensure memstream buffer is updated */
    if (fflush(file))
        goto mem_error;

    /* If the memstream is backed by a temporary file, no buffer is created,
       so create it manually. */
    if (!packager->output_buffer) {
        if (fseek(file, 0L, SEEK_END))
            goto mem_error;

        size = ftell(file);
        if (size == -1)
            goto mem_error;

        packager->output_buffer = malloc(size);
        GOTO_LABEL_ON_MEM_ERROR(packager->output_buffer, mem_error);

        rewind(file);
        if (fread((void *) packager->output_buffer, size, 1, file) < 1)
            goto mem_error;

        packager->output_buffer_size = size;
    }

    return fclose(file);

mem_error:
    fclose(file);
    return EOF;
}

/*
 * Create a new packager object.
 */
lxlsx_packager *
lxlsx_packager_new(const char *filename, const char *tmpdir, uint8_t use_zip64)
{
    zlib_filefunc_def filefunc;
    lxlsx_packager *packager = calloc(1, sizeof(lxlsx_packager));
    GOTO_LABEL_ON_MEM_ERROR(packager, mem_error);

    packager->buffer = calloc(1, LXLSX_ZIP_BUFFER_SIZE);
    GOTO_LABEL_ON_MEM_ERROR(packager->buffer, mem_error);

    packager->filename = NULL;
    packager->tmpdir = tmpdir;

    if (filename) {
        packager->filename = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(packager->filename, mem_error);
    }

    packager->buffer_size = LXLSX_ZIP_BUFFER_SIZE;
    packager->use_zip64 = use_zip64;

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

    packager->output_buffer = NULL;
    packager->output_buffer_size = 0;

    /* Create a zip container for the xlsx file. */
    if (packager->filename) {
#ifdef _WIN32
        packager->zipfile = _open_zipfile_win32(packager->filename);
#else
        packager->zipfile = zipOpen(packager->filename, 0);
#endif
    }
    else {
        fill_fopen_filefunc(&filefunc);
        filefunc.opaque = packager;
        filefunc.zopen_file = _fopen_memstream;
        filefunc.zclose_file = _fclose_memstream;
        packager->zipfile = zipOpen2(packager->filename, 0, NULL, &filefunc);
    }

    if (packager->zipfile == NULL)
        goto mem_error;

    return packager;

mem_error:
    lxlsx_packager_free(packager);
    return NULL;
}

/*
 * Free a packager object.
 */
void
lxlsx_packager_free(lxlsx_packager *packager)
{
    if (!packager)
        return;

    free((void *) packager->buffer);
    free((void *) packager->filename);
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
STATIC lxlsx_error
_write_workbook_file(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_error err;

    char *buffer = NULL;
    size_t buffer_size = 0;
    workbook->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!workbook->file)
        return LXLSX_ERROR_CREATING_TMPFILE;

    lxlsx_workbook_assemble_xml_file(workbook);

    err = _add_to_zip(self, workbook->file, &buffer, &buffer_size,
                      "xl/workbook.xml");
    fclose(workbook->file);
    free(buffer);
    RETURN_ON_ERROR(err);

    return LXLSX_NO_ERROR;
}

/*
 * Write the worksheet files.
 */
STATIC lxlsx_error
_write_worksheet_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/worksheets/sheet%d.xml", index++);

        if (worksheet->optimize_row)
            lxlsx_worksheet_write_single_row(worksheet);

        worksheet->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                             self->tmpdir);
        if (!worksheet->file)
            return LXLSX_ERROR_CREATING_TMPFILE;

        lxlsx_worksheet_assemble_xml_file(worksheet);

        err = _add_to_zip(self, worksheet->file, &buffer, &buffer_size,
                          sheetname);
        fclose(worksheet->file);
        free(buffer);
        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the chartsheet files.
 */
STATIC lxlsx_error
_write_chartsheet_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_chartsheet *chartsheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            chartsheet = sheet->u.chartsheet;
        else
            continue;

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/chartsheets/sheet%d.xml", index++);

        chartsheet->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                              self->tmpdir);
        if (!chartsheet->file)
            return LXLSX_ERROR_CREATING_TMPFILE;

        lxlsx_chartsheet_assemble_xml_file(chartsheet);

        err = _add_to_zip(self, chartsheet->file, &buffer, &buffer_size,
                          sheetname);
        fclose(chartsheet->file);
        free(buffer);
        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the /xl/media/image?.xml files.
 */
STATIC lxlsx_error
_write_image_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_object_properties *object_props;
    lxlsx_error err;
    FILE *image_stream;

    char filename[LXLSX_FILENAME_LENGTH] = { 0 };
    uint32_t index = 1;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (STAILQ_EMPTY(worksheet->image_props)
            && STAILQ_EMPTY(worksheet->embedded_image_props))
            continue;

        STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
                       list_pointers) {

            if (object_props->is_duplicate)
                continue;

            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/media/image%d.%s", index++,
                         object_props->extension);

            if (!object_props->is_image_buffer) {
                /* Check that the image file exists and can be opened. */
                image_stream = lxlsx_fopen(object_props->filename, "rb");
                if (!image_stream) {
                    LXLSX_WARN_FORMAT1("Error adding image to xlsx file: file "
                                     "doesn't exist or can't be opened: %s.",
                                     object_props->filename);
                    return LXLSX_ERROR_CREATING_TMPFILE;
                }

                err = _add_file_to_zip(self, image_stream, filename);
                fclose(image_stream);
            }
            else {
                err = _add_buffer_to_zip(self,
                                         object_props->image_buffer,
                                         object_props->image_buffer_size,
                                         filename);
            }

            RETURN_ON_ERROR(err);
        }

        STAILQ_FOREACH(object_props, worksheet->image_props, list_pointers) {

            if (object_props->is_duplicate)
                continue;

            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/media/image%d.%s", index++,
                         object_props->extension);

            if (!object_props->is_image_buffer) {
                /* Check that the image file exists and can be opened. */
                image_stream = lxlsx_fopen(object_props->filename, "rb");
                if (!image_stream) {
                    LXLSX_WARN_FORMAT1("Error adding image to xlsx file: file "
                                     "doesn't exist or can't be opened: %s.",
                                     object_props->filename);
                    return LXLSX_ERROR_CREATING_TMPFILE;
                }

                err = _add_file_to_zip(self, image_stream, filename);
                fclose(image_stream);
            }
            else {
                err = _add_buffer_to_zip(self,
                                         object_props->image_buffer,
                                         object_props->image_buffer_size,
                                         filename);
            }

            RETURN_ON_ERROR(err);
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the xl/vbaProject.bin file.
 */
STATIC lxlsx_error
_add_vba_project(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_error err;
    FILE *image_stream;

    if (!workbook->vba_project)
        return LXLSX_NO_ERROR;

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(workbook->vba_project, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("Error adding vbaProject.bin to xlsx file: "
                         "file doesn't exist or can't be opened: %s.",
                         workbook->vba_project);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    err = _add_file_to_zip(self, image_stream, "xl/vbaProject.bin");
    fclose(image_stream);
    RETURN_ON_ERROR(err);

    return LXLSX_NO_ERROR;
}

/*
 * Write the xl/vbaProjectSignature.bin file.
 */
STATIC lxlsx_error
_add_vba_project_signature(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_error err;
    FILE *image_stream;

    if (!workbook->vba_project_signature)
        return LXLSX_NO_ERROR;

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(workbook->vba_project_signature, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("Error adding vbaProjectSignature.bin to xlsx file: "
                         "file doesn't exist or can't be opened: %s.",
                         workbook->vba_project_signature);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    err = _add_file_to_zip(self, image_stream, "xl/vbaProjectSignature.bin");
    fclose(image_stream);
    RETURN_ON_ERROR(err);

    return LXLSX_NO_ERROR;
}

/*
 * Write the chart files.
 */
STATIC lxlsx_error
_write_chart_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_chart *chart;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(chart, workbook->ordered_charts, ordered_list_pointers) {

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/charts/chart%d.xml", index++);

        chart->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
        if (!chart->file)
            return LXLSX_ERROR_CREATING_TMPFILE;

        lxlsx_chart_assemble_xml_file(chart);

        err = _add_to_zip(self, chart->file, &buffer, &buffer_size,
                          sheetname);
        fclose(chart->file);
        free(buffer);
        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Count the chart files.
 */
uint32_t
_get_chart_count(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_chart *chart;
    uint32_t lxlsx_chart_count = 0;

    STAILQ_FOREACH(chart, workbook->ordered_charts, ordered_list_pointers) {
        lxlsx_chart_count++;
    }

    return lxlsx_chart_count;
}

/*
 * Write the drawing files.
 */
STATIC lxlsx_error
_write_drawing_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_drawing *drawing;
    char filename[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            worksheet = sheet->u.worksheet;

        drawing = worksheet->drawing;

        if (drawing) {
            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/drawings/drawing%d.xml", index++);

            drawing->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                               self->tmpdir);
            if (!drawing->file)
                return LXLSX_ERROR_CREATING_TMPFILE;

            lxlsx_drawing_assemble_xml_file(drawing);

            err = _add_to_zip(self, drawing->file, &buffer, &buffer_size,
                              filename);
            fclose(drawing->file);
            free(buffer);
            RETURN_ON_ERROR(err);
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Count  the drawing files.
 */
uint32_t
_get_drawing_count(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_drawing *drawing;
    uint32_t lxlsx_drawing_count = 0;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            worksheet = sheet->u.worksheet;

        drawing = worksheet->drawing;

        if (drawing)
            lxlsx_drawing_count++;
    }

    return lxlsx_drawing_count;
}

/*
 * Write the worksheet table files.
 */
STATIC lxlsx_error
_write_table_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_table *table;
    lxlsx_table_obj *lxlsx_table_obj;
    lxlsx_error err;

    char filename[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (STAILQ_EMPTY(worksheet->lxlsx_table_objs))
            continue;

        STAILQ_FOREACH(lxlsx_table_obj, worksheet->lxlsx_table_objs, list_pointers) {

            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/tables/table%d.xml", index++);

            table = lxlsx_table_new();
            if (!table) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                RETURN_ON_ERROR(err);
            }

            table->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                             self->tmpdir);
            if (!table->file) {
                lxlsx_table_free(table);
                return LXLSX_ERROR_CREATING_TMPFILE;
            }

            table->lxlsx_table_obj = lxlsx_table_obj;

            lxlsx_table_assemble_xml_file(table);

            err = _add_to_zip(self, table->file, &buffer, &buffer_size,
                              filename);
            fclose(table->file);
            free(buffer);
            lxlsx_table_free(table);
            RETURN_ON_ERROR(err);
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Count  the table files.
 */
uint32_t
_get_table_count(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    uint32_t lxlsx_table_count = 0;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            worksheet = sheet->u.worksheet;

        lxlsx_table_count += worksheet->lxlsx_table_count;
    }

    return lxlsx_table_count;
}

/*
 * Write the comment/header VML files.
 */
STATIC lxlsx_error
_write_vml_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_vml *vml;
    char filename[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (!worksheet->has_vml && !worksheet->has_header_vml)
            continue;

        if (worksheet->has_vml) {

            vml = lxlsx_vml_new();
            if (!vml)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/drawings/vmlDrawing%d.vml", index++);

            vml->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                           self->tmpdir);
            if (!vml->file) {
                lxlsx_vml_free(vml);
                return LXLSX_ERROR_CREATING_TMPFILE;
            }

            vml->comment_objs = worksheet->comment_objs;
            vml->button_objs = worksheet->button_objs;
            vml->lxlsx_vml_shape_id = worksheet->lxlsx_vml_shape_id;
            vml->comment_display_default = worksheet->comment_display_default;

            if (worksheet->lxlsx_vml_data_id_str) {
                vml->lxlsx_vml_data_id_str = worksheet->lxlsx_vml_data_id_str;
            }
            else {
                fclose(vml->file);
                free(buffer);
                lxlsx_vml_free(vml);
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            }

            lxlsx_vml_assemble_xml_file(vml);

            err = _add_to_zip(self, vml->file, &buffer, &buffer_size,
                              filename);

            fclose(vml->file);
            free(buffer);
            lxlsx_vml_free(vml);

            RETURN_ON_ERROR(err);
        }

        if (worksheet->has_header_vml) {

            err = _write_vml_drawing_rels_file(self, worksheet, index);
            RETURN_ON_ERROR(err);

            vml = lxlsx_vml_new();
            if (!vml)
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "xl/drawings/vmlDrawing%d.vml", index++);

            vml->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                           self->tmpdir);
            if (!vml->file) {
                lxlsx_vml_free(vml);
                return LXLSX_ERROR_CREATING_TMPFILE;
            }

            vml->image_objs = worksheet->header_image_objs;
            vml->lxlsx_vml_shape_id = worksheet->lxlsx_vml_header_id * 1024;

            if (worksheet->lxlsx_vml_header_id_str) {
                vml->lxlsx_vml_data_id_str = worksheet->lxlsx_vml_header_id_str;
            }
            else {
                fclose(vml->file);
                free(buffer);
                lxlsx_vml_free(vml);
                return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            }

            lxlsx_vml_assemble_xml_file(vml);

            err = _add_to_zip(self, vml->file, &buffer, &buffer_size,
                              filename);

            fclose(vml->file);
            free(buffer);
            lxlsx_vml_free(vml);

            RETURN_ON_ERROR(err);
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the comment files.
 */
STATIC lxlsx_error
_write_comment_files(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_comment *comment;
    char filename[LXLSX_FILENAME_LENGTH] = { 0 };
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (!worksheet->has_comments)
            continue;

        comment = lxlsx_comment_new();
        if (!comment)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "xl/comments%d.xml", index++);

        comment->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                           self->tmpdir);
        if (!comment->file) {
            lxlsx_comment_free(comment);
            return LXLSX_ERROR_CREATING_TMPFILE;
        }

        comment->comment_objs = worksheet->comment_objs;
        comment->comment_author = worksheet->comment_author;

        lxlsx_comment_assemble_xml_file(comment);

        err = _add_to_zip(self, comment->file, &buffer, &buffer_size,
                          filename);

        fclose(comment->file);
        free(buffer);
        lxlsx_comment_free(comment);

        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the sharedStrings.xml file.
 */
STATIC lxlsx_error
_write_shared_strings_file(lxlsx_packager *self)
{
    lxlsx_sst *sst = self->workbook->sst;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_error err;

    /* Skip the sharedStrings file if there are no shared strings. */
    if (!sst->string_count)
        return LXLSX_NO_ERROR;

    sst->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!sst->file)
        return LXLSX_ERROR_CREATING_TMPFILE;

    lxlsx_sst_assemble_xml_file(sst);

    err = _add_to_zip(self, sst->file, &buffer, &buffer_size,
                      "xl/sharedStrings.xml");
    fclose(sst->file);
    free(buffer);
    RETURN_ON_ERROR(err);

    return LXLSX_NO_ERROR;
}

/*
 * Write the app.xml file.
 */
STATIC lxlsx_error
_write_app_file(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_chartsheet *chartsheet;
    lxlsx_defined_name *defined_name;
    lxlsx_app *app;
    char *buffer = NULL;
    size_t buffer_size = 0;
    uint32_t named_range_count = 0;
    char *autofilter;
    char *has_range;
    char number[LXLSX_ATTR_32] = { 0 };
    lxlsx_error err = LXLSX_NO_ERROR;

    app = lxlsx_app_new();
    if (!app) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    app->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!app->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    if (self->workbook->num_worksheets) {
        lxlsx_snprintf(number, LXLSX_ATTR_32, "%d",
                     self->workbook->num_worksheets);
        lxlsx_app_add_heading_pair(app, "Worksheets", number);
    }

    if (self->workbook->num_chartsheets) {
        lxlsx_snprintf(number, LXLSX_ATTR_32, "%d",
                     self->workbook->num_chartsheets);
        lxlsx_app_add_heading_pair(app, "Charts", number);
    }

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (!sheet->is_chartsheet) {
            worksheet = sheet->u.worksheet;
            lxlsx_app_add_part_name(app, worksheet->name);
        }
    }

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet) {
            chartsheet = sheet->u.chartsheet;
            lxlsx_app_add_part_name(app, chartsheet->name);
        }
    }

    /* Add the Named Ranges parts. */
    TAILQ_FOREACH(defined_name, workbook->defined_names, list_pointers) {

        has_range = strchr(defined_name->formula, '!');
        autofilter = strstr(defined_name->lxlsx_app_name, "_FilterDatabase");

        /* Only store defined names with ranges (except for autofilters). */
        if (has_range && !autofilter) {
            lxlsx_app_add_part_name(app, defined_name->lxlsx_app_name);
            named_range_count++;
        }
    }

    /* Add the Named Range heading pairs. */
    if (named_range_count) {
        lxlsx_snprintf(number, LXLSX_ATTR_32, "%d", named_range_count);
        lxlsx_app_add_heading_pair(app, "Named Ranges", number);
    }

    /* Set the app/doc properties. */
    app->properties = workbook->properties;

    app->doc_security = workbook->read_only;

    lxlsx_app_assemble_xml_file(app);

    err = _add_to_zip(self, app->file, &buffer, &buffer_size,
                      "docProps/app.xml");

    fclose(app->file);
    free(buffer);

mem_error:
    lxlsx_app_free(app);

    return err;
}

/*
 * Write the core.xml file.
 */
STATIC lxlsx_error
_write_core_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_core *core = lxlsx_core_new();
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!core) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    core->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!core->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    core->properties = self->workbook->properties;

    lxlsx_core_assemble_xml_file(core);

    err = _add_to_zip(self, core->file, &buffer, &buffer_size,
                      "docProps/core.xml");

    fclose(core->file);
    free(buffer);

mem_error:
    lxlsx_core_free(core);

    return err;
}

/*
 * Write the metadata.xml file.
 */
STATIC lxlsx_error
_write_metadata_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_metadata *metadata;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!self->workbook->has_metadata)
        return LXLSX_NO_ERROR;

    metadata = lxlsx_metadata_new();

    if (!metadata) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    metadata->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!metadata->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    metadata->has_embedded_images = self->workbook->has_embedded_images;
    metadata->num_embedded_images = self->workbook->num_embedded_images;
    metadata->has_dynamic_functions = self->workbook->has_dynamic_functions;

    lxlsx_metadata_assemble_xml_file(metadata);

    err = _add_to_zip(self, metadata->file, &buffer, &buffer_size,
                      "xl/metadata.xml");

    fclose(metadata->file);
    free(buffer);

mem_error:
    lxlsx_metadata_free(metadata);

    return err;
}

/*
 * Write the rdrichvalue.xml file.
 */
STATIC lxlsx_error
_write_rich_value_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_rich_value *rich_value;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!self->workbook->has_embedded_images)
        return LXLSX_NO_ERROR;

    rich_value = lxlsx_rich_value_new();
    if (!rich_value) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rich_value->workbook = self->workbook;

    rich_value->file =
        lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rich_value->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_rich_value_assemble_xml_file(rich_value);

    err = _add_to_zip(self, rich_value->file, &buffer, &buffer_size,
                      "xl/richData/rdrichvalue.xml");

    fclose(rich_value->file);
    free(buffer);

mem_error:
    lxlsx_rich_value_free(rich_value);

    return err;
}

/*
 * Write the richValueRel.xml file.
 */
STATIC lxlsx_error
_write_rich_value_rel_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_rich_value_rel *lxlsx_rich_value_rel;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!self->workbook->has_embedded_images)
        return LXLSX_NO_ERROR;

    lxlsx_rich_value_rel = lxlsx_rich_value_rel_new();
    if (!lxlsx_rich_value_rel) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    lxlsx_rich_value_rel->num_embedded_images = self->workbook->num_embedded_images;

    lxlsx_rich_value_rel->file =
        lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!lxlsx_rich_value_rel->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_rich_value_rel_assemble_xml_file(lxlsx_rich_value_rel);

    err = _add_to_zip(self, lxlsx_rich_value_rel->file, &buffer, &buffer_size,
                      "xl/richData/richValueRel.xml");

    fclose(lxlsx_rich_value_rel->file);
    free(buffer);

mem_error:
    lxlsx_rich_value_rel_free(lxlsx_rich_value_rel);

    return err;
}

/*
 * Write the rdRichValueTypes.xml file.
 */
STATIC lxlsx_error
_write_rich_value_types_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_rich_value_types *lxlsx_rich_value_types;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!self->workbook->has_embedded_images)
        return LXLSX_NO_ERROR;

    lxlsx_rich_value_types = lxlsx_rich_value_types_new();
    if (!lxlsx_rich_value_types) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    lxlsx_rich_value_types->file =
        lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!lxlsx_rich_value_types->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_rich_value_types_assemble_xml_file(lxlsx_rich_value_types);

    err = _add_to_zip(self, lxlsx_rich_value_types->file, &buffer, &buffer_size,
                      "xl/richData/rdRichValueTypes.xml");

    fclose(lxlsx_rich_value_types->file);
    free(buffer);

mem_error:
    lxlsx_rich_value_types_free(lxlsx_rich_value_types);

    return err;
}

/*
 * Write the rdrichvaluestructure.xml file.
 */
STATIC lxlsx_error
_write_rich_value_structure_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_rich_value_structure *lxlsx_rich_value_structure;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!self->workbook->has_embedded_images)
        return LXLSX_NO_ERROR;

    lxlsx_rich_value_structure = lxlsx_rich_value_structure_new();
    if (!lxlsx_rich_value_structure) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    lxlsx_rich_value_structure->has_embedded_image_descriptions =
        self->workbook->has_embedded_image_descriptions;

    lxlsx_rich_value_structure->file =
        lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!lxlsx_rich_value_structure->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_rich_value_structure_assemble_xml_file(lxlsx_rich_value_structure);

    err = _add_to_zip(self, lxlsx_rich_value_structure->file, &buffer, &buffer_size,
                      "xl/richData/rdrichvaluestructure.xml");

    fclose(lxlsx_rich_value_structure->file);
    free(buffer);

mem_error:
    lxlsx_rich_value_structure_free(lxlsx_rich_value_structure);

    return err;
}

/*
 * Write the custom.xml file.
 */
STATIC lxlsx_error
_write_custom_file(lxlsx_packager *self)
{
    lxlsx_custom *custom;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (STAILQ_EMPTY(self->workbook->lxlsx_custom_properties))
        return LXLSX_NO_ERROR;

    custom = lxlsx_custom_new();
    if (!custom) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    custom->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!custom->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    custom->lxlsx_custom_properties = self->workbook->lxlsx_custom_properties;

    lxlsx_custom_assemble_xml_file(custom);

    err = _add_to_zip(self, custom->file, &buffer, &buffer_size,
                      "docProps/custom.xml");

    fclose(custom->file);
    free(buffer);

mem_error:
    lxlsx_custom_free(custom);
    return err;
}

/*
 * Write the theme.xml file.
 */
STATIC lxlsx_error
_write_theme_file(lxlsx_packager *self)
{
    lxlsx_error err = LXLSX_NO_ERROR;
    lxlsx_theme *theme = lxlsx_theme_new();
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!theme) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    theme->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!theme->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_theme_assemble_xml_file(theme);

    err = _add_to_zip(self, theme->file, &buffer, &buffer_size,
                      "xl/theme/theme1.xml");

    fclose(theme->file);
    free(buffer);

mem_error:
    lxlsx_theme_free(theme);

    return err;
}

/*
 * Write the styles.xml file.
 */
STATIC lxlsx_error
_write_styles_file(lxlsx_packager *self)
{
    lxlsx_styles *styles = lxlsx_styles_new();
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_hash_element *hash_element;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (!styles) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    /* Copy the unique and in-use formats from the workbook to the styles
     * xf_format list. */
    LXLSX_FOREACH_ORDERED(hash_element, self->workbook->used_xf_formats) {
        lxlsx_format *lxlsx_workbook_format = (lxlsx_format *) hash_element->value;
        lxlsx_format *style_format = lxlsx_format_new();

        if (!style_format) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto mem_error;
        }

        memcpy(style_format, lxlsx_workbook_format, sizeof(lxlsx_format));
        STAILQ_INSERT_TAIL(styles->xf_formats, style_format, list_pointers);
    }

    /* Copy the unique and in-use dxf formats from the workbook to the styles
     * dxf_format list. */
    LXLSX_FOREACH_ORDERED(hash_element, self->workbook->used_dxf_formats) {
        lxlsx_format *lxlsx_workbook_format = (lxlsx_format *) hash_element->value;
        lxlsx_format *style_format = lxlsx_format_new();

        if (!style_format) {
            err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
            goto mem_error;
        }

        memcpy(style_format, lxlsx_workbook_format, sizeof(lxlsx_format));
        STAILQ_INSERT_TAIL(styles->dxf_formats, style_format, list_pointers);
    }

    styles->font_count = self->workbook->font_count;
    styles->border_count = self->workbook->border_count;
    styles->fill_count = self->workbook->fill_count;
    styles->num_format_count = self->workbook->num_format_count;
    styles->xf_count = self->workbook->used_xf_formats->unique_count;
    styles->dxf_count = self->workbook->used_dxf_formats->unique_count;
    styles->has_comments = self->workbook->has_comments;

    styles->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!styles->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_styles_assemble_xml_file(styles);

    err = _add_to_zip(self, styles->file, &buffer, &buffer_size,
                      "xl/styles.xml");

    fclose(styles->file);
    free(buffer);

mem_error:
    lxlsx_styles_free(styles);

    return err;
}

/*
 * Write the ContentTypes.xml file.
 */
STATIC lxlsx_error
_write_content_types_file(lxlsx_packager *self)
{
    lxlsx_content_types *content_types = lxlsx_content_types_new();
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    char filename[LXLSX_MAX_ATTRIBUTE_LENGTH] = { 0 };
    uint32_t index = 1;
    uint32_t lxlsx_worksheet_index = 1;
    uint32_t lxlsx_chartsheet_index = 1;
    uint32_t lxlsx_drawing_count = _get_drawing_count(self);
    uint32_t lxlsx_chart_count = _get_chart_count(self);
    uint32_t lxlsx_table_count = _get_table_count(self);
    lxlsx_error err = LXLSX_NO_ERROR;

    if (!content_types) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    content_types->file = lxlsx_get_filehandle(&buffer, &buffer_size,
                                             self->tmpdir);
    if (!content_types->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    if (workbook->has_png)
        lxlsx_ct_add_default(content_types, "png", "image/png");

    if (workbook->has_jpeg)
        lxlsx_ct_add_default(content_types, "jpeg", "image/jpeg");

    if (workbook->has_bmp)
        lxlsx_ct_add_default(content_types, "bmp", "image/bmp");

    if (workbook->has_gif)
        lxlsx_ct_add_default(content_types, "gif", "image/gif");

    if (workbook->vba_project)
        lxlsx_ct_add_default(content_types, "bin",
                           "application/vnd.ms-office.vbaProject");

    if (workbook->vba_project)
        lxlsx_ct_add_override(content_types, "/xl/workbook.xml",
                            LXLSX_APP_MSEXCEL "sheet.macroEnabled.main+xml");
    else
        lxlsx_ct_add_override(content_types, "/xl/workbook.xml",
                            LXLSX_APP_DOCUMENT "spreadsheetml.sheet.main+xml");

    if (workbook->vba_project_signature)
        lxlsx_ct_add_override(content_types, "/xl/vbaProjectSignature.bin",
                            "application/vnd.ms-office.vbaProjectSignature");

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet) {
            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "/xl/chartsheets/sheet%d.xml", lxlsx_chartsheet_index++);
            lxlsx_ct_add_chartsheet_name(content_types, filename);
        }
        else {
            lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                         "/xl/worksheets/sheet%d.xml", lxlsx_worksheet_index++);
            lxlsx_ct_add_worksheet_name(content_types, filename);
        }
    }

    for (index = 1; index <= lxlsx_chart_count; index++) {
        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH, "/xl/charts/chart%d.xml",
                     index);
        lxlsx_ct_add_chart_name(content_types, filename);
    }

    for (index = 1; index <= lxlsx_drawing_count; index++) {
        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "/xl/drawings/drawing%d.xml", index);
        lxlsx_ct_add_drawing_name(content_types, filename);
    }

    for (index = 1; index <= lxlsx_table_count; index++) {
        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "/xl/tables/table%d.xml", index);
        lxlsx_ct_add_table_name(content_types, filename);
    }

    if (workbook->has_vml)
        lxlsx_ct_add_vml_name(content_types);

    for (index = 1; index <= workbook->comment_count; index++) {
        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "/xl/comments%d.xml", index);
        lxlsx_ct_add_comment_name(content_types, filename);
    }

    if (workbook->sst->string_count)
        lxlsx_ct_add_shared_strings(content_types);

    if (!STAILQ_EMPTY(self->workbook->lxlsx_custom_properties))
        lxlsx_ct_add_custom_properties(content_types);

    if (workbook->has_metadata)
        lxlsx_ct_add_metadata(content_types);

    if (workbook->has_embedded_images)
        lxlsx_ct_add_rich_value(content_types);

    lxlsx_content_types_assemble_xml_file(content_types);

    err = _add_to_zip(self, content_types->file, &buffer, &buffer_size,
                      "[Content_Types].xml");

    fclose(content_types->file);
    free(buffer);

mem_error:
    lxlsx_content_types_free(content_types);

    return err;
}

/*
 * Write the workbook .rels xml file.
 */
STATIC lxlsx_error
_write_workbook_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels = lxlsx_relationships_new();
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    uint32_t lxlsx_worksheet_index = 1;
    uint32_t lxlsx_chartsheet_index = 1;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (!rels) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rels->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet) {
            lxlsx_snprintf(sheetname,
                         LXLSX_FILENAME_LENGTH,
                         "chartsheets/sheet%d.xml", lxlsx_chartsheet_index++);
            lxlsx_add_document_relationship(rels, "/chartsheet", sheetname);
        }
        else {
            lxlsx_snprintf(sheetname,
                         LXLSX_FILENAME_LENGTH,
                         "worksheets/sheet%d.xml", lxlsx_worksheet_index++);
            lxlsx_add_document_relationship(rels, "/worksheet", sheetname);
        }
    }

    lxlsx_add_document_relationship(rels, "/theme", "theme/theme1.xml");
    lxlsx_add_document_relationship(rels, "/styles", "styles.xml");

    if (workbook->sst->string_count)
        lxlsx_add_document_relationship(rels, "/sharedStrings",
                                      "sharedStrings.xml");

    if (workbook->vba_project)
        lxlsx_add_ms_package_relationship(rels, "/vbaProject",
                                        "vbaProject.bin");

    if (workbook->has_metadata)
        lxlsx_add_document_relationship(rels, "/sheetMetadata", "metadata.xml");

    if (workbook->has_embedded_images)
        lxlsx_add_rich_value_relationship(rels);

    lxlsx_relationships_assemble_xml_file(rels);

    err = _add_to_zip(self, rels->file, &buffer, &buffer_size,
                      "xl/_rels/workbook.xml.rels");

    fclose(rels->file);
    free(buffer);

mem_error:
    lxlsx_free_relationships(rels);

    return err;
}

/*
 * Write the worksheet .rels files for worksheets that contain links to
 * external data such as hyperlinks or drawings.
 */
STATIC lxlsx_error
_write_worksheet_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_rel_tuple *rel;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    uint32_t index = 0;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        index++;

        if (STAILQ_EMPTY(worksheet->external_hyperlinks) &&
            STAILQ_EMPTY(worksheet->external_drawing_links) &&
            STAILQ_EMPTY(worksheet->external_table_links) &&
            !worksheet->external_vml_header_link &&
            !worksheet->external_vml_comment_link &&
            !worksheet->external_background_link &&
            !worksheet->external_comment_link)
            continue;

        rels = lxlsx_relationships_new();

        rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
        if (!rels->file) {
            lxlsx_free_relationships(rels);
            return LXLSX_ERROR_CREATING_TMPFILE;
        }

        STAILQ_FOREACH(rel, worksheet->external_hyperlinks, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        STAILQ_FOREACH(rel, worksheet->external_drawing_links, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        rel = worksheet->external_vml_comment_link;
        if (rel)
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        rel = worksheet->external_vml_header_link;
        if (rel)
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        rel = worksheet->external_background_link;
        if (rel)
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        STAILQ_FOREACH(rel, worksheet->external_table_links, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        rel = worksheet->external_comment_link;
        if (rel)
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/worksheets/_rels/sheet%d.xml.rels", index);

        lxlsx_relationships_assemble_xml_file(rels);

        err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

        fclose(rels->file);
        free(buffer);
        lxlsx_free_relationships(rels);

        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the chartsheet .rels files for chartsheets that contain links to
 * external data such as drawings.
 */
STATIC lxlsx_error
_write_chartsheet_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_rel_tuple *rel;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    uint32_t index = 0;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            continue;

        index++;

        if (STAILQ_EMPTY(worksheet->external_drawing_links))
            continue;

        rels = lxlsx_relationships_new();

        rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
        if (!rels->file) {
            lxlsx_free_relationships(rels);
            return LXLSX_ERROR_CREATING_TMPFILE;
        }

        STAILQ_FOREACH(rel, worksheet->external_hyperlinks, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        STAILQ_FOREACH(rel, worksheet->external_drawing_links, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);
        }

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/chartsheets/_rels/sheet%d.xml.rels", index);

        lxlsx_relationships_assemble_xml_file(rels);

        err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

        fclose(rels->file);
        free(buffer);
        lxlsx_free_relationships(rels);

        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the drawing .rels files for worksheets that contain charts or
 * drawings.
 */
STATIC lxlsx_error
_write_drawing_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_rel_tuple *rel;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    uint32_t index = 1;
    lxlsx_error err;

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            worksheet = sheet->u.chartsheet->worksheet;
        else
            worksheet = sheet->u.worksheet;

        if (STAILQ_EMPTY(worksheet->lxlsx_drawing_links))
            continue;

        rels = lxlsx_relationships_new();

        rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
        if (!rels->file) {
            lxlsx_free_relationships(rels);
            return LXLSX_ERROR_CREATING_TMPFILE;
        }

        STAILQ_FOREACH(rel, worksheet->lxlsx_drawing_links, list_pointers) {
            lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                           rel->target_mode);

        }

        lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                     "xl/drawings/_rels/drawing%d.xml.rels", index++);

        lxlsx_relationships_assemble_xml_file(rels);

        err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

        fclose(rels->file);
        free(buffer);
        lxlsx_free_relationships(rels);

        RETURN_ON_ERROR(err);
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write the vmlDrawing .rels files for worksheets that contain images in
 * headers or footers.
 */
STATIC lxlsx_error
_write_vml_drawing_rels_file(lxlsx_packager *self, lxlsx_worksheet *worksheet,
                             uint32_t index)
{
    lxlsx_relationships *rels;
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_rel_tuple *rel;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    lxlsx_error err = LXLSX_NO_ERROR;

    rels = lxlsx_relationships_new();

    rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rels->file) {
        lxlsx_free_relationships(rels);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    STAILQ_FOREACH(rel, worksheet->lxlsx_vml_drawing_links, list_pointers) {
        lxlsx_add_worksheet_relationship(rels, rel->type, rel->target,
                                       rel->target_mode);

    }

    lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                 "xl/drawings/_rels/vmlDrawing%d.vml.rels", index);

    lxlsx_relationships_assemble_xml_file(rels);

    err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

    fclose(rels->file);
    free(buffer);
    lxlsx_free_relationships(rels);

    return err;
}

/*
 * Write the vbaProject .rels xml file.
 */
STATIC lxlsx_error
_write_vba_project_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels;
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_error err = LXLSX_NO_ERROR;
    char *buffer = NULL;
    size_t buffer_size = 0;

    if (!workbook->vba_project_signature)
        return LXLSX_NO_ERROR;

    rels = lxlsx_relationships_new();
    if (!rels) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rels->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_add_ms_package_relationship(rels, "/vbaProjectSignature",
                                    "vbaProjectSignature.bin");

    lxlsx_relationships_assemble_xml_file(rels);

    err = _add_to_zip(self, rels->file, &buffer, &buffer_size,
                      "xl/_rels/vbaProject.bin.rels");

    fclose(rels->file);
    free(buffer);

mem_error:
    lxlsx_free_relationships(rels);

    return err;
}

/*
 * Write the richValueRel.xml.rels files for embedded images.
 */
STATIC lxlsx_error
_write_rich_value_rels_file(lxlsx_packager *self)
{
    lxlsx_workbook *workbook = self->workbook;
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_object_properties *object_props;

    lxlsx_relationships *rels;
    char *buffer = NULL;
    size_t buffer_size = 0;
    char sheetname[LXLSX_FILENAME_LENGTH] = { 0 };
    char target[LXLSX_FILENAME_LENGTH] = { 0 };
    lxlsx_error err = LXLSX_NO_ERROR;
    uint32_t index = 1;

    if (!workbook->has_embedded_images)
        return LXLSX_NO_ERROR;

    rels = lxlsx_relationships_new();

    rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rels->file) {
        lxlsx_free_relationships(rels);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    STAILQ_FOREACH(sheet, workbook->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (STAILQ_EMPTY(worksheet->embedded_image_props))
            continue;

        STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
                       list_pointers) {

            if (object_props->is_duplicate)
                continue;

            lxlsx_snprintf(target, LXLSX_FILENAME_LENGTH,
                         "../media/image%d.%s", index++,
                         object_props->extension);

            lxlsx_add_document_relationship(rels, "/image", target);

        }

    }

    lxlsx_snprintf(sheetname, LXLSX_FILENAME_LENGTH,
                 "xl/richData/_rels/richValueRel.xml.rels");

    lxlsx_relationships_assemble_xml_file(rels);

    err = _add_to_zip(self, rels->file, &buffer, &buffer_size, sheetname);

    fclose(rels->file);
    free(buffer);
    lxlsx_free_relationships(rels);

    return err;
}

/*
 * Write the _rels/.rels xml file.
 */
STATIC lxlsx_error
_write_root_rels_file(lxlsx_packager *self)
{
    lxlsx_relationships *rels = lxlsx_relationships_new();
    char *buffer = NULL;
    size_t buffer_size = 0;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (!rels) {
        err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        goto mem_error;
    }

    rels->file = lxlsx_get_filehandle(&buffer, &buffer_size, self->tmpdir);
    if (!rels->file) {
        err = LXLSX_ERROR_CREATING_TMPFILE;
        goto mem_error;
    }

    lxlsx_add_document_relationship(rels, "/officeDocument", "xl/workbook.xml");

    lxlsx_add_package_relationship(rels,
                                 "/metadata/core-properties",
                                 "docProps/core.xml");

    lxlsx_add_document_relationship(rels,
                                  "/extended-properties", "docProps/app.xml");

    if (!STAILQ_EMPTY(self->workbook->lxlsx_custom_properties))
        lxlsx_add_document_relationship(rels,
                                      "/custom-properties",
                                      "docProps/custom.xml");

    lxlsx_relationships_assemble_xml_file(rels);

    err = _add_to_zip(self, rels->file, &buffer, &buffer_size, "_rels/.rels");

    fclose(rels->file);
    free(buffer);

mem_error:
    lxlsx_free_relationships(rels);

    return err;
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

STATIC lxlsx_error
_add_file_to_zip(lxlsx_packager *self, FILE *file, const char *filename)
{
    int16_t error = ZIP_OK;
    size_t size_read;

    error = zipOpenNewFileInZip4_64(self->zipfile,
                                    filename,
                                    &self->zipfile_info,
                                    NULL, 0, NULL, 0, NULL,
                                    Z_DEFLATED, Z_DEFAULT_COMPRESSION, 0,
                                    -MAX_WBITS, DEF_MEM_LEVEL,
                                    Z_DEFAULT_STRATEGY, NULL, 0, 0, 0,
                                    self->use_zip64);

    if (error != ZIP_OK) {
        LXLSX_ERROR("Error adding member to zipfile");
        RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
    }

    fflush(file);
    rewind(file);

    size_read = fread((void *) self->buffer, 1, self->buffer_size, file);

    while (size_read) {

        if (size_read < self->buffer_size) {
            if (ferror(file)) {
                LXLSX_ERROR("Error reading member file data");
                RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
            }
        }

        error = zipWriteInFileInZip(self->zipfile,
                                    self->buffer, (unsigned int) size_read);

        if (error < 0) {
            LXLSX_ERROR("Error in writing member in the zipfile");
            RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
        }

        size_read =
            fread((void *) (void *) self->buffer, 1, self->buffer_size, file);
    }

    error = zipCloseFileInZip(self->zipfile);
    if (error != ZIP_OK) {
        LXLSX_ERROR("Error in closing member in the zipfile");
        RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
    }

    return LXLSX_NO_ERROR;
}

STATIC lxlsx_error
_add_buffer_to_zip(lxlsx_packager *self, const char *buffer, size_t buffer_size,
                   const char *filename)
{
    int16_t error = ZIP_OK;

    error = zipOpenNewFileInZip4_64(self->zipfile,
                                    filename,
                                    &self->zipfile_info,
                                    NULL, 0, NULL, 0, NULL,
                                    Z_DEFLATED, Z_DEFAULT_COMPRESSION, 0,
                                    -MAX_WBITS, DEF_MEM_LEVEL,
                                    Z_DEFAULT_STRATEGY, NULL, 0, 0, 0,
                                    self->use_zip64);

    if (error != ZIP_OK) {
        LXLSX_ERROR("Error adding member to zipfile");
        RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
    }

    error = zipWriteInFileInZip(self->zipfile,
                                buffer, (unsigned int) buffer_size);

    if (error < 0) {
        LXLSX_ERROR("Error in writing member in the zipfile");
        RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
    }

    error = zipCloseFileInZip(self->zipfile);
    if (error != ZIP_OK) {
        LXLSX_ERROR("Error in closing member in the zipfile");
        RETURN_ON_ZIP_ERROR(error, LXLSX_ERROR_ZIP_FILE_ADD);
    }

    return LXLSX_NO_ERROR;
}

STATIC lxlsx_error
_add_to_zip(lxlsx_packager *self, FILE *file, char **buffer,
            size_t *buffer_size, const char *filename)
{
    /* Flush to ensure buffer is updated when using a memory-backed file. */
    fflush(file);
    return *buffer ?
        _add_buffer_to_zip(self, *buffer, *buffer_size, filename) :
        _add_file_to_zip(self, file, filename);
}

/*
 * Write the xml files that make up the XLSX OPC package.
 */
lxlsx_error
lxlsx_create_package(lxlsx_packager *self)
{
    lxlsx_error error;
    int8_t zip_error;

    error = _write_content_types_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_root_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_workbook_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_worksheet_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_chartsheet_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_workbook_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_chart_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_drawing_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_vml_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_comment_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_table_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_shared_strings_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_custom_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_theme_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_styles_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_worksheet_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_chartsheet_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_drawing_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_image_files(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _add_vba_project(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _add_vba_project_signature(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_vba_project_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_core_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_metadata_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_rich_value_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_rich_value_rel_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_rich_value_types_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_rich_value_structure_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_rich_value_rels_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    error = _write_app_file(self);
    RETURN_AND_ZIPCLOSE_ON_ERROR(error);

    zip_error = zipClose(self->zipfile, NULL);
    if (zip_error) {
        RETURN_ON_ZIP_ERROR(zip_error, LXLSX_ERROR_ZIP_CLOSE);
    }

    return LXLSX_NO_ERROR;
}
