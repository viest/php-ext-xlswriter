/*****************************************************************************
 * workbook - A library for creating Excel XLSX workbook files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/workbook.h"
#include "lxlsx/utility.h"
#include "lxlsx/packager.h"
#include "lxlsx/hash_table.h"
#include "lxlsx/hash_table.h"

STATIC int _worksheet_name_cmp(lxlsx_worksheet_name *name1,
                               lxlsx_worksheet_name *name2);
STATIC int _chartsheet_name_cmp(lxlsx_chartsheet_name *name1,
                                lxlsx_chartsheet_name *name2);
STATIC int _image_md5_cmp(lxlsx_image_md5 *tuple1, lxlsx_image_md5 *tuple2);

#ifndef __clang_analyzer__
LXLSX_RB_GENERATE_WORKSHEET_NAMES(lxlsx_worksheet_names, lxlsx_worksheet_name,
                                tree_pointers, _worksheet_name_cmp);
LXLSX_RB_GENERATE_CHARTSHEET_NAMES(lxlsx_chartsheet_names, lxlsx_chartsheet_name,
                                 tree_pointers, _chartsheet_name_cmp);
LXLSX_RB_GENERATE_IMAGE_MD5S(lxlsx_image_md5s, lxlsx_image_md5,
                           tree_pointers, _image_md5_cmp);
#endif

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Comparators for the sheet names structure red/black tree.
 */
STATIC int
_worksheet_name_cmp(lxlsx_worksheet_name *name1, lxlsx_worksheet_name *name2)
{
    return lxlsx_strcasecmp(name1->name, name2->name);
}

STATIC int
_chartsheet_name_cmp(lxlsx_chartsheet_name *name1, lxlsx_chartsheet_name *name2)
{
    return lxlsx_strcasecmp(name1->name, name2->name);
}

STATIC int
_image_md5_cmp(lxlsx_image_md5 *tuple1, lxlsx_image_md5 *tuple2)
{
    return strcmp(tuple1->md5, tuple2->md5);
}

/*
 * Free workbook properties.
 */
STATIC void
_free_doc_properties(lxlsx_doc_properties *properties)
{
    if (properties) {
        free((void *) properties->title);
        free((void *) properties->subject);
        free((void *) properties->author);
        free((void *) properties->manager);
        free((void *) properties->company);
        free((void *) properties->category);
        free((void *) properties->keywords);
        free((void *) properties->comments);
        free((void *) properties->status);
        free((void *) properties->hyperlink_base);
    }

    free(properties);
}

/*
 * Free workbook custom property.
 */
STATIC void
_free_custom_doc_property(lxlsx_custom_property *lxlsx_custom_property)
{
    if (lxlsx_custom_property) {
        free(lxlsx_custom_property->name);
        if (lxlsx_custom_property->type == LXLSX_CUSTOM_STRING)
            free(lxlsx_custom_property->u.string);
    }

    free(lxlsx_custom_property);
}

/*
 * Free a workbook object.
 */
void
lxlsx_workbook_free(lxlsx_workbook *workbook)
{
    lxlsx_sheet *sheet;
    struct lxlsx_worksheet_name *lxlsx_worksheet_name;
    struct lxlsx_worksheet_name *next_worksheet_name;
    struct lxlsx_chartsheet_name *lxlsx_chartsheet_name;
    struct lxlsx_chartsheet_name *next_chartsheet_name;
    struct lxlsx_image_md5 *image_md5;
    struct lxlsx_image_md5 *next_image_md5;
    lxlsx_chart *chart;
    lxlsx_format *format;
    lxlsx_defined_name *defined_name;
    lxlsx_defined_name *defined_name_tmp;
    lxlsx_custom_property *lxlsx_custom_property;

    if (!workbook)
        return;

    _free_doc_properties(workbook->properties);

    free(workbook->filename);

    /* Free the sheets in the workbook. */
    if (workbook->sheets) {
        while (!STAILQ_EMPTY(workbook->sheets)) {
            sheet = STAILQ_FIRST(workbook->sheets);

            if (sheet->is_chartsheet)
                lxlsx_chartsheet_free(sheet->u.chartsheet);
            else
                lxlsx_worksheet_free(sheet->u.worksheet);

            STAILQ_REMOVE_HEAD(workbook->sheets, list_pointers);
            free(sheet);
        }
        free(workbook->sheets);
    }

    /* Free the sheet lists. The worksheet objects are freed above. */
    free(workbook->worksheets);
    free(workbook->chartsheets);

    /* Free the charts in the workbook. */
    if (workbook->charts) {
        while (!STAILQ_EMPTY(workbook->charts)) {
            chart = STAILQ_FIRST(workbook->charts);
            STAILQ_REMOVE_HEAD(workbook->charts, list_pointers);
            lxlsx_chart_free(chart);
        }
        free(workbook->charts);
    }

    /* Free the formats in the workbook. */
    if (workbook->formats) {
        while (!STAILQ_EMPTY(workbook->formats)) {
            format = STAILQ_FIRST(workbook->formats);
            STAILQ_REMOVE_HEAD(workbook->formats, list_pointers);
            lxlsx_format_free(format);
        }
        free(workbook->formats);
    }

    /* Free the defined_names in the workbook. */
    if (workbook->defined_names) {
        defined_name = TAILQ_FIRST(workbook->defined_names);
        while (defined_name) {

            defined_name_tmp = TAILQ_NEXT(defined_name, list_pointers);
            free(defined_name);
            defined_name = defined_name_tmp;
        }
        free(workbook->defined_names);
    }

    /* Free the lxlsx_custom_properties in the workbook. */
    if (workbook->lxlsx_custom_properties) {
        while (!STAILQ_EMPTY(workbook->lxlsx_custom_properties)) {
            lxlsx_custom_property = STAILQ_FIRST(workbook->lxlsx_custom_properties);
            STAILQ_REMOVE_HEAD(workbook->lxlsx_custom_properties, list_pointers);
            _free_custom_doc_property(lxlsx_custom_property);
        }
        free(workbook->lxlsx_custom_properties);
    }

    if (workbook->lxlsx_worksheet_names) {
        for (lxlsx_worksheet_name =
             RB_MIN(lxlsx_worksheet_names, workbook->lxlsx_worksheet_names);
             lxlsx_worksheet_name; lxlsx_worksheet_name = next_worksheet_name) {

            next_worksheet_name = RB_NEXT(lxlsx_worksheet_names,
                                          workbook->lxlsx_worksheet_name,
                                          lxlsx_worksheet_name);
            RB_REMOVE(lxlsx_worksheet_names, workbook->lxlsx_worksheet_names,
                      lxlsx_worksheet_name);
            free(lxlsx_worksheet_name);
        }

        free(workbook->lxlsx_worksheet_names);
    }

    if (workbook->lxlsx_chartsheet_names) {
        for (lxlsx_chartsheet_name =
             RB_MIN(lxlsx_chartsheet_names, workbook->lxlsx_chartsheet_names);
             lxlsx_chartsheet_name; lxlsx_chartsheet_name = next_chartsheet_name) {

            next_chartsheet_name = RB_NEXT(lxlsx_chartsheet_names,
                                           workbook->lxlsx_chartsheet_name,
                                           lxlsx_chartsheet_name);
            RB_REMOVE(lxlsx_chartsheet_names, workbook->lxlsx_chartsheet_names,
                      lxlsx_chartsheet_name);
            free(lxlsx_chartsheet_name);
        }

        free(workbook->lxlsx_chartsheet_names);
    }

    if (workbook->image_md5s) {
        for (image_md5 = RB_MIN(lxlsx_image_md5s, workbook->image_md5s);
             image_md5; image_md5 = next_image_md5) {

            next_image_md5 =
                RB_NEXT(lxlsx_image_md5s, workbook->image_md5, image_md5);
            RB_REMOVE(lxlsx_image_md5s, workbook->image_md5s, image_md5);
            free(image_md5->md5);
            free(image_md5);
        }

        free(workbook->image_md5s);
    }

    if (workbook->embedded_image_md5s) {
        for (image_md5 =
             RB_MIN(lxlsx_image_md5s, workbook->embedded_image_md5s); image_md5;
             image_md5 = next_image_md5) {

            next_image_md5 =
                RB_NEXT(lxlsx_image_md5s, workbook->embedded_image_md5s,
                        image_md5);
            RB_REMOVE(lxlsx_image_md5s, workbook->embedded_image_md5s,
                      image_md5);
            free(image_md5->md5);
            free(image_md5);
        }

        free(workbook->embedded_image_md5s);
    }

    if (workbook->header_image_md5s) {
        for (image_md5 = RB_MIN(lxlsx_image_md5s, workbook->header_image_md5s);
             image_md5; image_md5 = next_image_md5) {

            next_image_md5 =
                RB_NEXT(lxlsx_image_md5s, workbook->header_image_md5s,
                        image_md5);
            RB_REMOVE(lxlsx_image_md5s, workbook->header_image_md5s, image_md5);
            free(image_md5->md5);
            free(image_md5);
        }

        free(workbook->header_image_md5s);
    }

    if (workbook->background_md5s) {
        for (image_md5 = RB_MIN(lxlsx_image_md5s, workbook->background_md5s);
             image_md5; image_md5 = next_image_md5) {

            next_image_md5 =
                RB_NEXT(lxlsx_image_md5s, workbook->background_md5s, image_md5);
            RB_REMOVE(lxlsx_image_md5s, workbook->background_md5s, image_md5);
            free(image_md5->md5);
            free(image_md5);
        }

        free(workbook->background_md5s);
    }

    lxlsx_hash_free(workbook->used_xf_formats);
    lxlsx_hash_free(workbook->used_dxf_formats);
    lxlsx_sst_free(workbook->sst);
    free((void *) workbook->options.tmpdir);
    free(workbook->ordered_charts);
    free(workbook->vba_project);
    free(workbook->vba_project_signature);
    free(workbook->vba_codename);
    free(workbook);
}

/*
 * Set the default index for each format. This is only used for testing.
 */
void
lxlsx_workbook_set_default_xf_indices(lxlsx_workbook *self)
{
    lxlsx_format *format;
    int32_t index = 0;

    STAILQ_FOREACH(format, self->formats, list_pointers) {

        /* Skip the hyperlink format. */
        if (index != 1)
            lxlsx_format_get_xf_index(format);

        index++;
    }
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * font elements.
 */
STATIC void
_prepare_fonts(lxlsx_workbook *self)
{

    lxlsx_hash_table *fonts = lxlsx_hash_new(128, 1, 1);
    lxlsx_hash_element *hash_element;
    lxlsx_hash_element *used_format_element;
    uint16_t index = 0;

    LXLSX_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;
        lxlsx_font *key = lxlsx_format_get_font_key(format);

        if (key) {
            /* Look up the format in the hash table. */
            hash_element = lxlsx_hash_key_exists(fonts, key, sizeof(lxlsx_font));

            if (hash_element) {
                /* Font has already been used. */
                format->font_index = *(uint16_t *) hash_element->value;
                format->has_font = LXLSX_FALSE;
                free(key);
            }
            else {
                /* This is a new font. */
                uint16_t *font_index = calloc(1, sizeof(uint16_t));
                *font_index = index;
                format->font_index = index;
                format->has_font = LXLSX_TRUE;
                lxlsx_insert_hash_element(fonts, key, font_index,
                                        sizeof(lxlsx_font));
                index++;
            }
        }
    }

    lxlsx_hash_free(fonts);

    /* For DXF formats we only need to check if the properties have changed. */
    LXLSX_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;

        /* The only font properties that can change for a DXF format are:
         * color, bold, italic, underline and strikethrough. */
        if (format->font_color || format->bold || format->italic
            || format->underline || format->font_strikeout) {
            format->has_dxf_font = LXLSX_TRUE;
        }
    }

    self->font_count = index;
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * border elements.
 */
STATIC void
_prepare_borders(lxlsx_workbook *self)
{

    lxlsx_hash_table *borders = lxlsx_hash_new(128, 1, 1);
    lxlsx_hash_element *hash_element;
    lxlsx_hash_element *used_format_element;
    uint16_t index = 0;

    LXLSX_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;
        lxlsx_border *key = lxlsx_format_get_border_key(format);

        if (key) {
            /* Look up the format in the hash table. */
            hash_element =
                lxlsx_hash_key_exists(borders, key, sizeof(lxlsx_border));

            if (hash_element) {
                /* Border has already been used. */
                format->border_index = *(uint16_t *) hash_element->value;
                format->has_border = LXLSX_FALSE;
                free(key);
            }
            else {
                /* This is a new border. */
                uint16_t *border_index = calloc(1, sizeof(uint16_t));
                *border_index = index;
                format->border_index = index;
                format->has_border = 1;
                lxlsx_insert_hash_element(borders, key, border_index,
                                        sizeof(lxlsx_border));
                index++;
            }
        }
    }

    /* For DXF formats we only need to check if the properties have changed. */
    LXLSX_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;

        if (format->left || format->right || format->top || format->bottom) {
            format->has_dxf_border = LXLSX_TRUE;
        }
    }

    lxlsx_hash_free(borders);

    self->border_count = index;
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * fill elements.
 */
STATIC void
_prepare_fills(lxlsx_workbook *self)
{

    lxlsx_hash_table *fills = lxlsx_hash_new(128, 1, 1);
    lxlsx_hash_element *hash_element;
    lxlsx_hash_element *used_format_element;
    uint16_t index = 2;
    lxlsx_fill *default_fill_1 = NULL;
    lxlsx_fill *default_fill_2 = NULL;
    uint16_t *fill_index1 = NULL;
    uint16_t *fill_index2 = NULL;

    default_fill_1 = calloc(1, sizeof(lxlsx_fill));
    GOTO_LABEL_ON_MEM_ERROR(default_fill_1, mem_error);

    default_fill_2 = calloc(1, sizeof(lxlsx_fill));
    GOTO_LABEL_ON_MEM_ERROR(default_fill_2, mem_error);

    fill_index1 = calloc(1, sizeof(uint16_t));
    GOTO_LABEL_ON_MEM_ERROR(fill_index1, mem_error);

    fill_index2 = calloc(1, sizeof(uint16_t));
    GOTO_LABEL_ON_MEM_ERROR(fill_index2, mem_error);

    /* Add the default fills. */
    default_fill_1->pattern = LXLSX_PATTERN_NONE;
    default_fill_1->fg_color = LXLSX_COLOR_UNSET;
    default_fill_1->bg_color = LXLSX_COLOR_UNSET;
    *fill_index1 = 0;
    lxlsx_insert_hash_element(fills, default_fill_1, fill_index1,
                            sizeof(lxlsx_fill));

    default_fill_2->pattern = LXLSX_PATTERN_GRAY_125;
    default_fill_2->fg_color = LXLSX_COLOR_UNSET;
    default_fill_2->bg_color = LXLSX_COLOR_UNSET;
    *fill_index2 = 1;
    lxlsx_insert_hash_element(fills, default_fill_2, fill_index2,
                            sizeof(lxlsx_fill));

    /* For DXF formats we only need to check if the properties have changed. */
    LXLSX_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;

        if (format->pattern || format->bg_color || format->fg_color) {
            format->has_dxf_fill = LXLSX_TRUE;
            format->dxf_bg_color = format->bg_color;
            format->dxf_fg_color = format->fg_color;
        }
    }

    LXLSX_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;
        lxlsx_fill *key = lxlsx_format_get_fill_key(format);

        /* The following logical statements jointly take care of special */
        /* cases in relation to cell colors and patterns:                */
        /* 1. For a solid fill (pattern == 1) Excel reverses the role of */
        /*    foreground and background colors, and                      */
        /* 2. If the user specifies a foreground or background color     */
        /*    without a pattern they probably wanted a solid fill, so    */
        /*    we fill in the defaults.                                   */
        if (format->pattern == LXLSX_PATTERN_SOLID
            && format->bg_color != LXLSX_COLOR_UNSET
            && format->fg_color != LXLSX_COLOR_UNSET) {
            lxlsx_color_t tmp = format->fg_color;
            format->fg_color = format->bg_color;
            format->bg_color = tmp;
        }

        if (format->pattern <= LXLSX_PATTERN_SOLID
            && format->bg_color != LXLSX_COLOR_UNSET
            && format->fg_color == LXLSX_COLOR_UNSET) {
            format->fg_color = format->bg_color;
            format->bg_color = LXLSX_COLOR_UNSET;
            format->pattern = LXLSX_PATTERN_SOLID;
        }

        if (format->pattern <= LXLSX_PATTERN_SOLID
            && format->bg_color == LXLSX_COLOR_UNSET
            && format->fg_color != LXLSX_COLOR_UNSET) {
            format->pattern = LXLSX_PATTERN_SOLID;
        }

        if (key) {
            /* Look up the format in the hash table. */
            hash_element = lxlsx_hash_key_exists(fills, key, sizeof(lxlsx_fill));

            if (hash_element) {
                /* Fill has already been used. */
                format->fill_index = *(uint16_t *) hash_element->value;
                format->has_fill = LXLSX_FALSE;
                free(key);
            }
            else {
                /* This is a new fill. */
                uint16_t *fill_index = calloc(1, sizeof(uint16_t));
                *fill_index = index;
                format->fill_index = index;
                format->has_fill = 1;
                lxlsx_insert_hash_element(fills, key, fill_index,
                                        sizeof(lxlsx_fill));
                index++;
            }
        }
    }

    lxlsx_hash_free(fills);

    self->fill_count = index;

    return;

mem_error:
    free(fill_index2);
    free(fill_index1);
    free(default_fill_2);
    free(default_fill_1);
    lxlsx_hash_free(fills);
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * number format elements. Note, user defined records start from index 0xA4.
 */
STATIC void
_prepare_num_formats(lxlsx_workbook *self)
{

    lxlsx_hash_table *num_formats = lxlsx_hash_new(128, 0, 1);
    lxlsx_hash_element *hash_element;
    lxlsx_hash_element *used_format_element;
    uint16_t index = 0xA4;
    uint16_t num_format_count = 0;
    uint16_t *num_format_index;

    LXLSX_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;

        /* Format already has a number format index. */
        if (format->num_format_index)
            continue;

        /* Check if there is a user defined number format string. */
        if (*format->num_format) {
            char num_format[LXLSX_FORMAT_FIELD_LEN] = { 0 };
            lxlsx_snprintf(num_format, LXLSX_FORMAT_FIELD_LEN, "%s",
                         format->num_format);

            /* Look up the num_format in the hash table. */
            hash_element = lxlsx_hash_key_exists(num_formats, num_format,
                                               LXLSX_FORMAT_FIELD_LEN);

            if (hash_element) {
                /* Num_Format has already been used. */
                format->num_format_index = *(uint16_t *) hash_element->value;
            }
            else {
                /* This is a new num_format. */
                num_format_index = calloc(1, sizeof(uint16_t));
                *num_format_index = index;
                format->num_format_index = index;
                lxlsx_insert_hash_element(num_formats, format->num_format,
                                        num_format_index,
                                        LXLSX_FORMAT_FIELD_LEN);
                index++;
                num_format_count++;
            }
        }
    }

    LXLSX_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
        lxlsx_format *format = (lxlsx_format *) used_format_element->value;

        /* Format already has a number format index. */
        if (format->num_format_index)
            continue;

        /* Check if there is a user defined number format string. */
        if (*format->num_format) {
            char num_format[LXLSX_FORMAT_FIELD_LEN] = { 0 };
            lxlsx_snprintf(num_format, LXLSX_FORMAT_FIELD_LEN, "%s",
                         format->num_format);

            /* Look up the num_format in the hash table. */
            hash_element = lxlsx_hash_key_exists(num_formats, num_format,
                                               LXLSX_FORMAT_FIELD_LEN);

            if (hash_element) {
                /* Num_Format has already been used. */
                format->num_format_index = *(uint16_t *) hash_element->value;
            }
            else {
                /* This is a new num_format. */
                num_format_index = calloc(1, sizeof(uint16_t));
                *num_format_index = index;
                format->num_format_index = index;
                lxlsx_insert_hash_element(num_formats, format->num_format,
                                        num_format_index,
                                        LXLSX_FORMAT_FIELD_LEN);
                index++;
                /* Don't update num_format_count for DXF formats. */
            }
        }
    }

    lxlsx_hash_free(num_formats);

    self->num_format_count = num_format_count;
}

/*
 * Prepare workbook and sub-objects for writing.
 */
STATIC void
_prepare_workbook(lxlsx_workbook *self)
{
    /* Set the font index for the format objects. */
    _prepare_fonts(self);

    /* Set the number format index for the format objects. */
    _prepare_num_formats(self);

    /* Set the border index for the format objects. */
    _prepare_borders(self);

    /* Set the fill index for the format objects. */
    _prepare_fills(self);

}

/*
 * Compare two defined_name structures.
 */
static int
_compare_defined_names(lxlsx_defined_name *a, lxlsx_defined_name *b)
{
    int res = strcmp(a->normalised_name, b->normalised_name);

    /* Primary comparison based on defined name. */
    if (res)
        return res;

    /* Secondary comparison based on worksheet name. */
    res = strcmp(a->normalised_sheetname, b->normalised_sheetname);

    return res;
}

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
STATIC lxlsx_error
_store_defined_name(lxlsx_workbook *self, const char *name,
                    const char *lxlsx_app_name, const char *formula, int16_t index,
                    uint8_t hidden)
{
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_defined_name *defined_name;
    lxlsx_defined_name *list_defined_name;
    char name_copy[LXLSX_DEFINED_NAME_LENGTH];
    char *tmp_str;
    char *lxlsx_worksheet_name;

    /* Do some checks on the input data */
    if (!name || !formula)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_str_is_empty(name) || lxlsx_str_is_empty(formula))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    if (lxlsx_utf8_strlen(name) > LXLSX_DEFINED_NAME_LENGTH ||
        lxlsx_utf8_strlen(formula) > LXLSX_DEFINED_NAME_LENGTH) {
        return LXLSX_ERROR_128_STRING_LENGTH_EXCEEDED;
    }

    /* Allocate a new defined_name to be added to the linked list of names. */
    defined_name = calloc(1, sizeof(struct lxlsx_defined_name));
    RETURN_ON_MEM_ERROR(defined_name, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Copy the user input string. */
    lxlsx_strcpy(name_copy, name);

    /* Set the worksheet index or -1 for a global defined name. */
    defined_name->index = index;
    defined_name->hidden = hidden;

    /* Check for local defined names like like "Sheet1!name". */
    tmp_str = strchr(name_copy, '!');

    if (tmp_str == NULL) {
        /* The name is global. We just store the defined name string. */
        lxlsx_strcpy(defined_name->name, name_copy);
    }
    else {
        /* The name is worksheet local. We need to extract the sheet name
         * and map it to a sheet index. */

        /* Split the into the worksheet name and defined name. */
        *tmp_str = '\0';
        tmp_str++;
        lxlsx_worksheet_name = name_copy;

        if (lxlsx_str_is_empty(tmp_str) || lxlsx_str_is_empty(lxlsx_worksheet_name))
            goto mem_error;

        /* Remove any worksheet quoting. */
        if (lxlsx_worksheet_name[0] == '\'')
            lxlsx_worksheet_name++;
        if (strlen(lxlsx_worksheet_name) > 0
            && lxlsx_worksheet_name[strlen(lxlsx_worksheet_name) - 1] == '\'') {
            lxlsx_worksheet_name[strlen(lxlsx_worksheet_name) - 1] = '\0';
        }

        /* Search for worksheet name to get the equivalent worksheet index. */
        STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
            if (sheet->is_chartsheet)
                continue;
            else
                worksheet = sheet->u.worksheet;

            if (strcmp(lxlsx_worksheet_name, worksheet->name) == 0) {
                defined_name->index = worksheet->index;
                lxlsx_strcpy(defined_name->normalised_sheetname,
                           lxlsx_worksheet_name);
            }
        }

        /* If we didn't find the worksheet name we exit. */
        if (defined_name->index == -1)
            goto mem_error;

        lxlsx_strcpy(defined_name->name, tmp_str);
    }

    /* Print titles and repeat title pass in the name used for App.xml. */
    if (lxlsx_app_name) {
        lxlsx_strcpy(defined_name->lxlsx_app_name, lxlsx_app_name);
        lxlsx_strcpy(defined_name->normalised_sheetname, lxlsx_app_name);
    }
    else {
        lxlsx_strcpy(defined_name->lxlsx_app_name, name);
    }

    /* We need to normalize the defined names for sorting. This involves
     * removing any _xlnm namespace  and converting it to lowercase. */
    tmp_str = strstr(name_copy, "_xlnm.");

    if (tmp_str)
        lxlsx_strcpy(defined_name->normalised_name, defined_name->name + 6);
    else
        lxlsx_strcpy(defined_name->normalised_name, defined_name->name);

    lxlsx_str_tolower(defined_name->normalised_name);
    lxlsx_str_tolower(defined_name->normalised_sheetname);

    /* Strip leading "=" from the formula. */
    if (formula[0] == '=')
        lxlsx_strcpy(defined_name->formula, formula + 1);
    else
        lxlsx_strcpy(defined_name->formula, formula);

    /* We add the defined name to the list in sorted order. */
    list_defined_name = TAILQ_FIRST(self->defined_names);

    if (list_defined_name == NULL ||
        _compare_defined_names(defined_name, list_defined_name) < 1) {
        /* List is empty or defined name goes to the head. */
        TAILQ_INSERT_HEAD(self->defined_names, defined_name, list_pointers);
        return LXLSX_NO_ERROR;
    }

    TAILQ_FOREACH(list_defined_name, self->defined_names, list_pointers) {
        int res = _compare_defined_names(defined_name, list_defined_name);

        /* The entry already exists. We exit and don't overwrite. */
        if (res == 0)
            goto mem_error;

        /* New defined name is inserted in sorted order before other entries. */
        if (res < 0) {
            TAILQ_INSERT_BEFORE(list_defined_name, defined_name,
                                list_pointers);
            return LXLSX_NO_ERROR;
        }
    }

    /* If the entry wasn't less than any of the entries in the list we add it
     * to the end. */
    TAILQ_INSERT_TAIL(self->defined_names, defined_name, list_pointers);
    return LXLSX_NO_ERROR;

mem_error:
    free(defined_name);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Populate the data cache of a chart data series by reading the data from the
 * relevant worksheet and adding it to the cached in the range object as a
 * list of points.
 *
 * Note, the data cache isn't strictly required by Excel but it helps if the
 * chart is embedded in another application such as PowerPoint and it also
 * helps with comparison testing.
 */
STATIC void
_populate_range_data_cache(lxlsx_workbook *self, lxlsx_series_range *range)
{
    lxlsx_worksheet *worksheet;
    lxlsx_row_t row_num;
    lxlsx_col_t col_num;
    lxlsx_row *row_obj;
    lxlsx_cell *cell_obj;
    struct lxlsx_series_data_point *data_point;
    uint16_t num_data_points = 0;

    /* If ignore_cache is set then don't try to populate the cache. This flag
     * may be set manually, for testing, or due to a case where the cache
     * can't be calculated.
     */
    if (range->ignore_cache)
        return;

    /* Currently we only handle 2D ranges so ensure either the rows or cols
     * are the same.
     */
    if (range->first_row != range->last_row
        && range->first_col != range->last_col) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Check that the sheetname exists. */
    worksheet = lxlsx_workbook_get_worksheet_by_name(self, range->sheetname);
    if (!worksheet) {
        LXLSX_WARN_FORMAT2("lxlsx_workbook_add_chart(): worksheet name '%s' "
                         "in chart formula '%s' doesn't exist.",
                         range->sheetname, range->formula);
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* We can't read the data when worksheet optimization is on. */
    if (worksheet->optimize) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Iterate through the worksheet data and populate the range cache. */
    for (row_num = range->first_row; row_num <= range->last_row; row_num++) {
        row_obj = lxlsx_worksheet_find_row(worksheet, row_num);

        for (col_num = range->first_col; col_num <= range->last_col;
             col_num++) {

            data_point = calloc(1, sizeof(struct lxlsx_series_data_point));
            if (!data_point) {
                range->ignore_cache = LXLSX_TRUE;
                return;
            }

            cell_obj = lxlsx_worksheet_find_cell_in_row(row_obj, col_num);

            if (cell_obj) {
                if (cell_obj->type == NUMBER_CELL) {
                    data_point->number = cell_obj->u.number;
                }

                if (cell_obj->type == STRING_CELL) {
                    data_point->string = lxlsx_strdup(cell_obj->lxlsx_sst_string);
                    data_point->is_string = LXLSX_TRUE;
                    range->has_string_cache = LXLSX_TRUE;
                }
            }
            else {
                data_point->no_data = LXLSX_TRUE;
            }

            STAILQ_INSERT_TAIL(range->data_cache, data_point, list_pointers);
            num_data_points++;
        }
    }

    range->num_data_points = num_data_points;

}

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
STATIC void
_populate_range_dimensions(lxlsx_workbook *self, lxlsx_series_range *range)
{

    char formula[LXLSX_MAX_FORMULA_RANGE_LENGTH] = { 0 };
    char *tmp_str;
    char *sheetname;

    /* If neither the range formula or sheetname is defined then this probably
     * isn't a valid range.
     */
    if (!range->formula && !range->sheetname) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* If the sheetname is already defined it was already set via
     * lxlsx_chart_series_set_categories() or  lxlsx_chart_series_set_values().
     */
    if (range->sheetname)
        return;

    /* Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5) */
    if (range->formula[0] == '(') {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }

    /* Create a copy of the formula to modify and parse into parts. */
    lxlsx_snprintf(formula, LXLSX_MAX_FORMULA_RANGE_LENGTH, "%s", range->formula);

    /* Check for valid formula. Note, This needs stronger validation. */
    tmp_str = strchr(formula, '!');

    if (tmp_str == NULL) {
        range->ignore_cache = LXLSX_TRUE;
        return;
    }
    else {
        /* Check for empty string. */
        if (lxlsx_str_is_empty(tmp_str)) {
            range->ignore_cache = LXLSX_TRUE;
            return;
        }

        /* Split the formulas into sheetname and row-col data. */
        *tmp_str = '\0';
        tmp_str++;
        sheetname = formula;

        if (lxlsx_str_is_empty(tmp_str) || lxlsx_str_is_empty(sheetname)) {
            range->ignore_cache = LXLSX_TRUE;
            return;
        }

        /* Remove any worksheet quoting. */
        if (sheetname[0] == '\'')
            sheetname++;
        if (strlen(sheetname) > 0 && sheetname[strlen(sheetname) - 1] == '\'') {
            sheetname[strlen(sheetname) - 1] = '\0';
        }

        /* Check that the sheetname exists. */
        if (!lxlsx_workbook_get_worksheet_by_name(self, sheetname)) {
            LXLSX_WARN_FORMAT2("lxlsx_workbook_add_chart(): worksheet name '%s' "
                             "in chart formula '%s' doesn't exist.",
                             sheetname, range->formula);
            range->ignore_cache = LXLSX_TRUE;
            return;
        }

        range->sheetname = lxlsx_strdup(sheetname);
        range->first_row = lxlsx_name_to_row(tmp_str);
        range->first_col = lxlsx_name_to_col(tmp_str);

        if (strchr(tmp_str, ':')) {
            /* 2D range. */
            range->last_row = lxlsx_name_to_row_2(tmp_str);
            range->last_col = lxlsx_name_to_col_2(tmp_str);
        }
        else {
            /* 1D range. */
            range->last_row = range->first_row;
            range->last_col = range->first_col;
        }

    }
}

/* Set the range dimensions and set the data cache.
 */
STATIC void
_populate_range(lxlsx_workbook *self, lxlsx_series_range *range)
{
    if (!range)
        return;

    _populate_range_dimensions(self, range);
    _populate_range_data_cache(self, range);
}

/*
 * Add "cached" data to charts to provide the numCache and strCache data for
 * series and title/axis ranges.
 */
STATIC void
_add_chart_cache_data(lxlsx_workbook *self)
{
    lxlsx_chart *chart;
    lxlsx_chart_series *series;
    uint16_t i;

    STAILQ_FOREACH(chart, self->ordered_charts, ordered_list_pointers) {

        _populate_range(self, chart->title.range);
        _populate_range(self, chart->x_axis->title.range);
        _populate_range(self, chart->y_axis->title.range);

        if (STAILQ_EMPTY(chart->series_list))
            continue;

        STAILQ_FOREACH(series, chart->series_list, list_pointers) {
            _populate_range(self, series->categories);
            _populate_range(self, series->values);
            _populate_range(self, series->title.range);

            for (i = 0; i < series->data_label_count; i++) {
                lxlsx_chart_custom_label *data_label = &series->data_labels[i];
                _populate_range(self, data_label->range);
            }
        }
    }
}

/*
 * Store the image types used in the workbook to update the content types.
 */
STATIC void
_store_image_type(lxlsx_workbook *self, uint8_t image_type)
{
    if (image_type == LXLSX_IMAGE_PNG)
        self->has_png = LXLSX_TRUE;

    if (image_type == LXLSX_IMAGE_JPEG)
        self->has_jpeg = LXLSX_TRUE;

    if (image_type == LXLSX_IMAGE_BMP)
        self->has_bmp = LXLSX_TRUE;

    if (image_type == LXLSX_IMAGE_GIF)
        self->has_gif = LXLSX_TRUE;
}

/*
 * Iterate through the worksheets and set up any chart or image drawings.
 */
STATIC void
_prepare_drawings(lxlsx_workbook *self)
{
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_object_properties *object_props;
    uint32_t lxlsx_chart_ref_id = 0;
    uint32_t image_ref_id = 0;
    uint32_t ref_id = 0;
    uint32_t lxlsx_drawing_id = 0;
    uint8_t is_chartsheet;
    lxlsx_image_md5 tmp_image_md5;
    lxlsx_image_md5 *new_image_md5 = NULL;
    lxlsx_image_md5 *found_duplicate_image = NULL;
    uint8_t i;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet) {
            worksheet = sheet->u.chartsheet->worksheet;
            is_chartsheet = LXLSX_TRUE;
        }
        else {
            worksheet = sheet->u.worksheet;
            is_chartsheet = LXLSX_FALSE;
        }

        if (STAILQ_EMPTY(worksheet->image_props)
            && STAILQ_EMPTY(worksheet->embedded_image_props)
            && STAILQ_EMPTY(worksheet->lxlsx_chart_data)
            && !worksheet->has_header_vml && !worksheet->has_background_image) {
            continue;
        }

        lxlsx_drawing_id++;

        /* Prepare embedded worksheet images. */
        STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
                       list_pointers) {

            _store_image_type(self, object_props->image_type);

            /* Check for images with alt-text. */
            if (object_props->description)
                self->has_embedded_image_descriptions = LXLSX_TRUE;

            /* Check for duplicate images and only store the first instance. */
            if (object_props->md5) {
                tmp_image_md5.md5 = object_props->md5;
                found_duplicate_image = RB_FIND(lxlsx_image_md5s,
                                                self->embedded_image_md5s,
                                                &tmp_image_md5);
            }

            if (found_duplicate_image) {
                ref_id = found_duplicate_image->id;
                object_props->is_duplicate = LXLSX_TRUE;
            }
            else {
                image_ref_id++;
                ref_id = image_ref_id;
                self->num_embedded_images++;

#ifndef USE_NO_MD5
                new_image_md5 = calloc(1, sizeof(lxlsx_image_md5));
#endif
                if (new_image_md5 && object_props->md5) {
                    new_image_md5->id = ref_id;
                    new_image_md5->md5 = lxlsx_strdup(object_props->md5);

                    RB_INSERT(lxlsx_image_md5s, self->embedded_image_md5s,
                              new_image_md5);
                }
            }

            lxlsx_worksheet_set_error_cell(worksheet, object_props, ref_id);
        }

        /* Prepare background images. */
        if (worksheet->has_background_image) {

            object_props = worksheet->background_image;

            _store_image_type(self, object_props->image_type);

            /* Check for duplicate images and only store the first instance. */
            if (object_props->md5) {
                tmp_image_md5.md5 = object_props->md5;
                found_duplicate_image = RB_FIND(lxlsx_image_md5s,
                                                self->background_md5s,
                                                &tmp_image_md5);
            }

            if (found_duplicate_image) {
                ref_id = found_duplicate_image->id;
                object_props->is_duplicate = LXLSX_TRUE;
            }
            else {
                image_ref_id++;
                ref_id = image_ref_id;

#ifndef USE_NO_MD5
                new_image_md5 = calloc(1, sizeof(lxlsx_image_md5));
#endif
                if (new_image_md5 && object_props->md5) {
                    new_image_md5->id = ref_id;
                    new_image_md5->md5 = lxlsx_strdup(object_props->md5);

                    RB_INSERT(lxlsx_image_md5s, self->background_md5s,
                              new_image_md5);
                }
            }

            lxlsx_worksheet_prepare_background(worksheet, ref_id, object_props);
        }

        /* Prepare worksheet images. */
        STAILQ_FOREACH(object_props, worksheet->image_props, list_pointers) {

            /* Ignore background image added above. */
            if (object_props->is_background)
                continue;

            _store_image_type(self, object_props->image_type);

            /* Check for duplicate images and only store the first instance. */
            if (object_props->md5) {
                tmp_image_md5.md5 = object_props->md5;
                found_duplicate_image = RB_FIND(lxlsx_image_md5s,
                                                self->image_md5s,
                                                &tmp_image_md5);
            }

            if (found_duplicate_image) {
                ref_id = found_duplicate_image->id;
                object_props->is_duplicate = LXLSX_TRUE;
            }
            else {
                image_ref_id++;
                ref_id = image_ref_id;

#ifndef USE_NO_MD5
                new_image_md5 = calloc(1, sizeof(lxlsx_image_md5));
#endif
                if (new_image_md5 && object_props->md5) {
                    new_image_md5->id = ref_id;
                    new_image_md5->md5 = lxlsx_strdup(object_props->md5);

                    RB_INSERT(lxlsx_image_md5s, self->image_md5s,
                              new_image_md5);
                }
            }

            lxlsx_worksheet_prepare_image(worksheet, ref_id, lxlsx_drawing_id,
                                        object_props);
        }

        /* Prepare worksheet charts. */
        STAILQ_FOREACH(object_props, worksheet->lxlsx_chart_data, list_pointers) {
            lxlsx_chart_ref_id++;
            lxlsx_worksheet_prepare_chart(worksheet, lxlsx_chart_ref_id, lxlsx_drawing_id,
                                        object_props, is_chartsheet);
            if (object_props->chart)
                STAILQ_INSERT_TAIL(self->ordered_charts, object_props->chart,
                                   ordered_list_pointers);
        }

        /* Prepare worksheet header/footer images. */
        for (i = 0; i < LXLSX_HEADER_FOOTER_OBJS_MAX; i++) {

            object_props = *worksheet->header_footer_objs[i];
            if (!object_props)
                continue;

            _store_image_type(self, object_props->image_type);

            /* Check for duplicate images and only store the first instance. */
            if (object_props->md5) {
                tmp_image_md5.md5 = object_props->md5;
                found_duplicate_image = RB_FIND(lxlsx_image_md5s,
                                                self->header_image_md5s,
                                                &tmp_image_md5);
            }

            if (found_duplicate_image) {
                ref_id = found_duplicate_image->id;
                object_props->is_duplicate = LXLSX_TRUE;
            }
            else {
                image_ref_id++;
                ref_id = image_ref_id;

#ifndef USE_NO_MD5
                new_image_md5 = calloc(1, sizeof(lxlsx_image_md5));
#endif
                if (new_image_md5 && object_props->md5) {
                    new_image_md5->id = ref_id;
                    new_image_md5->md5 = lxlsx_strdup(object_props->md5);

                    RB_INSERT(lxlsx_image_md5s, self->header_image_md5s,
                              new_image_md5);
                }
            }

            lxlsx_worksheet_prepare_header_image(worksheet, ref_id,
                                               object_props);
        }

    }

    self->lxlsx_drawing_count = lxlsx_drawing_id;
}

/*
 * Iterate through the worksheets and set up the VML objects.
 */
STATIC void
_prepare_vml(lxlsx_workbook *self)
{
    lxlsx_worksheet *worksheet;
    lxlsx_sheet *sheet;
    uint32_t comment_id = 0;
    uint32_t lxlsx_vml_drawing_id = 0;
    uint32_t lxlsx_vml_data_id = 1;
    uint32_t lxlsx_vml_header_id = 0;
    uint32_t lxlsx_vml_shape_id = 1024;
    uint32_t comment_count = 0;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (!worksheet->has_vml && !worksheet->has_header_vml)
            continue;

        if (worksheet->has_vml) {
            self->has_vml = LXLSX_TRUE;
            if (worksheet->has_comments) {
                self->comment_count++;
                comment_id++;
                self->has_comments = LXLSX_TRUE;
            }

            lxlsx_vml_drawing_id++;

            comment_count = lxlsx_worksheet_prepare_vml_objects(worksheet,
                                                              lxlsx_vml_data_id,
                                                              lxlsx_vml_shape_id,
                                                              lxlsx_vml_drawing_id,
                                                              comment_id);

            /* Each VML should start with a shape id incremented by 1024. */
            lxlsx_vml_data_id += 1 * ((1024 + comment_count) / 1024);
            lxlsx_vml_shape_id += 1024 * ((1024 + comment_count) / 1024);
        }

        if (worksheet->has_header_vml) {
            self->has_vml = LXLSX_TRUE;
            lxlsx_vml_drawing_id++;
            lxlsx_vml_header_id++;
            lxlsx_worksheet_prepare_header_vml_objects(worksheet,
                                                     lxlsx_vml_header_id,
                                                     lxlsx_vml_drawing_id);
        }
    }
}

/*
 * Iterate through the worksheets and store any defined names used for print
 * ranges or repeat rows/columns.
 */
STATIC void
_prepare_defined_names(lxlsx_workbook *self)
{
    lxlsx_worksheet *worksheet;
    lxlsx_sheet *sheet;
    char lxlsx_app_name[LXLSX_DEFINED_NAME_LENGTH];
    char range[LXLSX_DEFINED_NAME_LENGTH];
    char area[LXLSX_MAX_CELL_RANGE_LENGTH];
    char first_col[8];
    char last_col[8];

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;
        /*
         * Check for autofilter settings and store them.
         */
        if (worksheet->autofilter.in_use) {

            lxlsx_snprintf(lxlsx_app_name, LXLSX_DEFINED_NAME_LENGTH,
                         "%s!_FilterDatabase", worksheet->quoted_name);

            lxlsx_rowcol_to_range_abs(area,
                                    worksheet->autofilter.first_row,
                                    worksheet->autofilter.first_col,
                                    worksheet->autofilter.last_row,
                                    worksheet->autofilter.last_col);

            lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH, "%s!%s",
                         worksheet->quoted_name, area);

            /* Autofilters are the only defined name to set the hidden flag. */
            _store_defined_name(self, "_xlnm._FilterDatabase", lxlsx_app_name,
                                range, worksheet->index, LXLSX_TRUE);
        }

        /*
         * Check for Print Area settings and store them.
         */
        if (worksheet->print_area.in_use) {

            lxlsx_snprintf(lxlsx_app_name, LXLSX_DEFINED_NAME_LENGTH,
                         "%s!Print_Area", worksheet->quoted_name);

            /* Check for print area that is the max row range. */
            if (worksheet->print_area.first_row == 0
                && worksheet->print_area.last_row == LXLSX_ROW_MAX - 1) {

                lxlsx_col_to_name(first_col,
                                worksheet->print_area.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->print_area.last_col, LXLSX_FALSE);

                lxlsx_snprintf(area, LXLSX_MAX_CELL_RANGE_LENGTH - 1, "$%s:$%s",
                             first_col, last_col);

            }
            /* Check for print area that is the max column range. */
            else if (worksheet->print_area.first_col == 0
                     && worksheet->print_area.last_col == LXLSX_COL_MAX - 1) {

                lxlsx_snprintf(area, LXLSX_MAX_CELL_RANGE_LENGTH - 1, "$%d:$%d",
                             worksheet->print_area.first_row + 1,
                             worksheet->print_area.last_row + 1);

            }
            else {
                lxlsx_rowcol_to_range_abs(area,
                                        worksheet->print_area.first_row,
                                        worksheet->print_area.first_col,
                                        worksheet->print_area.last_row,
                                        worksheet->print_area.last_col);
            }

            lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH, "%s!%s",
                         worksheet->quoted_name, area);

            _store_defined_name(self, "_xlnm.Print_Area", lxlsx_app_name,
                                range, worksheet->index, LXLSX_FALSE);
        }

        /*
         * Check for repeat rows/cols. aka, Print Titles and store them.
         */
        if (worksheet->repeat_rows.in_use || worksheet->repeat_cols.in_use) {
            if (worksheet->repeat_rows.in_use
                && worksheet->repeat_cols.in_use) {
                lxlsx_snprintf(lxlsx_app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_col_to_name(first_col,
                                worksheet->repeat_cols.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->repeat_cols.last_col, LXLSX_FALSE);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%s:$%s,%s!$%d:$%d",
                             worksheet->quoted_name, first_col,
                             last_col, worksheet->quoted_name,
                             worksheet->repeat_rows.first_row + 1,
                             worksheet->repeat_rows.last_row + 1);

                _store_defined_name(self, "_xlnm.Print_Titles", lxlsx_app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
            else if (worksheet->repeat_rows.in_use) {

                lxlsx_snprintf(lxlsx_app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%d:$%d", worksheet->quoted_name,
                             worksheet->repeat_rows.first_row + 1,
                             worksheet->repeat_rows.last_row + 1);

                _store_defined_name(self, "_xlnm.Print_Titles", lxlsx_app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
            else if (worksheet->repeat_cols.in_use) {
                lxlsx_snprintf(lxlsx_app_name, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!Print_Titles", worksheet->quoted_name);

                lxlsx_col_to_name(first_col,
                                worksheet->repeat_cols.first_col, LXLSX_FALSE);

                lxlsx_col_to_name(last_col,
                                worksheet->repeat_cols.last_col, LXLSX_FALSE);

                lxlsx_snprintf(range, LXLSX_DEFINED_NAME_LENGTH,
                             "%s!$%s:$%s", worksheet->quoted_name,
                             first_col, last_col);

                _store_defined_name(self, "_xlnm.Print_Titles", lxlsx_app_name,
                                    range, worksheet->index, LXLSX_FALSE);
            }
        }
    }
}

/*
 * Iterate through the worksheets and set up the table objects.
 */
STATIC void
_prepare_tables(lxlsx_workbook *self)
{
    lxlsx_worksheet *worksheet;
    lxlsx_sheet *sheet;
    uint32_t lxlsx_table_id = 0;
    uint32_t lxlsx_table_count = 0;

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        lxlsx_table_count = worksheet->lxlsx_table_count;

        if (lxlsx_table_count == 0)
            continue;

        lxlsx_worksheet_prepare_tables(worksheet, lxlsx_table_id + 1);

        lxlsx_table_id += lxlsx_table_count;
    }
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
_workbook_xml_declaration(lxlsx_workbook *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <workbook> element.
 */
STATIC void
_write_workbook(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = "http://schemas.openxmlformats.org"
        "/spreadsheetml/2006/main";
    char xmlns_r[] = "http://schemas.openxmlformats.org"
        "/officeDocument/2006/relationships";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxlsx_xml_start_tag(self->file, "workbook", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <fileVersion> element.
 */
STATIC void
_write_file_version(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("appName", "xl");
    LXLSX_PUSH_ATTRIBUTES_STR("lastEdited", "4");
    LXLSX_PUSH_ATTRIBUTES_STR("lowestEdited", "4");
    LXLSX_PUSH_ATTRIBUTES_STR("rupBuild", "4505");

    if (self->vba_project)
        LXLSX_PUSH_ATTRIBUTES_STR("codeName",
                                "{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}");

    lxlsx_xml_empty_tag(self->file, "fileVersion", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <fileSharing> element.
 */
STATIC void
_workbook_write_file_sharing(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (self->read_only == 0)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("readOnlyRecommended", "1");

    lxlsx_xml_empty_tag(self->file, "fileSharing", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <workbookPr> element.
 */
STATIC void
_write_workbook_pr(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (self->vba_codename)
        LXLSX_PUSH_ATTRIBUTES_STR("codeName", self->vba_codename);

    if (self->use_1904_epoch)
        LXLSX_PUSH_ATTRIBUTES_STR("date1904", "1");

    LXLSX_PUSH_ATTRIBUTES_STR("defaultThemeVersion", "124226");

    lxlsx_xml_empty_tag(self->file, "workbookPr", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <workbookView> element.
 */
STATIC void
_write_workbook_view(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xWindow", "240");
    LXLSX_PUSH_ATTRIBUTES_STR("yWindow", "15");
    LXLSX_PUSH_ATTRIBUTES_INT("windowWidth", self->window_width);
    LXLSX_PUSH_ATTRIBUTES_INT("windowHeight", self->window_height);

    if (self->first_sheet)
        LXLSX_PUSH_ATTRIBUTES_INT("firstSheet", self->first_sheet);

    if (self->active_sheet)
        LXLSX_PUSH_ATTRIBUTES_INT("activeTab", self->active_sheet);

    lxlsx_xml_empty_tag(self->file, "workbookView", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <bookViews> element.
 */
STATIC void
_write_book_views(lxlsx_workbook *self)
{
    lxlsx_xml_start_tag(self->file, "bookViews", NULL);

    _write_workbook_view(self);

    lxlsx_xml_end_tag(self->file, "bookViews");
}

/*
 * Write the <sheet> element.
 */
STATIC void
_write_sheet(lxlsx_workbook *self, const char *name, uint32_t sheet_id,
             uint8_t hidden)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH] = "rId1";

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", sheet_id);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("name", name);
    LXLSX_PUSH_ATTRIBUTES_INT("sheetId", sheet_id);

    if (hidden)
        LXLSX_PUSH_ATTRIBUTES_STR("state", "hidden");

    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "sheet", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sheets> element.
 */
STATIC void
_write_sheets(lxlsx_workbook *self)
{
    lxlsx_sheet *sheet;
    lxlsx_worksheet *worksheet;
    lxlsx_chartsheet *chartsheet;

    lxlsx_xml_start_tag(self->file, "sheets", NULL);

    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet) {
            chartsheet = sheet->u.chartsheet;
            _write_sheet(self, chartsheet->name, chartsheet->index + 1,
                         chartsheet->hidden);
        }
        else {
            worksheet = sheet->u.worksheet;
            _write_sheet(self, worksheet->name, worksheet->index + 1,
                         worksheet->hidden);
        }
    }

    lxlsx_xml_end_tag(self->file, "sheets");
}

/*
 * Write the <calcPr> element.
 */
STATIC void
_write_calc_pr(lxlsx_workbook *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("calcId", "124519");
    LXLSX_PUSH_ATTRIBUTES_STR("fullCalcOnLoad", "1");

    lxlsx_xml_empty_tag(self->file, "calcPr", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <definedName> element.
 */
STATIC void
_write_defined_name(lxlsx_workbook *self, lxlsx_defined_name *defined_name)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("name", defined_name->name);

    if (defined_name->index != -1)
        LXLSX_PUSH_ATTRIBUTES_INT("localSheetId", defined_name->index);

    if (defined_name->hidden)
        LXLSX_PUSH_ATTRIBUTES_INT("hidden", 1);

    lxlsx_xml_data_element(self->file, "definedName", defined_name->formula,
                         &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

STATIC void
_write_defined_names(lxlsx_workbook *self)
{
    lxlsx_defined_name *defined_name;

    if (TAILQ_EMPTY(self->defined_names))
        return;

    lxlsx_xml_start_tag(self->file, "definedNames", NULL);

    TAILQ_FOREACH(defined_name, self->defined_names, list_pointers) {
        _write_defined_name(self, defined_name);
    }

    lxlsx_xml_end_tag(self->file, "definedNames");
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
lxlsx_workbook_assemble_xml_file(lxlsx_workbook *self)
{
    /* Prepare workbook and sub-objects for writing. */
    _prepare_workbook(self);

    /* Write the XML declaration. */
    _workbook_xml_declaration(self);

    /* Write the root workbook element. */
    _write_workbook(self);

    /* Write the XLSX file version. */
    _write_file_version(self);

    /* Write the fileSharing element. */
    _workbook_write_file_sharing(self);

    /* Write the workbook properties. */
    _write_workbook_pr(self);

    /* Write the workbook view properties. */
    _write_book_views(self);

    /* Write the worksheet names and ids. */
    _write_sheets(self);

    /* Write the workbook defined names. */
    _write_defined_names(self);

    /* Write the workbook calculation properties. */
    _write_calc_pr(self);

    /* Close the workbook tag. */
    lxlsx_xml_end_tag(self->file, "workbook");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Create a new workbook object.
 */
lxlsx_workbook *
lxlsx_workbook_new(const char *filename)
{
    return lxlsx_workbook_new_opt(filename, NULL);
}

/* Deprecated function name for backwards compatibility. */
lxlsx_workbook *
new_workbook(const char *filename)
{
    return lxlsx_workbook_new_opt(filename, NULL);
}

/* Deprecated function name for backwards compatibility. */
lxlsx_workbook *
new_workbook_opt(const char *filename, lxlsx_workbook_options *options)
{
    return lxlsx_workbook_new_opt(filename, options);
}

/*
 * Create a new workbook object with options.
 */
lxlsx_workbook *
lxlsx_workbook_new_opt(const char *filename, lxlsx_workbook_options *options)
{
    lxlsx_format *format;
    lxlsx_workbook *workbook;

    /* Create the workbook object. */
    workbook = calloc(1, sizeof(lxlsx_workbook));
    GOTO_LABEL_ON_MEM_ERROR(workbook, mem_error);
    workbook->filename = lxlsx_strdup(filename);

    /* Add the sheets list. */
    workbook->sheets = calloc(1, sizeof(struct lxlsx_sheets));
    GOTO_LABEL_ON_MEM_ERROR(workbook->sheets, mem_error);
    STAILQ_INIT(workbook->sheets);

    /* Add the worksheets list. */
    workbook->worksheets = calloc(1, sizeof(struct lxlsx_worksheets));
    GOTO_LABEL_ON_MEM_ERROR(workbook->worksheets, mem_error);
    STAILQ_INIT(workbook->worksheets);

    /* Add the chartsheets list. */
    workbook->chartsheets = calloc(1, sizeof(struct lxlsx_chartsheets));
    GOTO_LABEL_ON_MEM_ERROR(workbook->chartsheets, mem_error);
    STAILQ_INIT(workbook->chartsheets);

    /* Add the worksheet names tree. */
    workbook->lxlsx_worksheet_names = calloc(1, sizeof(struct lxlsx_worksheet_names));
    GOTO_LABEL_ON_MEM_ERROR(workbook->lxlsx_worksheet_names, mem_error);
    RB_INIT(workbook->lxlsx_worksheet_names);

    /* Add the chartsheet names tree. */
    workbook->lxlsx_chartsheet_names = calloc(1,
                                        sizeof(struct lxlsx_chartsheet_names));
    GOTO_LABEL_ON_MEM_ERROR(workbook->lxlsx_chartsheet_names, mem_error);
    RB_INIT(workbook->lxlsx_chartsheet_names);

    /* Add the image MD5 tree. */
    workbook->image_md5s = calloc(1, sizeof(struct lxlsx_image_md5s));
    GOTO_LABEL_ON_MEM_ERROR(workbook->image_md5s, mem_error);
    RB_INIT(workbook->image_md5s);

    /* Add the embedded image MD5 tree. */
    workbook->embedded_image_md5s = calloc(1, sizeof(struct lxlsx_image_md5s));
    GOTO_LABEL_ON_MEM_ERROR(workbook->embedded_image_md5s, mem_error);
    RB_INIT(workbook->embedded_image_md5s);

    /* Add the header image MD5 tree. */
    workbook->header_image_md5s = calloc(1, sizeof(struct lxlsx_image_md5s));
    GOTO_LABEL_ON_MEM_ERROR(workbook->header_image_md5s, mem_error);
    RB_INIT(workbook->header_image_md5s);

    /* Add the background image MD5 tree. */
    workbook->background_md5s = calloc(1, sizeof(struct lxlsx_image_md5s));
    GOTO_LABEL_ON_MEM_ERROR(workbook->background_md5s, mem_error);
    RB_INIT(workbook->background_md5s);

    /* Add the charts list. */
    workbook->charts = calloc(1, sizeof(struct lxlsx_charts));
    GOTO_LABEL_ON_MEM_ERROR(workbook->charts, mem_error);
    STAILQ_INIT(workbook->charts);

    /* Add the ordered charts list to track chart insertion order. */
    workbook->ordered_charts = calloc(1, sizeof(struct lxlsx_charts));
    GOTO_LABEL_ON_MEM_ERROR(workbook->ordered_charts, mem_error);
    STAILQ_INIT(workbook->ordered_charts);

    /* Add the formats list. */
    workbook->formats = calloc(1, sizeof(struct lxlsx_formats));
    GOTO_LABEL_ON_MEM_ERROR(workbook->formats, mem_error);
    STAILQ_INIT(workbook->formats);

    /* Add the defined_names list. */
    workbook->defined_names = calloc(1, sizeof(struct lxlsx_defined_names));
    GOTO_LABEL_ON_MEM_ERROR(workbook->defined_names, mem_error);
    TAILQ_INIT(workbook->defined_names);

    /* Add the shared strings table. */
    workbook->sst = lxlsx_sst_new();
    GOTO_LABEL_ON_MEM_ERROR(workbook->sst, mem_error);

    /* Add the default workbook properties. */
    workbook->properties = calloc(1, sizeof(lxlsx_doc_properties));
    GOTO_LABEL_ON_MEM_ERROR(workbook->properties, mem_error);

    /* Add a hash table to track format indices. */
    workbook->used_xf_formats = lxlsx_hash_new(128, 1, 0);
    GOTO_LABEL_ON_MEM_ERROR(workbook->used_xf_formats, mem_error);

    /* Add a hash table to track format indices. */
    workbook->used_dxf_formats = lxlsx_hash_new(128, 1, 0);
    GOTO_LABEL_ON_MEM_ERROR(workbook->used_dxf_formats, mem_error);

    /* Add the worksheets list. */
    workbook->lxlsx_custom_properties =
        calloc(1, sizeof(struct lxlsx_custom_properties));
    GOTO_LABEL_ON_MEM_ERROR(workbook->lxlsx_custom_properties, mem_error);
    STAILQ_INIT(workbook->lxlsx_custom_properties);

    /* Add the default cell format. */
    format = lxlsx_workbook_add_format(workbook);
    GOTO_LABEL_ON_MEM_ERROR(format, mem_error);

    /* Initialize its index. */
    lxlsx_format_get_xf_index(format);

    /* Add the default hyperlink format. */
    format = lxlsx_workbook_add_format(workbook);
    GOTO_LABEL_ON_MEM_ERROR(format, mem_error);
    lxlsx_format_set_hyperlink(format);
    workbook->default_url_format = format;

    if (options) {
        workbook->options.constant_memory = options->constant_memory;
        workbook->options.tmpdir = lxlsx_strdup(options->tmpdir);
        workbook->options.use_zip64 = options->use_zip64;
        workbook->options.output_buffer = options->output_buffer;
        workbook->options.output_buffer_size = options->output_buffer_size;
    }

    workbook->max_url_length = 2079;
    workbook->window_width = 16095;
    workbook->window_height = 9660;

    return workbook;

mem_error:
    lxlsx_workbook_free(workbook);
    workbook = NULL;
    return NULL;
}

/*
 * Add a new worksheet to the Excel workbook.
 */
lxlsx_worksheet *
lxlsx_workbook_add_worksheet(lxlsx_workbook *self, const char *sheetname)
{
    lxlsx_sheet *sheet = NULL;
    lxlsx_worksheet *worksheet = NULL;
    lxlsx_worksheet_name *lxlsx_worksheet_name = NULL;
    lxlsx_error error;
    lxlsx_worksheet_init_data init_data =
        { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
    char *new_name = NULL;

    if (sheetname) {
        /* Use the user supplied name. */
        init_data.name = lxlsx_strdup(sheetname);
        init_data.quoted_name = lxlsx_quote_sheetname(sheetname);
    }
    else {
        /* Use the default SheetN name. */
        new_name = malloc(LXLSX_MAX_SHEETNAME_LENGTH);
        GOTO_LABEL_ON_MEM_ERROR(new_name, mem_error);

        lxlsx_snprintf(new_name, LXLSX_MAX_SHEETNAME_LENGTH, "Sheet%d",
                     self->num_worksheets + 1);
        init_data.name = new_name;
        init_data.quoted_name = lxlsx_strdup(new_name);
    }

    /* Check that the worksheet name is valid. */
    error = lxlsx_workbook_validate_sheet_name(self, init_data.name);
    if (error) {
        LXLSX_WARN_FORMAT2("lxlsx_workbook_add_worksheet(): worksheet name '%s' has "
                         "error: %s", init_data.name, lxlsx_strerror(error));
        goto mem_error;
    }

    /* Create a struct to find/store the worksheet name/pointer. */
    lxlsx_worksheet_name = calloc(1, sizeof(struct lxlsx_worksheet_name));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_worksheet_name, mem_error);

    /* Initialize the metadata to pass to the worksheet. */
    init_data.hidden = 0;
    init_data.index = self->num_sheets;
    init_data.sst = self->sst;
    init_data.optimize = self->options.constant_memory;
    init_data.active_sheet = &self->active_sheet;
    init_data.first_sheet = &self->first_sheet;
    init_data.tmpdir = self->options.tmpdir;
    init_data.default_url_format = self->default_url_format;
    init_data.max_url_length = self->max_url_length;
    init_data.use_1904_epoch = self->use_1904_epoch;

    /* Create a new worksheet object. */
    worksheet = lxlsx_worksheet_new(&init_data);
    GOTO_LABEL_ON_MEM_ERROR(worksheet, mem_error);

    /* Add it to the worksheet list. */
    self->num_worksheets++;
    STAILQ_INSERT_TAIL(self->worksheets, worksheet, list_pointers);

    /* Create a new sheet object. */
    sheet = calloc(1, sizeof(lxlsx_sheet));
    GOTO_LABEL_ON_MEM_ERROR(sheet, mem_error);
    sheet->u.worksheet = worksheet;

    /* Add it to the worksheet list. */
    self->num_sheets++;
    STAILQ_INSERT_TAIL(self->sheets, sheet, list_pointers);

    /* Store the worksheet so we can look it up by name. */
    lxlsx_worksheet_name->name = init_data.name;
    lxlsx_worksheet_name->worksheet = worksheet;
    RB_INSERT(lxlsx_worksheet_names, self->lxlsx_worksheet_names, lxlsx_worksheet_name);

    return worksheet;

mem_error:
    free((void *) init_data.name);
    free((void *) init_data.quoted_name);
    free(lxlsx_worksheet_name);
    free(worksheet);
    return NULL;
}

/*
 * Add a new chartsheet to the Excel workbook.
 */
lxlsx_chartsheet *
lxlsx_workbook_add_chartsheet(lxlsx_workbook *self, const char *sheetname)
{
    lxlsx_sheet *sheet = NULL;
    lxlsx_chartsheet *chartsheet = NULL;
    lxlsx_chartsheet_name *lxlsx_chartsheet_name = NULL;
    lxlsx_error error;
    lxlsx_worksheet_init_data init_data =
        { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
    char *new_name = NULL;

    if (sheetname) {
        /* Use the user supplied name. */
        init_data.name = lxlsx_strdup(sheetname);
        init_data.quoted_name = lxlsx_quote_sheetname(sheetname);
    }
    else {
        /* Use the default SheetN name. */
        new_name = malloc(LXLSX_MAX_SHEETNAME_LENGTH);
        GOTO_LABEL_ON_MEM_ERROR(new_name, mem_error);

        lxlsx_snprintf(new_name, LXLSX_MAX_SHEETNAME_LENGTH, "Chart%d",
                     self->num_chartsheets + 1);
        init_data.name = new_name;
        init_data.quoted_name = lxlsx_strdup(new_name);
    }

    /* Check that the chartsheet name is valid. */
    error = lxlsx_workbook_validate_sheet_name(self, init_data.name);
    if (error) {
        LXLSX_WARN_FORMAT2
            ("lxlsx_workbook_add_chartsheet(): chartsheet name '%s' has "
             "error: %s", init_data.name, lxlsx_strerror(error));
        goto mem_error;
    }

    /* Create a struct to find/store the chartsheet name/pointer. */
    lxlsx_chartsheet_name = calloc(1, sizeof(struct lxlsx_chartsheet_name));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_chartsheet_name, mem_error);

    /* Initialize the metadata to pass to the chartsheet. */
    init_data.hidden = 0;
    init_data.index = self->num_sheets;
    init_data.sst = self->sst;
    init_data.optimize = self->options.constant_memory;
    init_data.active_sheet = &self->active_sheet;
    init_data.first_sheet = &self->first_sheet;
    init_data.tmpdir = self->options.tmpdir;

    /* Create a new chartsheet object. */
    chartsheet = lxlsx_chartsheet_new(&init_data);
    GOTO_LABEL_ON_MEM_ERROR(chartsheet, mem_error);

    /* Add it to the chartsheet list. */
    self->num_chartsheets++;
    STAILQ_INSERT_TAIL(self->chartsheets, chartsheet, list_pointers);

    /* Create a new sheet object. */
    sheet = calloc(1, sizeof(lxlsx_sheet));
    GOTO_LABEL_ON_MEM_ERROR(sheet, mem_error);
    sheet->is_chartsheet = LXLSX_TRUE;
    sheet->u.chartsheet = chartsheet;

    /* Add it to the chartsheet list. */
    self->num_sheets++;
    STAILQ_INSERT_TAIL(self->sheets, sheet, list_pointers);

    /* Store the chartsheet so we can look it up by name. */
    lxlsx_chartsheet_name->name = init_data.name;
    lxlsx_chartsheet_name->chartsheet = chartsheet;
    RB_INSERT(lxlsx_chartsheet_names, self->lxlsx_chartsheet_names, lxlsx_chartsheet_name);

    return chartsheet;

mem_error:
    free((void *) init_data.name);
    free((void *) init_data.quoted_name);
    free(lxlsx_chartsheet_name);
    free(chartsheet);
    return NULL;
}

/*
 * Add a new chart to the Excel workbook.
 */
lxlsx_chart *
lxlsx_workbook_add_chart(lxlsx_workbook *self, uint8_t type)
{
    lxlsx_chart *chart;

    if (type == LXLSX_CHART_NONE || type > LXLSX_CHART_RADAR_FILLED) {
        LXLSX_WARN_FORMAT1("lxlsx_workbook_add_chart(): invalid chart type: %d",
                         type);
        return NULL;
    }

    /* Create a new chart object. */
    chart = lxlsx_chart_new(type);

    if (chart)
        STAILQ_INSERT_TAIL(self->charts, chart, list_pointers);

    return chart;
}

/*
 * Add a new format to the Excel workbook.
 */
lxlsx_format *
lxlsx_workbook_add_format(lxlsx_workbook *self)
{
    /* Create a new format object. */
    lxlsx_format *format = lxlsx_format_new();
    RETURN_ON_MEM_ERROR(format, NULL);

    format->xf_format_indices = self->used_xf_formats;
    format->dxf_format_indices = self->used_dxf_formats;
    format->num_xf_formats = &self->num_xf_formats;

    STAILQ_INSERT_TAIL(self->formats, format, list_pointers);

    return format;
}

/*
 * Call finalization code and close file.
 */
lxlsx_error
lxlsx_workbook_close(lxlsx_workbook *self)
{
    lxlsx_sheet *sheet = NULL;
    lxlsx_worksheet *worksheet = NULL;
    lxlsx_packager *packager = NULL;
    lxlsx_error error = LXLSX_NO_ERROR;
    char codename[LXLSX_MAX_SHEETNAME_LENGTH] = { 0 };

    /* Add a default worksheet if non have been added. */
    if (!self->num_sheets)
        lxlsx_workbook_add_worksheet(self, NULL);

    /* Ensure that at least one worksheet has been selected. */
    if (self->active_sheet == 0) {
        sheet = STAILQ_FIRST(self->sheets);
        if (!sheet->is_chartsheet) {
            worksheet = sheet->u.worksheet;
            worksheet->selected = LXLSX_TRUE;
            worksheet->hidden = 0;
        }
    }

    /* Set the active sheet and check if a metadata file is needed. */
    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
        if (sheet->is_chartsheet)
            continue;
        else
            worksheet = sheet->u.worksheet;

        if (worksheet->index == self->active_sheet)
            worksheet->active = LXLSX_TRUE;

        if (worksheet->has_dynamic_functions) {
            self->has_metadata = LXLSX_TRUE;
            self->has_dynamic_functions = LXLSX_TRUE;

        }

        if (!STAILQ_EMPTY(worksheet->embedded_image_props)) {
            self->has_metadata = LXLSX_TRUE;
            self->has_embedded_images = LXLSX_TRUE;
        }
    }

    /* Set workbook and worksheet VBA codenames if a macro has been added. */
    if (self->vba_project) {
        if (!self->vba_codename)
            lxlsx_workbook_set_vba_name(self, "ThisWorkbook");

        STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
            if (sheet->is_chartsheet)
                continue;
            else
                worksheet = sheet->u.worksheet;

            if (!worksheet->vba_codename) {
                lxlsx_snprintf(codename, LXLSX_MAX_SHEETNAME_LENGTH, "Sheet%d",
                             worksheet->index + 1);

                lxlsx_worksheet_set_vba_name(worksheet, codename);
            }
        }
    }

    /* Prepare the worksheet VML elements such as comments. */
    _prepare_vml(self);

    /* Set the defined names for the worksheets such as Print Titles. */
    _prepare_defined_names(self);

    /* Prepare the drawings, charts and images. */
    _prepare_drawings(self);

    /* Add cached data to charts. */
    _add_chart_cache_data(self);

    /* Set the table ids for the worksheet tables. */
    _prepare_tables(self);

    /* Create a packager object to assemble sub-elements into a zip file. */
    packager = lxlsx_packager_new(self->filename,
                                self->options.tmpdir,
                                self->options.use_zip64);

    /* If the packager fails it is generally due to a zip permission error. */
    if (packager == NULL) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Error creating '%s'. "
                   "System error = %s\n", self->filename, strerror(errno));

        error = LXLSX_ERROR_CREATING_XLSX_FILE;
        goto mem_error;
    }

    /* Set the workbook object in the packager. */
    packager->workbook = self;

    /* Assemble all the sub-files in the xlsx package. */
    error = lxlsx_create_package(packager);

    if (!self->filename) {
        *self->options.output_buffer = packager->output_buffer;
        *self->options.output_buffer_size = packager->output_buffer_size;
    }

    /* Error and non-error conditions fall through to the cleanup code. */
    if (error == LXLSX_ERROR_CREATING_TMPFILE) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Error creating tmpfile(s) to assemble '%s'. "
                   "System error = %s\n", self->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_FILE_OPERATION then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_FILE_OPERATION) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip ZIP_ERRNO error while creating xlsx file '%s'. "
                   "System error = %s\n", self->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_PARAMETER_ERROR then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_PARAMETER_ERROR) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip ZIP_PARAMERROR error while creating xlsx file '%s'. "
                   "System error = %s\n", self->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_BAD_ZIP_FILE then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_BAD_ZIP_FILE) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip ZIP_BADZIPFILE error while creating xlsx file '%s'. "
                   "This may require the use_zip64 option for large files. "
                   "System error = %s\n", self->filename, strerror(errno));
    }

    /* If LXLSX_ERROR_ZIP_INTERNAL_ERROR then errno is set by zip. */
    if (error == LXLSX_ERROR_ZIP_INTERNAL_ERROR) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip ZIP_INTERNALERROR error while creating xlsx file '%s'. "
                   "System error = %s\n", self->filename, strerror(errno));
    }

    /* The next 2 error conditions don't set errno. */
    if (error == LXLSX_ERROR_ZIP_FILE_ADD) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip error adding file to xlsx file '%s'.\n",
                   self->filename);
    }

    if (error == LXLSX_ERROR_ZIP_CLOSE) {
        LXLSX_PRINTF(LXLSX_STDERR "[ERROR] lxlsx_workbook_close(): "
                   "Zip error closing xlsx file '%s'.\n", self->filename);
    }

mem_error:
    lxlsx_packager_free(packager);
    lxlsx_workbook_free(self);
    return error;
}

/*
 * Create a defined name in Excel. We handle global/workbook level names and
 * local/worksheet names.
 */
lxlsx_error
lxlsx_workbook_define_name(lxlsx_workbook *self, const char *name,
                     const char *formula)
{
    return _store_defined_name(self, name, NULL, formula, -1, LXLSX_FALSE);
}

/*
 * Set the document properties such as Title, Author etc.
 */
lxlsx_error
lxlsx_workbook_set_properties(lxlsx_workbook *self, lxlsx_doc_properties *user_props)
{
    lxlsx_doc_properties *doc_props;

    /* Free any existing properties. */
    _free_doc_properties(self->properties);

    doc_props = calloc(1, sizeof(lxlsx_doc_properties));
    GOTO_LABEL_ON_MEM_ERROR(doc_props, mem_error);

    /* Copy the user properties to an internal structure. */
    if (user_props->title) {
        doc_props->title = lxlsx_strdup(user_props->title);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->title, mem_error);
    }

    if (user_props->subject) {
        doc_props->subject = lxlsx_strdup(user_props->subject);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->subject, mem_error);
    }

    if (user_props->author) {
        doc_props->author = lxlsx_strdup(user_props->author);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->author, mem_error);
    }

    if (user_props->manager) {
        doc_props->manager = lxlsx_strdup(user_props->manager);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->manager, mem_error);
    }

    if (user_props->company) {
        doc_props->company = lxlsx_strdup(user_props->company);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->company, mem_error);
    }

    if (user_props->category) {
        doc_props->category = lxlsx_strdup(user_props->category);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->category, mem_error);
    }

    if (user_props->keywords) {
        doc_props->keywords = lxlsx_strdup(user_props->keywords);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->keywords, mem_error);
    }

    if (user_props->comments) {
        doc_props->comments = lxlsx_strdup(user_props->comments);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->comments, mem_error);
    }

    if (user_props->status) {
        doc_props->status = lxlsx_strdup(user_props->status);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->status, mem_error);
    }

    if (user_props->hyperlink_base) {
        doc_props->hyperlink_base = lxlsx_strdup(user_props->hyperlink_base);
        GOTO_LABEL_ON_MEM_ERROR(doc_props->hyperlink_base, mem_error);
    }

    doc_props->created = user_props->created;

    self->properties = doc_props;

    return LXLSX_NO_ERROR;

mem_error:
    _free_doc_properties(doc_props);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Set a string custom document property.
 */
lxlsx_error
lxlsx_workbook_set_custom_property_string(lxlsx_workbook *self, const char *name,
                                    const char *value)
{
    lxlsx_custom_property *lxlsx_custom_property;

    if (!name) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_string(): "
                        "parameter 'name' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_string(): "
                        "parameter 'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    if (!value) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_string(): "
                        "parameter 'value' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_utf8_strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_string(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    if (lxlsx_utf8_strlen(value) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_string(): parameter "
                        "'value' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    lxlsx_custom_property = calloc(1, sizeof(struct lxlsx_custom_property));
    RETURN_ON_MEM_ERROR(lxlsx_custom_property, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    lxlsx_custom_property->name = lxlsx_strdup(name);
    lxlsx_custom_property->u.string = lxlsx_strdup(value);
    lxlsx_custom_property->type = LXLSX_CUSTOM_STRING;

    STAILQ_INSERT_TAIL(self->lxlsx_custom_properties, lxlsx_custom_property,
                       list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Set a double number custom document property.
 */
lxlsx_error
lxlsx_workbook_set_custom_property_number(lxlsx_workbook *self, const char *name,
                                    double value)
{
    lxlsx_custom_property *lxlsx_custom_property;

    if (!name) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_number(): parameter "
                        "'name' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_number(): parameter "
                        "'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    if (lxlsx_utf8_strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_number(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    lxlsx_custom_property = calloc(1, sizeof(struct lxlsx_custom_property));
    RETURN_ON_MEM_ERROR(lxlsx_custom_property, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    lxlsx_custom_property->name = lxlsx_strdup(name);
    lxlsx_custom_property->u.number = value;
    lxlsx_custom_property->type = LXLSX_CUSTOM_DOUBLE;

    STAILQ_INSERT_TAIL(self->lxlsx_custom_properties, lxlsx_custom_property,
                       list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Set a integer number custom document property.
 */
lxlsx_error
lxlsx_workbook_set_custom_property_integer(lxlsx_workbook *self, const char *name,
                                     int32_t value)
{
    lxlsx_custom_property *lxlsx_custom_property;

    if (!name) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_integer(): parameter "
                        "'name' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_integer(): parameter "
                        "'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    if (strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_integer(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    lxlsx_custom_property = calloc(1, sizeof(struct lxlsx_custom_property));
    RETURN_ON_MEM_ERROR(lxlsx_custom_property, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    lxlsx_custom_property->name = lxlsx_strdup(name);
    lxlsx_custom_property->u.integer = value;
    lxlsx_custom_property->type = LXLSX_CUSTOM_INTEGER;

    STAILQ_INSERT_TAIL(self->lxlsx_custom_properties, lxlsx_custom_property,
                       list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Set a boolean custom document property.
 */
lxlsx_error
lxlsx_workbook_set_custom_property_boolean(lxlsx_workbook *self, const char *name,
                                     uint8_t value)
{
    lxlsx_custom_property *lxlsx_custom_property;

    if (!name) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_boolean(): parameter "
                        "'name' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_boolean(): parameter "
                        "'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    if (lxlsx_utf8_strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_boolean(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    lxlsx_custom_property = calloc(1, sizeof(struct lxlsx_custom_property));
    RETURN_ON_MEM_ERROR(lxlsx_custom_property, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    lxlsx_custom_property->name = lxlsx_strdup(name);
    lxlsx_custom_property->u.boolean = value;
    lxlsx_custom_property->type = LXLSX_CUSTOM_BOOLEAN;

    STAILQ_INSERT_TAIL(self->lxlsx_custom_properties, lxlsx_custom_property,
                       list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Set a datetime custom document property.
 */
lxlsx_error
lxlsx_workbook_set_custom_property_datetime(lxlsx_workbook *self, const char *name,
                                      lxlsx_datetime *datetime)
{
    lxlsx_custom_property *lxlsx_custom_property;

    if (!name) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_datetime(): parameter "
                        "'name' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_datetime(): parameter "
                        "'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    if (lxlsx_utf8_strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_datetime(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (!datetime) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_custom_property_datetime(): parameter "
                        "'datetime' cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_datetime_validate(datetime) != LXLSX_NO_ERROR) {
        return LXLSX_ERROR_DATETIME_VALIDATION;
    }

    /* Create a struct to hold the custom property. */
    lxlsx_custom_property = calloc(1, sizeof(struct lxlsx_custom_property));
    RETURN_ON_MEM_ERROR(lxlsx_custom_property, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    lxlsx_custom_property->name = lxlsx_strdup(name);

    memcpy(&lxlsx_custom_property->u.datetime, datetime, sizeof(lxlsx_datetime));
    lxlsx_custom_property->type = LXLSX_CUSTOM_DATETIME;

    STAILQ_INSERT_TAIL(self->lxlsx_custom_properties, lxlsx_custom_property,
                       list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Get a worksheet object from its name.
 */
lxlsx_worksheet *
lxlsx_workbook_get_worksheet_by_name(lxlsx_workbook *self, const char *name)
{
    lxlsx_worksheet_name lookup;
    lxlsx_worksheet_name *found;

    if (!name)
        return NULL;

    lookup.name = name;
    found = RB_FIND(lxlsx_worksheet_names,
                    self->lxlsx_worksheet_names, &lookup);

    if (found)
        return found->worksheet;
    else
        return NULL;
}

/*
 * Get a chartsheet object from its name.
 */
lxlsx_chartsheet *
lxlsx_workbook_get_chartsheet_by_name(lxlsx_workbook *self, const char *name)
{
    lxlsx_chartsheet_name lookup;
    lxlsx_chartsheet_name *found;

    if (!name)
        return NULL;

    lookup.name = name;
    found = RB_FIND(lxlsx_chartsheet_names,
                    self->lxlsx_chartsheet_names, &lookup);

    if (found)
        return found->chartsheet;
    else
        return NULL;
}

/*
 * Get the default URL format.
 */
lxlsx_format *
lxlsx_workbook_get_default_url_format(lxlsx_workbook *self)
{
    return self->default_url_format;
}

/*
 * Unset the default URL format.
 */
void
lxlsx_workbook_unset_default_url_format(lxlsx_workbook *self)
{
    self->default_url_format->hyperlink = LXLSX_FALSE;
    self->default_url_format->xf_id = 0;
    self->default_url_format->underline = LXLSX_UNDERLINE_NONE;
    self->default_url_format->theme = 0;
}

/*
 * Validate the worksheet name based on Excel's rules.
 */
lxlsx_error
lxlsx_workbook_validate_sheet_name(lxlsx_workbook *self, const char *sheetname)
{
    if (sheetname == NULL)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    /* Check for empty worksheet name. */
    if (lxlsx_str_is_empty(sheetname))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    /* Check the UTF-8 length of the worksheet name. */
    if (lxlsx_utf8_strlen(sheetname) > LXLSX_SHEETNAME_MAX)
        return LXLSX_ERROR_SHEETNAME_LENGTH_EXCEEDED;

    /* Check that the worksheet name doesn't contain invalid characters. */
    if (strpbrk(sheetname, "[]:*?/\\"))
        return LXLSX_ERROR_INVALID_SHEETNAME_CHARACTER;

    /* Check that the worksheet doesn't start or end with an apostrophe. */
    if (sheetname[0] == '\'' || sheetname[strlen(sheetname) - 1] == '\'')
        return LXLSX_ERROR_SHEETNAME_START_END_APOSTROPHE;

    /* Check if the worksheet name is already in use. */
    if (lxlsx_workbook_get_worksheet_by_name(self, sheetname))
        return LXLSX_ERROR_SHEETNAME_ALREADY_USED;

    /* Check if the chartsheet name is already in use. */
    if (lxlsx_workbook_get_chartsheet_by_name(self, sheetname))
        return LXLSX_ERROR_SHEETNAME_ALREADY_USED;

    return LXLSX_NO_ERROR;
}

/*
 * Add a vbaProject binary to the Excel workbook.
 */
lxlsx_error
lxlsx_workbook_add_vba_project(lxlsx_workbook *self, const char *filename)
{
    FILE *filehandle;

    if (!filename) {
        LXLSX_WARN("lxlsx_workbook_add_vba_project(): "
                 "project filename must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the vbaProject file exists and can be opened. */
    filehandle = lxlsx_fopen(filename, "rb");
    if (!filehandle) {
        LXLSX_WARN_FORMAT1("lxlsx_workbook_add_vba_project(): "
                         "project file doesn't exist or can't be opened: %s.",
                         filename);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    fclose(filehandle);

    self->vba_project = lxlsx_strdup(filename);

    return LXLSX_NO_ERROR;
}

/*
 * Add a vbaProject binary and a vbaProjectSignature binary to the Excel workbook.
 */
lxlsx_error
lxlsx_workbook_add_signed_vba_project(lxlsx_workbook *self,
                                const char *vba_project,
                                const char *signature)
{
    FILE *filehandle;

    lxlsx_error error = lxlsx_workbook_add_vba_project(self, vba_project);
    if (error != LXLSX_NO_ERROR)
        return error;

    if (!signature) {
        LXLSX_WARN("lxlsx_workbook_add_signed_vba_project(): "
                 "signature filename must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the vbaProjectSignature file exists and can be opened. */
    filehandle = lxlsx_fopen(signature, "rb");
    if (!filehandle) {
        LXLSX_WARN_FORMAT1("lxlsx_workbook_add_signed_vba_project(): "
                         "signature file doesn't exist or can't be opened: %s.",
                         signature);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    fclose(filehandle);

    self->vba_project_signature = lxlsx_strdup(signature);

    return LXLSX_NO_ERROR;
}

/*
 * Set the VBA name for the workbook.
 */
lxlsx_error
lxlsx_workbook_set_vba_name(lxlsx_workbook *self, const char *name)
{
    if (!name) {
        LXLSX_WARN("lxlsx_workbook_set_vba_name(): " "name must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_str_is_empty(name)) {
        LXLSX_WARN_FORMAT("lxlsx_workbook_set_vba_name(): parameter "
                        "'name' cannot be an empty string.");
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    self->vba_codename = lxlsx_strdup(name);

    return LXLSX_NO_ERROR;
}

/*
 * Set the Excel "Read-only recommended" save option.
 */
void
lxlsx_workbook_read_only_recommended(lxlsx_workbook *self)
{
    self->read_only = 2;
}

/*
 * Use the 1904 epoch for dates in the workbook.
 */
void
lxlsx_workbook_use_1904_epoch(lxlsx_workbook *self)
{
    self->use_1904_epoch = LXLSX_TRUE;
}

/*
 * Set the size of a workbook window.
 */
void
lxlsx_workbook_set_size(lxlsx_workbook *workbook, uint16_t width, uint16_t height)
{
    /* Convert the width/height to twips at 96 dpi. */
    if (width)
        workbook->window_width = width * 1440 / 96;

    if (height)
        workbook->window_height = height * 1440 / 96;

}
