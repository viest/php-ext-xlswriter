/*****************************************************************************
 * worksheet - A library for creating Excel XLSX worksheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#ifdef USE_FMEMOPEN
#define _POSIX_C_SOURCE 200809L
#endif

#include "libxlsx/xmlwriter.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/format.h"
#include "libxlsx/utility.h"

#ifdef USE_OPENSSL_MD5
#include <openssl/md5.h>
#else
#ifndef USE_NO_MD5
#include "libxlsx/third_party/md5.h"
#endif
#endif

#define LXLSX_STR_MAX                      32767
#define LXLSX_BUFFER_SIZE                  4096
#define LXLSX_PRINT_ACROSS                 1
#define LXLSX_VALIDATION_MAX_TITLE_LENGTH  32
#define LXLSX_VALIDATION_MAX_STRING_LENGTH 255
#define LXLSX_THIS_ROW "[#This Row],"
/*
 * Forward declarations.
 */
STATIC void _worksheet_write_rows(lxlsx_worksheet *self);
STATIC int _row_cmp(lxlsx_row *row1, lxlsx_row *row2);
STATIC int _cell_cmp(lxlsx_cell *cell1, lxlsx_cell *cell2);
STATIC int _drawing_rel_id_cmp(lxlsx_drawing_rel_id *tuple1,
                               lxlsx_drawing_rel_id *tuple2);
STATIC int _cond_format_hash_cmp(lxlsx_cond_format_hash_element *elem_1,
                                 lxlsx_cond_format_hash_element *elem_2);

#ifndef __clang_analyzer__
LXLSX_RB_GENERATE_ROW(lxlsx_table_rows, lxlsx_row, tree_pointers, _row_cmp);
LXLSX_RB_GENERATE_CELL(lxlsx_table_cells, lxlsx_cell, tree_pointers, _cell_cmp);
LXLSX_RB_GENERATE_DRAWING_REL_IDS(lxlsx_drawing_rel_ids, lxlsx_drawing_rel_id,
                                tree_pointers, _drawing_rel_id_cmp);
LXLSX_RB_GENERATE_VML_DRAWING_REL_IDS(lxlsx_vml_drawing_rel_ids,
                                    lxlsx_drawing_rel_id, tree_pointers,
                                    _drawing_rel_id_cmp);
LXLSX_RB_GENERATE_COND_FORMAT_HASH(lxlsx_cond_format_hash,
                                 lxlsx_cond_format_hash_element, tree_pointers,
                                 _cond_format_hash_cmp);
#endif

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Find but don't create a row object for a given row number.
 */
lxlsx_row *
lxlsx_worksheet_find_row(lxlsx_worksheet *self, lxlsx_row_t row_num)
{
    lxlsx_row tmp_row;

    tmp_row.row_num = row_num;

    return RB_FIND(lxlsx_table_rows, self->table, &tmp_row);
}

/*
 * Find but don't create a cell object for a given row object and col number.
 */
lxlsx_cell *
lxlsx_worksheet_find_cell_in_row(lxlsx_row *row, lxlsx_col_t col_num)
{
    lxlsx_cell tmp_cell;

    if (!row)
        return NULL;

    tmp_cell.col_num = col_num;

    return RB_FIND(lxlsx_table_cells, row->cells, &tmp_cell);
}

/*
 * Create a new worksheet object.
 */
lxlsx_worksheet *
lxlsx_worksheet_new(lxlsx_worksheet_init_data *init_data)
{
    lxlsx_worksheet *worksheet = calloc(1, sizeof(lxlsx_worksheet));
    GOTO_LABEL_ON_MEM_ERROR(worksheet, mem_error);

    worksheet->table = calloc(1, sizeof(struct lxlsx_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->table, mem_error);
    RB_INIT(worksheet->table);

    worksheet->hyperlinks = calloc(1, sizeof(struct lxlsx_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->hyperlinks, mem_error);
    RB_INIT(worksheet->hyperlinks);

    worksheet->comments = calloc(1, sizeof(struct lxlsx_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->comments, mem_error);
    RB_INIT(worksheet->comments);

    /* Initialize the cached rows. */
    worksheet->table->cached_row_num = LXLSX_ROW_MAX + 1;
    worksheet->hyperlinks->cached_row_num = LXLSX_ROW_MAX + 1;
    worksheet->comments->cached_row_num = LXLSX_ROW_MAX + 1;

    if (init_data && init_data->optimize) {
        worksheet->array = calloc(LXLSX_COL_MAX, sizeof(struct lxlsx_cell *));
        GOTO_LABEL_ON_MEM_ERROR(worksheet->array, mem_error);
    }

    worksheet->col_options =
        calloc(LXLSX_COL_META_MAX, sizeof(lxlsx_col_options *));
    worksheet->col_options_max = LXLSX_COL_META_MAX;
    GOTO_LABEL_ON_MEM_ERROR(worksheet->col_options, mem_error);

    worksheet->col_formats = calloc(LXLSX_COL_META_MAX, sizeof(lxlsx_format *));
    worksheet->col_formats_max = LXLSX_COL_META_MAX;
    GOTO_LABEL_ON_MEM_ERROR(worksheet->col_formats, mem_error);

    worksheet->optimize_row = calloc(1, sizeof(struct lxlsx_row));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_row, mem_error);
    worksheet->optimize_row->height = LXLSX_DEF_ROW_HEIGHT;

    worksheet->merged_ranges = calloc(1, sizeof(struct lxlsx_merged_ranges));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->merged_ranges, mem_error);
    STAILQ_INIT(worksheet->merged_ranges);

    worksheet->image_props = calloc(1, sizeof(struct lxlsx_image_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->image_props, mem_error);
    STAILQ_INIT(worksheet->image_props);

    worksheet->embedded_image_props =
        calloc(1, sizeof(struct lxlsx_embedded_image_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->embedded_image_props, mem_error);
    STAILQ_INIT(worksheet->embedded_image_props);

    worksheet->lxlsx_chart_data = calloc(1, sizeof(struct lxlsx_chart_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_chart_data, mem_error);
    STAILQ_INIT(worksheet->lxlsx_chart_data);

    worksheet->comment_objs = calloc(1, sizeof(struct lxlsx_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->comment_objs, mem_error);
    STAILQ_INIT(worksheet->comment_objs);

    worksheet->header_image_objs = calloc(1, sizeof(struct lxlsx_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->header_image_objs, mem_error);
    STAILQ_INIT(worksheet->header_image_objs);

    worksheet->button_objs = calloc(1, sizeof(struct lxlsx_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->button_objs, mem_error);
    STAILQ_INIT(worksheet->button_objs);

    worksheet->selections = calloc(1, sizeof(struct lxlsx_selections));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->selections, mem_error);
    STAILQ_INIT(worksheet->selections);

    worksheet->data_validations =
        calloc(1, sizeof(struct lxlsx_data_validations));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->data_validations, mem_error);
    STAILQ_INIT(worksheet->data_validations);

    worksheet->lxlsx_table_objs = calloc(1, sizeof(struct lxlsx_table_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_table_objs, mem_error);
    STAILQ_INIT(worksheet->lxlsx_table_objs);

    worksheet->external_hyperlinks = calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_hyperlinks, mem_error);
    STAILQ_INIT(worksheet->external_hyperlinks);

    worksheet->external_drawing_links =
        calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_drawing_links, mem_error);
    STAILQ_INIT(worksheet->external_drawing_links);

    worksheet->lxlsx_drawing_links = calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_drawing_links, mem_error);
    STAILQ_INIT(worksheet->lxlsx_drawing_links);

    worksheet->lxlsx_vml_drawing_links = calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_vml_drawing_links, mem_error);
    STAILQ_INIT(worksheet->lxlsx_vml_drawing_links);

    worksheet->external_table_links =
        calloc(1, sizeof(struct lxlsx_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_table_links, mem_error);
    STAILQ_INIT(worksheet->external_table_links);

    if (init_data && init_data->optimize) {
        FILE *tmpfile;

        worksheet->optimize_buffer = NULL;
        worksheet->optimize_buffer_size = 0;
        tmpfile = lxlsx_get_filehandle(&worksheet->optimize_buffer,
                                     &worksheet->optimize_buffer_size,
                                     init_data->tmpdir);
        if (!tmpfile) {
            LXLSX_ERROR("Error creating tmpfile() for worksheet in "
                      "'constant_memory' mode.");
            goto mem_error;
        }

        worksheet->optimize_tmpfile = tmpfile;
        GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_tmpfile, mem_error);
        worksheet->file = worksheet->optimize_tmpfile;
    }

    worksheet->lxlsx_drawing_rel_ids =
        calloc(1, sizeof(struct lxlsx_drawing_rel_ids));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_drawing_rel_ids, mem_error);
    RB_INIT(worksheet->lxlsx_drawing_rel_ids);

    worksheet->lxlsx_vml_drawing_rel_ids =
        calloc(1, sizeof(struct lxlsx_vml_drawing_rel_ids));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->lxlsx_vml_drawing_rel_ids, mem_error);
    RB_INIT(worksheet->lxlsx_vml_drawing_rel_ids);

    worksheet->conditional_formats =
        calloc(1, sizeof(struct lxlsx_cond_format_hash));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->conditional_formats, mem_error);
    RB_INIT(worksheet->conditional_formats);

    /* Initialize the worksheet dimensions. */
    worksheet->dim_rowmax = 0;
    worksheet->dim_colmax = 0;
    worksheet->dim_rowmin = LXLSX_ROW_MAX;
    worksheet->dim_colmin = LXLSX_COL_MAX;

    worksheet->default_row_height = LXLSX_DEF_ROW_HEIGHT;
    worksheet->default_row_pixels = 20;
    worksheet->default_col_pixels = 64;

    /* Initialize the page setup properties. */
    worksheet->fit_height = 0;
    worksheet->fit_width = 0;
    worksheet->page_start = 0;
    worksheet->print_scale = 100;
    worksheet->fit_page = 0;
    worksheet->orientation = LXLSX_TRUE;
    worksheet->page_order = 0;
    worksheet->page_setup_changed = LXLSX_FALSE;
    worksheet->page_view = LXLSX_FALSE;
    worksheet->paper_size = 0;
    worksheet->vertical_dpi = 0;
    worksheet->horizontal_dpi = 0;
    worksheet->margin_left = 0.7;
    worksheet->margin_right = 0.7;
    worksheet->margin_top = 0.75;
    worksheet->margin_bottom = 0.75;
    worksheet->margin_header = 0.3;
    worksheet->margin_footer = 0.3;
    worksheet->print_gridlines = 0;
    worksheet->screen_gridlines = 1;
    worksheet->print_options_changed = 0;
    worksheet->zoom = 100;
    worksheet->zoom_scale_normal = LXLSX_TRUE;
    worksheet->show_zeros = LXLSX_TRUE;
    worksheet->outline_on = LXLSX_TRUE;
    worksheet->outline_style = LXLSX_TRUE;
    worksheet->outline_below = LXLSX_TRUE;
    worksheet->outline_right = LXLSX_FALSE;
    worksheet->tab_color = LXLSX_COLOR_UNSET;
    worksheet->max_url_length = 2079;
    worksheet->comment_display_default = LXLSX_COMMENT_DISPLAY_HIDDEN;

    worksheet->header_footer_objs[0] = &worksheet->header_left_object_props;
    worksheet->header_footer_objs[1] = &worksheet->header_center_object_props;
    worksheet->header_footer_objs[2] = &worksheet->header_right_object_props;
    worksheet->header_footer_objs[3] = &worksheet->footer_left_object_props;
    worksheet->header_footer_objs[4] = &worksheet->footer_center_object_props;
    worksheet->header_footer_objs[5] = &worksheet->footer_right_object_props;

    if (init_data) {
        worksheet->name = init_data->name;
        worksheet->quoted_name = init_data->quoted_name;
        worksheet->tmpdir = init_data->tmpdir;
        worksheet->index = init_data->index;
        worksheet->hidden = init_data->hidden;
        worksheet->sst = init_data->sst;
        worksheet->optimize = init_data->optimize;
        worksheet->active_sheet = init_data->active_sheet;
        worksheet->first_sheet = init_data->first_sheet;
        worksheet->default_url_format = init_data->default_url_format;
        worksheet->max_url_length = init_data->max_url_length;
        worksheet->use_1904_epoch = init_data->use_1904_epoch;
    }

    return worksheet;

mem_error:
    lxlsx_worksheet_free(worksheet);
    return NULL;
}

/*
 * Free vml object.
 */
STATIC void
_free_vml_object(lxlsx_vml_obj *lxlsx_vml_obj)
{
    if (!lxlsx_vml_obj)
        return;

    free(lxlsx_vml_obj->author);
    free(lxlsx_vml_obj->font_name);
    free(lxlsx_vml_obj->text);
    free(lxlsx_vml_obj->image_position);
    free(lxlsx_vml_obj->name);
    free(lxlsx_vml_obj->macro);

    free(lxlsx_vml_obj);
}

/*
 * Free autofilter rule object.
 */
STATIC void
_free_filter_rule(lxlsx_filter_rule_obj *rule_obj)
{
    uint16_t i;

    if (!rule_obj)
        return;

    free(rule_obj->value1_string);
    free(rule_obj->value2_string);

    if (rule_obj->list) {
        for (i = 0; i < rule_obj->num_list_filters; i++)
            free(rule_obj->list[i]);

        free(rule_obj->list);
    }

    free(rule_obj);
}

/*
 * Free autofilter rules.
 */
STATIC void
_free_filter_rules(lxlsx_worksheet *worksheet)
{
    uint16_t i;

    if (!worksheet->filter_rules)
        return;

    for (i = 0; i < worksheet->num_filter_rules; i++)
        _free_filter_rule(worksheet->filter_rules[i]);

    free(worksheet->filter_rules);
}

/*
 * Free a worksheet cell.
 */
STATIC void
_free_cell(lxlsx_cell *cell)
{
    if (!cell)
        return;

    if (cell->type != NUMBER_CELL && cell->type != STRING_CELL
        && cell->type != BLANK_CELL && cell->type != BOOLEAN_CELL
        && cell->type != ERROR_CELL) {

        free((void *) cell->u.string);
    }

    free(cell->user_data1);
    free(cell->user_data2);

    _free_vml_object(cell->comment);

    free(cell);
}

/*
 * Free a worksheet row.
 */
STATIC void
_free_row(lxlsx_row *row)
{
    lxlsx_cell *cell;
    lxlsx_cell *next_cell;

    if (!row)
        return;

    for (cell = RB_MIN(lxlsx_table_cells, row->cells); cell; cell = next_cell) {
        next_cell = RB_NEXT(lxlsx_table_cells, row->cells, cell);
        RB_REMOVE(lxlsx_table_cells, row->cells, cell);
        _free_cell(cell);
    }

    free(row->cells);
    free(row);
}

/*
 * Free a worksheet image_options.
 */
STATIC void
_free_object_properties(lxlsx_object_properties *object_property)
{
    if (!object_property)
        return;

    free(object_property->filename);
    free(object_property->description);
    free(object_property->extension);
    free(object_property->url);
    free(object_property->tip);
    free(object_property->image_buffer);
    free(object_property->md5);
    free(object_property->image_position);
    free(object_property);
    object_property = NULL;
}

/*
 * Free a worksheet data_validation.
 */
STATIC void
_free_data_validation(lxlsx_data_val_obj *data_validation)
{
    if (!data_validation)
        return;

    free(data_validation->value_formula);
    free(data_validation->maximum_formula);
    free(data_validation->input_title);
    free(data_validation->input_message);
    free(data_validation->error_title);
    free(data_validation->error_message);
    free(data_validation->minimum_formula);

    free(data_validation);
}

/*
 * Free a worksheet conditional format obj.
 */
STATIC void
_free_cond_format(lxlsx_cond_format_obj *cond_format)
{
    if (!cond_format)
        return;

    free(cond_format->min_value_string);
    free(cond_format->mid_value_string);
    free(cond_format->max_value_string);
    free(cond_format->type_string);
    free(cond_format->guid);

    free(cond_format);
}

/*
 * Free a relationship structure.
 */
STATIC void
_free_relationship(lxlsx_rel_tuple *relationship)
{
    if (!relationship)
        return;

    free(relationship->type);
    free(relationship->target);
    free(relationship->target_mode);

    free(relationship);
}

/*
 * Free a worksheet table column object.
 */
STATIC void
_free_worksheet_table_column(lxlsx_table_column *column)
{
    if (!column)
        return;

    free((void *) column->header);
    free((void *) column->formula);
    free((void *) column->total_string);

    free(column);
}

/*
 * Free a worksheet table  object.
 */
STATIC void
_free_worksheet_table(lxlsx_table_obj *table)
{
    uint16_t i;

    if (!table)
        return;

    for (i = 0; i < table->num_cols; i++)
        _free_worksheet_table_column(table->columns[i]);

    free(table->name);
    free(table->total_string);
    free(table->columns);

    free(table);
}

/*
 * Free a worksheet object.
 */
void
lxlsx_worksheet_free(lxlsx_worksheet *worksheet)
{
    lxlsx_row *row;
    lxlsx_row *next_row;
    lxlsx_col_t col;
    lxlsx_merged_range *merged_range;
    lxlsx_object_properties *object_props;
    lxlsx_vml_obj *lxlsx_vml_obj;
    lxlsx_selection *selection;
    lxlsx_data_val_obj *data_validation;
    lxlsx_rel_tuple *relationship;
    lxlsx_cond_format_obj *cond_format;
    lxlsx_table_obj *lxlsx_table_obj;
    struct lxlsx_drawing_rel_id *lxlsx_drawing_rel_id;
    struct lxlsx_drawing_rel_id *next_drawing_rel_id;
    struct lxlsx_cond_format_hash_element *cond_format_elem;
    struct lxlsx_cond_format_hash_element *next_cond_format_elem;

    if (!worksheet)
        return;

    free(worksheet->edit_sheet_name);

    if (worksheet->col_options) {
        for (col = 0; col < worksheet->col_options_max; col++) {
            if (worksheet->col_options[col])
                free(worksheet->col_options[col]);
        }
    }

    free(worksheet->col_options);
    free(worksheet->col_sizes);
    free(worksheet->col_formats);

    if (worksheet->table) {
        for (row = RB_MIN(lxlsx_table_rows, worksheet->table); row;
             row = next_row) {

            next_row = RB_NEXT(lxlsx_table_rows, worksheet->table, row);
            RB_REMOVE(lxlsx_table_rows, worksheet->table, row);
            _free_row(row);
        }

        free(worksheet->table);
    }

    if (worksheet->hyperlinks) {
        for (row = RB_MIN(lxlsx_table_rows, worksheet->hyperlinks); row;
             row = next_row) {

            next_row = RB_NEXT(lxlsx_table_rows, worksheet->hyperlinks, row);
            RB_REMOVE(lxlsx_table_rows, worksheet->hyperlinks, row);
            _free_row(row);
        }

        free(worksheet->hyperlinks);
    }

    if (worksheet->comments) {
        for (row = RB_MIN(lxlsx_table_rows, worksheet->comments); row;
             row = next_row) {

            next_row = RB_NEXT(lxlsx_table_rows, worksheet->comments, row);
            RB_REMOVE(lxlsx_table_rows, worksheet->comments, row);
            _free_row(row);
        }

        free(worksheet->comments);
    }

    if (worksheet->merged_ranges) {
        while (!STAILQ_EMPTY(worksheet->merged_ranges)) {
            merged_range = STAILQ_FIRST(worksheet->merged_ranges);
            STAILQ_REMOVE_HEAD(worksheet->merged_ranges, list_pointers);
            free(merged_range);
        }

        free(worksheet->merged_ranges);
    }

    if (worksheet->image_props) {
        while (!STAILQ_EMPTY(worksheet->image_props)) {
            object_props = STAILQ_FIRST(worksheet->image_props);
            STAILQ_REMOVE_HEAD(worksheet->image_props, list_pointers);
            _free_object_properties(object_props);
        }

        free(worksheet->image_props);
    }

    if (worksheet->embedded_image_props) {
        while (!STAILQ_EMPTY(worksheet->embedded_image_props)) {
            object_props = STAILQ_FIRST(worksheet->embedded_image_props);
            STAILQ_REMOVE_HEAD(worksheet->embedded_image_props,
                               list_pointers);
            _free_object_properties(object_props);
        }

        free(worksheet->embedded_image_props);
    }

    if (worksheet->lxlsx_chart_data) {
        while (!STAILQ_EMPTY(worksheet->lxlsx_chart_data)) {
            object_props = STAILQ_FIRST(worksheet->lxlsx_chart_data);
            STAILQ_REMOVE_HEAD(worksheet->lxlsx_chart_data, list_pointers);
            _free_object_properties(object_props);
        }

        free(worksheet->lxlsx_chart_data);
    }

    /* Just free the list. The list objects are freed from the RB tree. */
    free(worksheet->comment_objs);

    if (worksheet->header_image_objs) {
        while (!STAILQ_EMPTY(worksheet->header_image_objs)) {
            lxlsx_vml_obj = STAILQ_FIRST(worksheet->header_image_objs);
            STAILQ_REMOVE_HEAD(worksheet->header_image_objs, list_pointers);
            _free_vml_object(lxlsx_vml_obj);
        }

        free(worksheet->header_image_objs);
    }

    if (worksheet->button_objs) {
        while (!STAILQ_EMPTY(worksheet->button_objs)) {
            lxlsx_vml_obj = STAILQ_FIRST(worksheet->button_objs);
            STAILQ_REMOVE_HEAD(worksheet->button_objs, list_pointers);
            _free_vml_object(lxlsx_vml_obj);
        }

        free(worksheet->button_objs);
    }

    if (worksheet->selections) {
        while (!STAILQ_EMPTY(worksheet->selections)) {
            selection = STAILQ_FIRST(worksheet->selections);
            STAILQ_REMOVE_HEAD(worksheet->selections, list_pointers);
            free(selection);
        }

        free(worksheet->selections);
    }

    if (worksheet->lxlsx_table_objs) {
        while (!STAILQ_EMPTY(worksheet->lxlsx_table_objs)) {
            lxlsx_table_obj = STAILQ_FIRST(worksheet->lxlsx_table_objs);
            STAILQ_REMOVE_HEAD(worksheet->lxlsx_table_objs, list_pointers);
            _free_worksheet_table(lxlsx_table_obj);
        }

        free(worksheet->lxlsx_table_objs);
    }

    if (worksheet->data_validations) {
        while (!STAILQ_EMPTY(worksheet->data_validations)) {
            data_validation = STAILQ_FIRST(worksheet->data_validations);
            STAILQ_REMOVE_HEAD(worksheet->data_validations, list_pointers);
            _free_data_validation(data_validation);
        }

        free(worksheet->data_validations);
    }

    while (!STAILQ_EMPTY(worksheet->external_hyperlinks)) {
        relationship = STAILQ_FIRST(worksheet->external_hyperlinks);
        STAILQ_REMOVE_HEAD(worksheet->external_hyperlinks, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->external_hyperlinks);

    while (!STAILQ_EMPTY(worksheet->external_drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->external_drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->external_drawing_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->external_drawing_links);

    while (!STAILQ_EMPTY(worksheet->lxlsx_drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->lxlsx_drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->lxlsx_drawing_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->lxlsx_drawing_links);

    while (!STAILQ_EMPTY(worksheet->lxlsx_vml_drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->lxlsx_vml_drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->lxlsx_vml_drawing_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->lxlsx_vml_drawing_links);

    while (!STAILQ_EMPTY(worksheet->external_table_links)) {
        relationship = STAILQ_FIRST(worksheet->external_table_links);
        STAILQ_REMOVE_HEAD(worksheet->external_table_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->external_table_links);

    if (worksheet->lxlsx_drawing_rel_ids) {
        for (lxlsx_drawing_rel_id =
             RB_MIN(lxlsx_drawing_rel_ids, worksheet->lxlsx_drawing_rel_ids);
             lxlsx_drawing_rel_id; lxlsx_drawing_rel_id = next_drawing_rel_id) {

            next_drawing_rel_id =
                RB_NEXT(lxlsx_drawing_rel_ids, worksheet->lxlsx_drawing_rel_id,
                        lxlsx_drawing_rel_id);
            RB_REMOVE(lxlsx_drawing_rel_ids, worksheet->lxlsx_drawing_rel_ids,
                      lxlsx_drawing_rel_id);
            free(lxlsx_drawing_rel_id->target);
            free(lxlsx_drawing_rel_id);
        }

        free(worksheet->lxlsx_drawing_rel_ids);
    }

    if (worksheet->lxlsx_vml_drawing_rel_ids) {
        for (lxlsx_drawing_rel_id =
             RB_MIN(lxlsx_vml_drawing_rel_ids, worksheet->lxlsx_vml_drawing_rel_ids);
             lxlsx_drawing_rel_id; lxlsx_drawing_rel_id = next_drawing_rel_id) {

            next_drawing_rel_id =
                RB_NEXT(lxlsx_vml_drawing_rel_ids, worksheet->lxlsx_drawing_rel_id,
                        lxlsx_drawing_rel_id);
            RB_REMOVE(lxlsx_vml_drawing_rel_ids, worksheet->lxlsx_vml_drawing_rel_ids,
                      lxlsx_drawing_rel_id);
            free(lxlsx_drawing_rel_id->target);
            free(lxlsx_drawing_rel_id);
        }

        free(worksheet->lxlsx_vml_drawing_rel_ids);
    }

    if (worksheet->conditional_formats) {
        for (cond_format_elem =
             RB_MIN(lxlsx_cond_format_hash, worksheet->conditional_formats);
             cond_format_elem; cond_format_elem = next_cond_format_elem) {

            next_cond_format_elem = RB_NEXT(lxlsx_cond_format_hash,
                                            worksheet->conditional_formats,
                                            cond_format_elem);
            RB_REMOVE(lxlsx_cond_format_hash,
                      worksheet->conditional_formats, cond_format_elem);

            while (!STAILQ_EMPTY(cond_format_elem->cond_formats)) {
                cond_format = STAILQ_FIRST(cond_format_elem->cond_formats);
                STAILQ_REMOVE_HEAD(cond_format_elem->cond_formats,
                                   list_pointers);
                _free_cond_format(cond_format);
            }

            free(cond_format_elem->cond_formats);
            free(cond_format_elem);
        }

        free(worksheet->conditional_formats);
    }

    _free_relationship(worksheet->external_vml_comment_link);
    _free_relationship(worksheet->external_comment_link);
    _free_relationship(worksheet->external_vml_header_link);
    _free_relationship(worksheet->external_background_link);

    _free_filter_rules(worksheet);

    if (worksheet->array) {
        for (col = 0; col < LXLSX_COL_MAX; col++) {
            _free_cell(worksheet->array[col]);
        }
        free(worksheet->array);
    }

    if (worksheet->optimize_row)
        free(worksheet->optimize_row);

    if (worksheet->drawing)
        lxlsx_drawing_free(worksheet->drawing);

    free(worksheet->hbreaks);
    free(worksheet->vbreaks);
    free((void *) worksheet->name);
    free((void *) worksheet->quoted_name);
    free(worksheet->vba_codename);
    free(worksheet->lxlsx_vml_data_id_str);
    free(worksheet->lxlsx_vml_header_id_str);
    free(worksheet->comment_author);
    free(worksheet->ignore_number_stored_as_text);
    free(worksheet->ignore_eval_error);
    free(worksheet->ignore_formula_differs);
    free(worksheet->ignore_formula_range);
    free(worksheet->ignore_formula_unlocked);
    free(worksheet->ignore_empty_cell_reference);
    free(worksheet->ignore_list_data_validation);
    free(worksheet->ignore_calculated_column);
    free(worksheet->ignore_two_digit_text_year);
    free(worksheet->header);
    free(worksheet->footer);

    free(worksheet);
    worksheet = NULL;
}

/*
 * Create a new worksheet row object.
 */
STATIC lxlsx_row *
_new_row(lxlsx_row_t row_num)
{
    lxlsx_row *row = calloc(1, sizeof(lxlsx_row));

    if (row) {
        row->row_num = row_num;
        row->cells = calloc(1, sizeof(struct lxlsx_table_cells));
        row->height = LXLSX_DEF_ROW_HEIGHT;

        if (row->cells)
            RB_INIT(row->cells);
        else
            LXLSX_MEM_ERROR();
    }
    else {
        LXLSX_MEM_ERROR();
    }

    return row;
}

/*
 * Create a new worksheet number cell object.
 */
STATIC lxlsx_cell *
_new_number_cell(lxlsx_row_t row_num,
                 lxlsx_col_t col_num, double value, lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = NUMBER_CELL;
    cell->format = format;
    cell->u.number = value;

    return cell;
}

/*
 * Create a new worksheet string cell object.
 */
STATIC lxlsx_cell *
_new_string_cell(lxlsx_row_t row_num,
                 lxlsx_col_t col_num, int32_t string_id, char *lxlsx_sst_string,
                 lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = STRING_CELL;
    cell->format = format;
    cell->u.string_id = string_id;
    cell->lxlsx_sst_string = lxlsx_sst_string;

    return cell;
}

/*
 * Create a new worksheet inline_string cell object.
 */
STATIC lxlsx_cell *
_new_inline_string_cell(lxlsx_row_t row_num,
                        lxlsx_col_t col_num, char *string, lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = INLINE_STRING_CELL;
    cell->format = format;
    cell->u.string = string;

    return cell;
}

/*
 * Create a new worksheet inline_string cell object for rich strings.
 */
STATIC lxlsx_cell *
_new_inline_rich_string_cell(lxlsx_row_t row_num,
                             lxlsx_col_t col_num, const char *string,
                             lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = INLINE_RICH_STRING_CELL;
    cell->format = format;
    cell->u.string = string;

    return cell;
}

/*
 * Create a new worksheet formula cell object.
 */
STATIC lxlsx_cell *
_new_formula_cell(lxlsx_row_t row_num,
                  lxlsx_col_t col_num, char *formula, lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = FORMULA_CELL;
    cell->format = format;
    cell->u.string = formula;

    return cell;
}

/*
 * Create a new worksheet array formula cell object.
 */
STATIC lxlsx_cell *
_new_array_formula_cell(lxlsx_row_t row_num, lxlsx_col_t col_num, char *formula,
                        char *range, lxlsx_format *format, uint8_t is_dynamic)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->format = format;
    cell->u.string = formula;
    cell->user_data1 = range;

    if (is_dynamic)
        cell->type = DYNAMIC_ARRAY_FORMULA_CELL;
    else
        cell->type = ARRAY_FORMULA_CELL;

    return cell;
}

/*
 * Create a new worksheet blank cell object.
 */
STATIC lxlsx_cell *
_new_blank_cell(lxlsx_row_t row_num, lxlsx_col_t col_num, lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = BLANK_CELL;
    cell->format = format;

    return cell;
}

/*
 * Create a new worksheet boolean cell object.
 */
STATIC lxlsx_cell *
_new_boolean_cell(lxlsx_row_t row_num, lxlsx_col_t col_num, int value,
                  lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = BOOLEAN_CELL;
    cell->format = format;
    cell->u.number = value;

    return cell;
}

/*
 * Create a new worksheet error cell object.
 */
STATIC lxlsx_cell *
_new_error_cell(lxlsx_row_t row_num, lxlsx_col_t col_num, uint32_t value,
                lxlsx_format *format)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = ERROR_CELL;
    cell->format = format;
    cell->u.number = value;

    return cell;
}

/*
 * Create a new comment cell object.
 */
STATIC lxlsx_cell *
_new_comment_cell(lxlsx_row_t row_num, lxlsx_col_t col_num,
                  lxlsx_vml_obj *comment_obj)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = COMMENT;
    cell->comment = comment_obj;

    return cell;
}

/*
 * Create a new worksheet hyperlink cell object.
 */
STATIC lxlsx_cell *
_new_hyperlink_cell(lxlsx_row_t row_num, lxlsx_col_t col_num,
                    enum cell_types link_type, char *url, char *string,
                    char *tooltip)
{
    lxlsx_cell *cell = calloc(1, sizeof(lxlsx_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = link_type;
    cell->u.string = url;
    cell->user_data1 = string;
    cell->user_data2 = tooltip;

    return cell;
}

/*
 * Get or create the row object for a given row number.
 */
STATIC lxlsx_row *
_get_row_list(struct lxlsx_table_rows *table, lxlsx_row_t row_num)
{
    lxlsx_row *row;
    lxlsx_row *existing_row;

    if (table->cached_row_num == row_num)
        return table->cached_row;

    /* Create a new row and try and insert it. */
    row = _new_row(row_num);
    existing_row = RB_INSERT(lxlsx_table_rows, table, row);

    /* If existing_row is not NULL, then it already existed. Free new row */
    /* and return existing_row. */
    if (existing_row) {
        _free_row(row);
        row = existing_row;
    }

    table->cached_row = row;
    table->cached_row_num = row_num;

    return row;
}

/*
 * Get or create the row object for a given row number.
 */
STATIC lxlsx_row *
_get_row(lxlsx_worksheet *self, lxlsx_row_t row_num)
{
    lxlsx_row *row;

    if (!self->optimize) {
        row = _get_row_list(self->table, row_num);
        return row;
    }
    else {
        if (row_num < self->optimize_row->row_num) {
            return NULL;
        }
        else if (row_num == self->optimize_row->row_num) {
            return self->optimize_row;
        }
        else {
            /* Flush row. */
            lxlsx_worksheet_write_single_row(self);
            row = self->optimize_row;
            row->row_num = row_num;
            return row;
        }
    }
}

/*
 * Insert a cell object in the cell list of a row object.
 */
STATIC void
_insert_cell_list(struct lxlsx_table_cells *cell_list,
                  lxlsx_cell *cell, lxlsx_col_t col_num)
{
    lxlsx_cell *existing_cell;

    cell->col_num = col_num;

    existing_cell = RB_INSERT(lxlsx_table_cells, cell_list, cell);

    /* If existing_cell is not NULL, then that cell already existed. */
    /* Remove existing_cell and add new one in again. */
    if (existing_cell) {
        RB_REMOVE(lxlsx_table_cells, cell_list, existing_cell);

        /* Add it in again. */
        RB_INSERT(lxlsx_table_cells, cell_list, cell);
        _free_cell(existing_cell);
    }

    return;
}

/*
 * Insert a cell object into the cell list or array.
 */
STATIC void
_insert_cell(lxlsx_worksheet *self, lxlsx_row_t row_num, lxlsx_col_t col_num,
             lxlsx_cell *cell)
{
    lxlsx_row *row = _get_row(self, row_num);

    if (!self->optimize) {
        row->data_changed = LXLSX_TRUE;
        _insert_cell_list(row->cells, cell, col_num);
    }
    else {
        if (row) {
            row->data_changed = LXLSX_TRUE;

            /* Overwrite an existing cell if necessary. */
            if (self->array[col_num])
                _free_cell(self->array[col_num]);

            self->array[col_num] = cell;
        }
    }
}

/*
 * Insert a blank placeholder cell in the cells RB tree in the same position
 * as a comment so that the rows "spans" calculation is correct. Since the
 * blank cell doesn't have a format it is ignored when writing. If there is
 * already a cell in the required position we don't have add a new cell.
 */
STATIC void
_insert_cell_placeholder(lxlsx_worksheet *self, lxlsx_row_t row_num,
                         lxlsx_col_t col_num)
{
    lxlsx_row *row;
    lxlsx_cell *cell;

    /* The spans calculation isn't required in constant_memory mode. */
    if (self->optimize)
        return;

    cell = _new_blank_cell(row_num, col_num, NULL);
    if (!cell)
        return;

    /* Only add a cell if one doesn't already exist. */
    row = _get_row(self, row_num);
    if (!RB_FIND(lxlsx_table_cells, row->cells, cell)) {
        _insert_cell_list(row->cells, cell, col_num);
    }
    else {
        _free_cell(cell);
    }
}

/*
 * Insert a hyperlink object into the hyperlink RB tree.
 */
STATIC void
_insert_hyperlink(lxlsx_worksheet *self, lxlsx_row_t row_num, lxlsx_col_t col_num,
                  lxlsx_cell *link)
{
    lxlsx_row *row = _get_row_list(self->hyperlinks, row_num);

    _insert_cell_list(row->cells, link, col_num);
}

/*
 * Insert a comment into the comment RB tree.
 */
STATIC void
_insert_comment(lxlsx_worksheet *self, lxlsx_row_t row_num, lxlsx_col_t col_num,
                lxlsx_cell *link)
{
    lxlsx_row *row = _get_row_list(self->comments, row_num);

    _insert_cell_list(row->cells, link, col_num);
}

/*
 * Next power of two for column reallocs. Taken from bithacks in the public
 * domain.
 */
STATIC lxlsx_col_t
_next_power_of_two(uint16_t col)
{
    col--;
    col |= col >> 1;
    col |= col >> 2;
    col |= col >> 4;
    col |= col >> 8;
    col++;

    return col;
}

/*
 * Check that row and col are within the allowed Excel range and store max
 * and min values for use in other methods/elements.
 *
 * The ignore_row/ignore_col flags are used to indicate that we wish to
 * perform the dimension check without storing the value.
 */
STATIC lxlsx_error
_check_dimensions(lxlsx_worksheet *self,
                  lxlsx_row_t row_num,
                  lxlsx_col_t col_num, int8_t ignore_row, int8_t ignore_col)
{
    if (row_num >= LXLSX_ROW_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    if (col_num >= LXLSX_COL_MAX)
        return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    /* In optimization mode we don't change dimensions for rows that are */
    /* already written. */
    if (!ignore_row && !ignore_col && self->optimize) {
        if (row_num < self->optimize_row->row_num)
            return LXLSX_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
    }

    if (!ignore_row) {
        if (row_num < self->dim_rowmin)
            self->dim_rowmin = row_num;
        if (row_num > self->dim_rowmax)
            self->dim_rowmax = row_num;
    }

    if (!ignore_col) {
        if (col_num < self->dim_colmin)
            self->dim_colmin = col_num;
        if (col_num > self->dim_colmax)
            self->dim_colmax = col_num;
    }

    return LXLSX_NO_ERROR;
}

/*
 * Comparator for the row structure red/black tree.
 */
STATIC int
_row_cmp(lxlsx_row *row1, lxlsx_row *row2)
{
    if (row1->row_num > row2->row_num)
        return 1;
    if (row1->row_num < row2->row_num)
        return -1;
    return 0;
}

/*
 * Comparator for the cell structure red/black tree.
 */
STATIC int
_cell_cmp(lxlsx_cell *cell1, lxlsx_cell *cell2)
{
    if (cell1->col_num > cell2->col_num)
        return 1;
    if (cell1->col_num < cell2->col_num)
        return -1;
    return 0;
}

/*
 * Comparator for the image/hyperlink relationship ids.
 */
STATIC int
_drawing_rel_id_cmp(lxlsx_drawing_rel_id *rel_id1, lxlsx_drawing_rel_id *rel_id2)
{
    return strcmp(rel_id1->target, rel_id2->target);
}

/*
 * Comparator for the conditional format RB hash elements.
 */
STATIC int
_cond_format_hash_cmp(lxlsx_cond_format_hash_element *elem_1,
                      lxlsx_cond_format_hash_element *elem_2)
{
    return strcmp(elem_1->sqref, elem_2->sqref);
}

/*
 * Get the index used to address a drawing rel link.
 */
STATIC uint32_t
_get_drawing_rel_index(lxlsx_worksheet *self, char *target)
{
    lxlsx_drawing_rel_id tmp_drawing_rel_id;
    lxlsx_drawing_rel_id *found_duplicate_target = NULL;
    lxlsx_drawing_rel_id *new_drawing_rel_id = NULL;

    if (target) {
        tmp_drawing_rel_id.target = target;
        found_duplicate_target = RB_FIND(lxlsx_drawing_rel_ids,
                                         self->lxlsx_drawing_rel_ids,
                                         &tmp_drawing_rel_id);
    }

    if (found_duplicate_target) {
        return found_duplicate_target->id;
    }
    else {
        self->lxlsx_drawing_rel_id++;

        if (target) {
            new_drawing_rel_id = calloc(1, sizeof(lxlsx_drawing_rel_id));

            if (new_drawing_rel_id) {
                new_drawing_rel_id->id = self->lxlsx_drawing_rel_id;
                new_drawing_rel_id->target = lxlsx_strdup(target);

                RB_INSERT(lxlsx_drawing_rel_ids, self->lxlsx_drawing_rel_ids,
                          new_drawing_rel_id);
            }
        }

        return self->lxlsx_drawing_rel_id;
    }
}

/*
 * find the index used to address a drawing rel link.
 */
STATIC uint32_t
_find_drawing_rel_index(lxlsx_worksheet *self, char *target)
{
    lxlsx_drawing_rel_id tmp_drawing_rel_id;
    lxlsx_drawing_rel_id *found_duplicate_target = NULL;

    if (!target)
        return 0;

    tmp_drawing_rel_id.target = target;
    found_duplicate_target = RB_FIND(lxlsx_drawing_rel_ids,
                                     self->lxlsx_drawing_rel_ids,
                                     &tmp_drawing_rel_id);

    if (found_duplicate_target)
        return found_duplicate_target->id;
    else
        return 0;
}

/*
 * Get the index used to address a VMLdrawing rel link.
 */
STATIC uint32_t
_get_vml_drawing_rel_index(lxlsx_worksheet *self, char *target)
{
    lxlsx_drawing_rel_id tmp_drawing_rel_id;
    lxlsx_drawing_rel_id *found_duplicate_target = NULL;
    lxlsx_drawing_rel_id *new_drawing_rel_id = NULL;

    if (target) {
        tmp_drawing_rel_id.target = target;
        found_duplicate_target = RB_FIND(lxlsx_vml_drawing_rel_ids,
                                         self->lxlsx_vml_drawing_rel_ids,
                                         &tmp_drawing_rel_id);
    }

    if (found_duplicate_target) {
        return found_duplicate_target->id;
    }
    else {
        self->lxlsx_vml_drawing_rel_id++;

        if (target) {
            new_drawing_rel_id = calloc(1, sizeof(lxlsx_drawing_rel_id));

            if (new_drawing_rel_id) {
                new_drawing_rel_id->id = self->lxlsx_vml_drawing_rel_id;
                new_drawing_rel_id->target = lxlsx_strdup(target);

                RB_INSERT(lxlsx_vml_drawing_rel_ids, self->lxlsx_vml_drawing_rel_ids,
                          new_drawing_rel_id);
            }
        }

        return self->lxlsx_vml_drawing_rel_id;
    }
}

/*
 * find the index used to address a VML drawing rel link.
 */
STATIC uint32_t
_find_vml_drawing_rel_index(lxlsx_worksheet *self, char *target)
{
    lxlsx_drawing_rel_id tmp_drawing_rel_id;
    lxlsx_drawing_rel_id *found_duplicate_target = NULL;

    if (!target)
        return 0;

    tmp_drawing_rel_id.target = target;
    found_duplicate_target = RB_FIND(lxlsx_vml_drawing_rel_ids,
                                     self->lxlsx_vml_drawing_rel_ids,
                                     &tmp_drawing_rel_id);

    if (found_duplicate_target)
        return found_duplicate_target->id;
    else
        return 0;
}

/*
 * Simple replacement for libgen.h basename() for compatibility with MSVC. It
 * handles forward and back slashes. It doesn't copy exactly the return
 * format of basename().
 */
const char *
lxlsx_basename(const char *path)
{

    const char *forward_slash;
    const char *back_slash;

    if (!path)
        return NULL;

    forward_slash = strrchr(path, '/');
    back_slash = strrchr(path, '\\');

    if (!forward_slash && !back_slash)
        return path;

    if (forward_slash > back_slash)
        return forward_slash + 1;
    else
        return back_slash + 1;
}

/* Function to count the total concatenated length of the strings in a
 * validation list array, including commas. */
size_t
_validation_list_length(const char **list)
{
    uint8_t i = 0;
    size_t length = 0;

    if (!list || !list[0])
        return 0;

    while (list[i] && length < LXLSX_VALIDATION_MAX_STRING_LENGTH) {
        /* Include commas in the length. */
        length += 1 + lxlsx_utf8_strlen(list[i]);
        i++;
    }

    /* Adjust the count for extraneous comma at end. */
    length--;

    return length;
}

/* Function to convert an array of strings into a CSV string for data
 * validation lists. */
char *
_validation_list_to_csv(const char **list)
{
    uint8_t i = 0;
    char *str;

    /* Create a buffer for the concatenated, and quoted, string. */
    /* Allow for 4 byte UTF-8 chars and add 3 bytes for quotes and EOL. */
    str = calloc(1, LXLSX_VALIDATION_MAX_STRING_LENGTH * 4 + 3);
    if (!str)
        return NULL;

    /* Add the start quote and first element. */
    strcat(str, "\"");
    strcat(str, list[0]);

    /* Add the other elements preceded by a comma. */
    i = 1;
    while (list[i]) {
        strcat(str, ",");
        strcat(str, list[i]);
        i++;
    }

    /* Add the end quote. */
    strcat(str, "\"");

    return str;
}

STATIC double
_pixels_to_width(double pixels)
{
    double max_digit_width = 7.0;
    double padding = 5.0;
    double width;

    if (pixels == LXLSX_DEF_COL_WIDTH_PIXELS)
        width = LXLSX_DEF_COL_WIDTH;
    else if (pixels <= 12.0)
        width = pixels / (max_digit_width + padding);
    else
        width = (pixels - padding) / max_digit_width;

    return width;
}

STATIC double
_pixels_to_height(double pixels)
{
    if (pixels == LXLSX_DEF_ROW_HEIGHT_PIXELS)
        return LXLSX_DEF_ROW_HEIGHT;
    else
        return pixels * 0.75;
}

/* Check and set if an autofilter is a standard or custom filter. */
void
_set_custom_filter(lxlsx_filter_rule_obj *rule_obj)
{
    rule_obj->is_custom = LXLSX_TRUE;

    if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_EQUAL_TO)
        rule_obj->is_custom = LXLSX_FALSE;

    if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_BLANKS)
        rule_obj->is_custom = LXLSX_FALSE;

    if (rule_obj->criteria2 != LXLSX_FILTER_CRITERIA_NONE) {
        if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_EQUAL_TO)
            rule_obj->is_custom = LXLSX_FALSE;

        if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_BLANKS)
            rule_obj->is_custom = LXLSX_FALSE;

        if (rule_obj->type == LXLSX_FILTER_TYPE_AND)
            rule_obj->is_custom = LXLSX_TRUE;
    }

    if (rule_obj->value1_string && strpbrk(rule_obj->value1_string, "*?"))
        rule_obj->is_custom = LXLSX_TRUE;

    if (rule_obj->value2_string && strpbrk(rule_obj->value2_string, "*?"))
        rule_obj->is_custom = LXLSX_TRUE;
}

/* Check and copy user input for table styles in lxlsx_worksheet_add_table(). */
void
_check_and_copy_table_style(lxlsx_table_obj *lxlsx_table_obj,
                            lxlsx_table_options *user_options)
{
    if (!user_options)
        return;

    /* Set the defaults. */
    lxlsx_table_obj->style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM;
    lxlsx_table_obj->style_type_number = 9;

    if (user_options->style_type > LXLSX_TABLE_STYLE_TYPE_DARK) {
        LXLSX_WARN_FORMAT1
            ("lxlsx_worksheet_add_table(): invalid style_type = %d. "
             "Using default TableStyleMedium9", user_options->style_type);

        lxlsx_table_obj->style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM;
        lxlsx_table_obj->style_type_number = 9;
    }
    else {
        lxlsx_table_obj->style_type = user_options->style_type;
    }

    /* Each type (light, medium and dark) has a different number of styles. */
    if (user_options->style_type == LXLSX_TABLE_STYLE_TYPE_LIGHT) {
        if (user_options->style_type_number > 21) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXLSX_TABLE_STYLE_TYPE_LIGHT. "
                             "Using default TableStyleMedium9",
                             user_options->style_type);

            lxlsx_table_obj->style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM;
            lxlsx_table_obj->style_type_number = 9;
        }
        else {
            lxlsx_table_obj->style_type_number = user_options->style_type_number;
        }
    }

    if (user_options->style_type == LXLSX_TABLE_STYLE_TYPE_MEDIUM) {
        if (user_options->style_type_number < 1
            || user_options->style_type_number > 28) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXLSX_TABLE_STYLE_TYPE_MEDIUM. "
                             "Using default TableStyleMedium9",
                             user_options->style_type_number);

            lxlsx_table_obj->style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM;
            lxlsx_table_obj->style_type_number = 9;
        }
        else {
            lxlsx_table_obj->style_type_number = user_options->style_type_number;
        }
    }

    if (user_options->style_type == LXLSX_TABLE_STYLE_TYPE_DARK) {
        if (user_options->style_type_number < 1
            || user_options->style_type_number > 11) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXLSX_TABLE_STYLE_TYPE_DARK. "
                             "Using default TableStyleMedium9",
                             user_options->style_type_number);

            lxlsx_table_obj->style_type = LXLSX_TABLE_STYLE_TYPE_MEDIUM;
            lxlsx_table_obj->style_type_number = 9;
        }
        else {
            lxlsx_table_obj->style_type_number = user_options->style_type_number;
        }
    }
}

/* Set the defaults for table columns in lxlsx_worksheet_add_table(). */
lxlsx_error
_set_default_table_columns(lxlsx_table_obj *lxlsx_table_obj)
{

    char col_name[LXLSX_ATTR_32];
    char *header;
    uint16_t i;
    lxlsx_table_column *column;
    uint16_t num_cols = lxlsx_table_obj->num_cols;
    lxlsx_table_column **columns = lxlsx_table_obj->columns;

    for (i = 0; i < num_cols; i++) {
        lxlsx_snprintf(col_name, LXLSX_ATTR_32, "Column%d", i + 1);

        column = calloc(num_cols, sizeof(lxlsx_table_column));
        RETURN_ON_MEM_ERROR(column, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

        header = lxlsx_strdup(col_name);
        if (!header) {
            free(column);
            RETURN_ON_MEM_ERROR(header, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
        }
        columns[i] = column;
        columns[i]->header = header;
    }

    return LXLSX_NO_ERROR;
}

/*  Convert Excel 2010 style "@" structural references to the Excel 2007 style
 *  "[#This Row]" in table formulas. This is the format that Excel uses to
 *  store the references. */
char *
_expand_table_formula(const char *formula)
{
    char *expanded;
    const char *ptr;
    size_t i;
    size_t ref_count = 0;
    size_t expanded_len;

    ptr = formula;

    while (*ptr) {
        if (*ptr == '@')
            ref_count++;

        ptr++;
    }

    if (ref_count == 0) {
        /* String doesn't need to be expanded. Just copy it. */
        expanded = lxlsx_strdup_formula(formula);
    }
    else {
        /* Convert "@" in the formula string to "[#This Row],".  */
        expanded_len = strlen(formula) + (sizeof(LXLSX_THIS_ROW) * ref_count);
        expanded = calloc(1, expanded_len);

        if (!expanded)
            return NULL;

        i = 0;
        ptr = formula;
        /* Ignore the = in the formula. */
        if (*ptr == '=')
            ptr++;

        /* Do the "@" expansion. */
        while (*ptr) {
            if (*ptr == '@') {
                strcat(&expanded[i], LXLSX_THIS_ROW);
                i += sizeof(LXLSX_THIS_ROW) - 1;
            }
            else {
                expanded[i] = *ptr;
                i++;
            }

            ptr++;
        }
    }

    return expanded;
}

/* Set user values for table columns in lxlsx_worksheet_add_table(). */
lxlsx_error
_set_custom_table_columns(lxlsx_table_obj *lxlsx_table_obj,
                          lxlsx_table_options *user_options)
{
    char *str;
    uint16_t i;
    lxlsx_table_column *table_column;
    lxlsx_table_column *user_column;
    uint16_t num_cols = lxlsx_table_obj->num_cols;
    lxlsx_table_column **user_columns = user_options->columns;

    for (i = 0; i < num_cols; i++) {

        user_column = user_columns[i];
        table_column = lxlsx_table_obj->columns[i];

        /* NULL indicates end of user input array. */
        if (user_column == NULL)
            return LXLSX_NO_ERROR;

        if (user_column->header) {
            if (lxlsx_utf8_strlen(user_column->header) > 255) {
                LXLSX_WARN_FORMAT("lxlsx_worksheet_add_table(): column parameter "
                                "'header' exceeds Excel length limit of 255.");
                return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
            }

            str = lxlsx_strdup(user_column->header);
            RETURN_ON_MEM_ERROR(str, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

            /* Free the default column header. */
            free((void *) table_column->header);
            table_column->header = str;
        }

        if (user_column->total_string) {
            str = lxlsx_strdup(user_column->total_string);
            RETURN_ON_MEM_ERROR(str, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

            table_column->total_string = str;
        }

        if (user_column->formula) {
            str = _expand_table_formula(user_column->formula);
            RETURN_ON_MEM_ERROR(str, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

            table_column->formula = str;
        }

        table_column->format = user_column->format;
        table_column->total_value = user_column->total_value;
        table_column->header_format = user_column->header_format;
        table_column->total_function = user_column->total_function;
    }

    return LXLSX_NO_ERROR;
}

/* Write a worksheet table column formula like SUBTOTAL(109,[Column1]). */
void
_write_column_function(lxlsx_worksheet *self, lxlsx_row_t row, lxlsx_col_t col,
                       lxlsx_table_column *column)
{
    size_t offset;
    char formula[LXLSX_MAX_ATTRIBUTE_LENGTH];
    lxlsx_format *format = column->format;
    uint8_t total_function = column->total_function;
    double value = column->total_value;
    const char *header = column->header;

    /* Write the subtotal formula number. */
    lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH, "SUBTOTAL(%d,[",
                 total_function);

    /* Copy the header string but escape any special characters. Note, this is
     * guaranteed to fit in the 2k buffer since the header is max 255
     * characters, checked in _set_custom_table_columns(). */
    offset = strlen(formula);
    while (*header) {
        switch (*header) {
            case '\'':
            case '#':
            case '[':
            case ']':
                formula[offset++] = '\'';
                formula[offset] = *header;
                break;
            default:
                formula[offset] = *header;
                break;
        }
        offset++;
        header++;
    }

    /* Write the end of the string. */
    memcpy(&formula[offset], "])\0", sizeof("])\0"));

    lxlsx_worksheet_write_formula_num(self, row, col, formula, format, value);
}

/* Write a worksheet table column formula. */
void
_write_column_formula(lxlsx_worksheet *self, lxlsx_row_t first_row,
                      lxlsx_row_t last_row, lxlsx_col_t col,
                      lxlsx_table_column *column)
{
    lxlsx_row_t row;
    const char *formula = column->formula;
    lxlsx_format *format = column->format;

    for (row = first_row; row <= last_row; row++)
        lxlsx_worksheet_write_formula(self, row, col, formula, format);
}

/* Set the defaults for table columns in lxlsx_worksheet_add_table(). */
void
_write_table_column_data(lxlsx_worksheet *self, lxlsx_table_obj *lxlsx_table_obj)
{
    uint16_t i;
    lxlsx_table_column *column;
    lxlsx_table_column **columns = lxlsx_table_obj->columns;

    lxlsx_col_t col;
    lxlsx_row_t first_row = lxlsx_table_obj->first_row;
    lxlsx_col_t first_col = lxlsx_table_obj->first_col;
    lxlsx_row_t last_row = lxlsx_table_obj->last_row;
    lxlsx_row_t first_data_row = first_row;
    lxlsx_row_t last_data_row = last_row;

    if (!lxlsx_table_obj->no_header_row)
        first_data_row++;

    if (lxlsx_table_obj->total_row)
        last_data_row--;

    for (i = 0; i < lxlsx_table_obj->num_cols; i++) {
        col = first_col + i;
        column = columns[i];

        if (lxlsx_table_obj->no_header_row == LXLSX_FALSE)
            lxlsx_worksheet_write_string(self, first_row, col, column->header,
                                   column->header_format);

        if (column->total_string)
            lxlsx_worksheet_write_string(self, last_row, col, column->total_string,
                                   NULL);

        if (column->total_function)
            _write_column_function(self, last_row, col, column);

        if (column->formula)
            _write_column_formula(self, first_data_row, last_data_row, col,
                                  column);
    }
}

/*
 * Check that there are sufficient data rows in a worksheet table.
 */
lxlsx_error
_check_table_rows(lxlsx_row_t first_row, lxlsx_row_t last_row,
                  lxlsx_table_options *user_options)
{
    lxlsx_row_t num_non_header_rows = last_row - first_row;

    if (user_options && user_options->no_header_row == LXLSX_TRUE)
        num_non_header_rows++;

    if (num_non_header_rows == 0) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_add_table(): "
                        "table must have at least 1 non-header row.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    return LXLSX_NO_ERROR;
}

/*
 * Check that the the table name is valid.
 */
lxlsx_error
_check_table_name(lxlsx_table_options *user_options)
{
    const char *name;
    char *ptr;
    char first[2] = { 0, 0 };

    if (!user_options)
        return LXLSX_NO_ERROR;

    if (!user_options->name)
        return LXLSX_NO_ERROR;

    name = user_options->name;

    /* Check table name length. */
    if (lxlsx_utf8_strlen(name) > 255) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_add_table(): "
                        "Table name exceeds Excel's limit of 255.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Check some short invalid names. */
    if (strlen(name) == 1
        && (name[0] == 'C' || name[0] == 'c' || name[0] == 'R'
            || name[0] == 'r')) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_add_table(): "
                         "invalid table name \"%s\".", name);
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Check for invalid characters in Table name, while trying to allow
     * for utf8 strings. */
    ptr = strpbrk(name, " !\"#$%&'()*+,-/:;<=>?@[\\]^`{|}~");
    if (ptr) {
        LXLSX_WARN_FORMAT2("lxlsx_worksheet_add_table(): "
                         "invalid character '%c' in table name \"%s\".",
                         *ptr, name);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check for invalid initial character in Table name, while trying to allow
     * for utf8 strings. */
    first[0] = name[0];
    ptr = strpbrk(first, " !\"#$%&'()*+,-./0123456789:;<=>?@[\\]^`{|}~");
    if (ptr) {
        LXLSX_WARN_FORMAT2("lxlsx_worksheet_add_table(): "
                         "invalid first character '%c' in table name \"%s\".",
                         *ptr, name);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    return LXLSX_NO_ERROR;
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
_worksheet_xml_declaration(lxlsx_worksheet *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <worksheet> element.
 */
STATIC void
_worksheet_write_worksheet(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main";
    char xmlns_r[] = "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships";
    char xmlns_mc[] = "http://schemas.openxmlformats.org/"
        "markup-compatibility/2006";
    char xmlns_x14ac[] = "http://schemas.microsoft.com/"
        "office/spreadsheetml/2009/9/ac";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    if (self->excel_version == 2010) {
        LXLSX_PUSH_ATTRIBUTES_STR("xmlns:mc", xmlns_mc);
        LXLSX_PUSH_ATTRIBUTES_STR("xmlns:x14ac", xmlns_x14ac);
        LXLSX_PUSH_ATTRIBUTES_STR("mc:Ignorable", "x14ac");
    }

    lxlsx_xml_start_tag(self->file, "worksheet", &attributes);
    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <dimension> element.
 */
STATIC void
_worksheet_write_dimension(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char ref[LXLSX_MAX_CELL_RANGE_LENGTH];
    lxlsx_row_t dim_rowmin = self->dim_rowmin;
    lxlsx_row_t dim_rowmax = self->dim_rowmax;
    lxlsx_col_t dim_colmin = self->dim_colmin;
    lxlsx_col_t dim_colmax = self->dim_colmax;

    if (dim_rowmin == LXLSX_ROW_MAX && dim_colmin == LXLSX_COL_MAX) {
        /* If the rows and cols are still the defaults then no dimensions have
         * been set and we use the default range "A1". */
        lxlsx_rowcol_to_range(ref, 0, 0, 0, 0);
    }
    else if (dim_rowmin == LXLSX_ROW_MAX && dim_colmin != LXLSX_COL_MAX) {
        /* If the rows aren't set but the columns are then the dimensions have
         * been changed via set_column(). */
        lxlsx_rowcol_to_range(ref, 0, dim_colmin, 0, dim_colmax);
    }
    else {
        lxlsx_rowcol_to_range(ref, dim_rowmin, dim_colmin, dim_rowmax,
                            dim_colmax);
    }

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", ref);

    lxlsx_xml_empty_tag(self->file, "dimension", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <pane> element for freeze panes.
 */
STATIC void
_worksheet_write_freeze_panes(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    lxlsx_selection *selection;
    lxlsx_selection *user_selection;
    lxlsx_row_t row = self->panes.first_row;
    lxlsx_col_t col = self->panes.first_col;
    lxlsx_row_t top_row = self->panes.top_row;
    lxlsx_col_t left_col = self->panes.left_col;

    char row_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char col_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char top_left_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char active_pane[LXLSX_PANE_NAME_LENGTH];

    /* If there is a user selection we remove it from the list and use it. */
    if (!STAILQ_EMPTY(self->selections)) {
        user_selection = STAILQ_FIRST(self->selections);
        STAILQ_REMOVE_HEAD(self->selections, list_pointers);
    }
    else {
        /* or else create a new blank selection. */
        user_selection = calloc(1, sizeof(lxlsx_selection));
        RETURN_VOID_ON_MEM_ERROR(user_selection);
    }

    LXLSX_INIT_ATTRIBUTES();

    lxlsx_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* Set the active pane. */
    if (row && col) {
        lxlsx_strcpy(active_pane, "bottomRight");

        lxlsx_rowcol_to_cell(row_cell, row, 0);
        lxlsx_rowcol_to_cell(col_cell, 0, col);

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "topRight");
            lxlsx_strcpy(selection->active_cell, col_cell);
            lxlsx_strcpy(selection->sqref, col_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomLeft");
            lxlsx_strcpy(selection->active_cell, row_cell);
            lxlsx_strcpy(selection->sqref, row_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomRight");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else if (col) {
        lxlsx_strcpy(active_pane, "topRight");

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "topRight");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else {
        lxlsx_strcpy(active_pane, "bottomLeft");

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomLeft");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }

    if (col)
        LXLSX_PUSH_ATTRIBUTES_INT("xSplit", col);

    if (row)
        LXLSX_PUSH_ATTRIBUTES_INT("ySplit", row);

    LXLSX_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);
    LXLSX_PUSH_ATTRIBUTES_STR("activePane", active_pane);

    if (self->panes.type == FREEZE_PANES)
        LXLSX_PUSH_ATTRIBUTES_STR("state", "frozen");
    else if (self->panes.type == FREEZE_SPLIT_PANES)
        LXLSX_PUSH_ATTRIBUTES_STR("state", "frozenSplit");

    lxlsx_xml_empty_tag(self->file, "pane", &attributes);

    free(user_selection);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Convert column width from user units to pane split width.
 */
STATIC uint32_t
_worksheet_calculate_x_split_width(double x_split)
{
    uint32_t width;
    uint32_t pixels;
    uint32_t points;
    uint32_t twips;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;

    /* Convert to pixels. */
    if (x_split < 1.0) {
        pixels = (uint32_t) (x_split * (max_digit_width + padding) + 0.5);
    }
    else {
        pixels = (uint32_t) (x_split * max_digit_width + 0.5) + 5;
    }

    /* Convert to points. */
    points = (pixels * 3) / 4;

    /* Convert to twips (twentieths of a point). */
    twips = points * 20;

    /* Add offset/padding. */
    width = twips + 390;

    return width;
}

/*
 * Write the <pane> element for split panes.
 */
STATIC void
_worksheet_write_split_panes(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    lxlsx_selection *selection;
    lxlsx_selection *user_selection;
    lxlsx_row_t row = self->panes.first_row;
    lxlsx_col_t col = self->panes.first_col;
    lxlsx_row_t top_row = self->panes.top_row;
    lxlsx_col_t left_col = self->panes.left_col;
    double x_split = self->panes.x_split;
    double y_split = self->panes.y_split;
    uint8_t has_selection = LXLSX_FALSE;

    char row_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char col_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char top_left_cell[LXLSX_MAX_CELL_NAME_LENGTH];
    char active_pane[LXLSX_PANE_NAME_LENGTH];

    /* If there is a user selection we remove it from the list and use it. */
    if (!STAILQ_EMPTY(self->selections)) {
        user_selection = STAILQ_FIRST(self->selections);
        STAILQ_REMOVE_HEAD(self->selections, list_pointers);
        has_selection = LXLSX_TRUE;
    }
    else {
        /* or else create a new blank selection. */
        user_selection = calloc(1, sizeof(lxlsx_selection));
        RETURN_VOID_ON_MEM_ERROR(user_selection);
    }

    LXLSX_INIT_ATTRIBUTES();

    /* Convert the row and col to 1/20 twip units with padding. */
    if (y_split > 0.0)
        y_split = (uint32_t) (20 * y_split + 300);

    if (x_split > 0.0)
        x_split = _worksheet_calculate_x_split_width(x_split);

    /* For non-explicit topLeft definitions, estimate the cell offset based on
     * the pixels dimensions. This is only a workaround and doesn't take
     * adjusted cell dimensions into account.
     */
    if (top_row == row && left_col == col) {
        top_row = (lxlsx_row_t) (0.5 + (y_split - 300.0) / 20.0 / 15.0);
        left_col = (lxlsx_col_t) (0.5 + (x_split - 390.0) / 20.0 / 3.0 / 16.0);
    }

    lxlsx_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* If there is no selection set the active cell to the top left cell. */
    if (!has_selection) {
        lxlsx_strcpy(user_selection->active_cell, top_left_cell);
        lxlsx_strcpy(user_selection->sqref, top_left_cell);
    }

    /* Set the active pane. */
    if (y_split > 0.0 && x_split > 0.0) {
        lxlsx_strcpy(active_pane, "bottomRight");

        lxlsx_rowcol_to_cell(row_cell, top_row, 0);
        lxlsx_rowcol_to_cell(col_cell, 0, left_col);

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "topRight");
            lxlsx_strcpy(selection->active_cell, col_cell);
            lxlsx_strcpy(selection->sqref, col_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomLeft");
            lxlsx_strcpy(selection->active_cell, row_cell);
            lxlsx_strcpy(selection->sqref, row_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomRight");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else if (x_split > 0.0) {
        lxlsx_strcpy(active_pane, "topRight");

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "topRight");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else {
        lxlsx_strcpy(active_pane, "bottomLeft");

        selection = calloc(1, sizeof(lxlsx_selection));
        if (selection) {
            lxlsx_strcpy(selection->pane, "bottomLeft");
            lxlsx_strcpy(selection->active_cell, user_selection->active_cell);
            lxlsx_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }

    if (x_split > 0.0)
        LXLSX_PUSH_ATTRIBUTES_DBL("xSplit", x_split);

    if (y_split > 0.0)
        LXLSX_PUSH_ATTRIBUTES_DBL("ySplit", y_split);

    LXLSX_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);

    if (has_selection)
        LXLSX_PUSH_ATTRIBUTES_STR("activePane", active_pane);

    lxlsx_xml_empty_tag(self->file, "pane", &attributes);

    free(user_selection);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <selection> element.
 */
STATIC void
_worksheet_write_selection(lxlsx_worksheet *self, lxlsx_selection *selection)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (*selection->pane)
        LXLSX_PUSH_ATTRIBUTES_STR("pane", selection->pane);

    if (*selection->active_cell)
        LXLSX_PUSH_ATTRIBUTES_STR("activeCell", selection->active_cell);

    if (*selection->sqref)
        LXLSX_PUSH_ATTRIBUTES_STR("sqref", selection->sqref);

    lxlsx_xml_empty_tag(self->file, "selection", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <selection> elements.
 */
STATIC void
_worksheet_write_selections(lxlsx_worksheet *self)
{
    lxlsx_selection *selection;

    STAILQ_FOREACH(selection, self->selections, list_pointers) {
        _worksheet_write_selection(self, selection);
    }
}

/*
 * Write the frozen or split <pane> elements.
 */
STATIC void
_worksheet_write_panes(lxlsx_worksheet *self)
{
    if (self->panes.type == NO_PANES)
        return;

    else if (self->panes.type == FREEZE_PANES)
        _worksheet_write_freeze_panes(self);

    else if (self->panes.type == FREEZE_SPLIT_PANES)
        _worksheet_write_freeze_panes(self);

    else if (self->panes.type == SPLIT_PANES)
        _worksheet_write_split_panes(self);
}

/*
 * Write the <sheetView> element.
 */
STATIC void
_worksheet_write_sheet_view(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Hide screen gridlines if required */
    if (!self->screen_gridlines)
        LXLSX_PUSH_ATTRIBUTES_STR("showGridLines", "0");

    /* Hide zeroes in cells. */
    if (!self->show_zeros)
        LXLSX_PUSH_ATTRIBUTES_STR("showZeros", "0");

    /* Display worksheet right to left for Hebrew, Arabic and others. */
    if (self->right_to_left)
        LXLSX_PUSH_ATTRIBUTES_STR("rightToLeft", "1");

    /* Show that the sheet tab is selected. */
    if (self->selected)
        LXLSX_PUSH_ATTRIBUTES_STR("tabSelected", "1");

    /* Turn outlines off. Also required in the outlinePr element. */
    if (!self->outline_on)
        LXLSX_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

    /* Set the page view/layout mode if required. */
    if (self->page_view)
        LXLSX_PUSH_ATTRIBUTES_STR("view", "pageLayout");

    /* Set the top left cell if required. */
    if (self->top_left_cell[0])
        LXLSX_PUSH_ATTRIBUTES_STR("topLeftCell", self->top_left_cell);

    /* Set the zoom level. */
    if (self->zoom != 100 && !self->page_view) {
        LXLSX_PUSH_ATTRIBUTES_INT("zoomScale", self->zoom);

        if (self->zoom_scale_normal)
            LXLSX_PUSH_ATTRIBUTES_INT("zoomScaleNormal", self->zoom);
    }

    LXLSX_PUSH_ATTRIBUTES_STR("workbookViewId", "0");

    if (self->panes.type != NO_PANES || !STAILQ_EMPTY(self->selections)) {
        lxlsx_xml_start_tag(self->file, "sheetView", &attributes);
        _worksheet_write_panes(self);
        _worksheet_write_selections(self);
        lxlsx_xml_end_tag(self->file, "sheetView");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "sheetView", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetViews> element.
 */
STATIC void
_worksheet_write_sheet_views(lxlsx_worksheet *self)
{
    lxlsx_xml_start_tag(self->file, "sheetViews", NULL);

    /* Write the sheetView element. */
    _worksheet_write_sheet_view(self);

    lxlsx_xml_end_tag(self->file, "sheetViews");
}

/*
 * Write the <sheetFormatPr> element.
 */
STATIC void
_worksheet_write_sheet_format_pr(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_DBL("defaultRowHeight", self->default_row_height);

    if (self->default_row_height != LXLSX_DEF_ROW_HEIGHT)
        LXLSX_PUSH_ATTRIBUTES_STR("customHeight", "1");

    if (self->default_row_zeroed)
        LXLSX_PUSH_ATTRIBUTES_STR("zeroHeight", "1");

    if (self->outline_row_level)
        LXLSX_PUSH_ATTRIBUTES_INT("outlineLevelRow", self->outline_row_level);

    if (self->outline_col_level)
        LXLSX_PUSH_ATTRIBUTES_INT("outlineLevelCol", self->outline_col_level);

    if (self->excel_version == 2010)
        LXLSX_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

    lxlsx_xml_empty_tag(self->file, "sheetFormatPr", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetData> element.
 */
STATIC void
_worksheet_write_sheet_data(lxlsx_worksheet *self)
{
    if (RB_EMPTY(self->table)) {
        lxlsx_xml_empty_tag(self->file, "sheetData", NULL);
    }
    else {
        lxlsx_xml_start_tag(self->file, "sheetData", NULL);
        _worksheet_write_rows(self);
        lxlsx_xml_end_tag(self->file, "sheetData");
    }
}

/*
 * Write the <sheetData> element when the memory optimization is on. In which
 * case we read the data stored in the temp file and rewrite it to the XML
 * sheet file.
 */
STATIC void
_worksheet_write_optimized_sheet_data(lxlsx_worksheet *self)
{
    size_t read_size = 1;
    char buffer[LXLSX_BUFFER_SIZE];

    if (self->dim_rowmin == LXLSX_ROW_MAX) {
        /* If the dimensions aren't defined then there is no data to write. */
        lxlsx_xml_empty_tag(self->file, "sheetData", NULL);
    }
    else {

        lxlsx_xml_start_tag(self->file, "sheetData", NULL);

        /* Flush the temp file. */
        fflush(self->optimize_tmpfile);

        if (self->optimize_buffer) {
            /* Ignore return value. There is no easy way to raise error. */
            (void) fwrite(self->optimize_buffer, self->optimize_buffer_size,
                          1, self->file);
        }
        else {
            /* Rewind the temp file. */
            rewind(self->optimize_tmpfile);
            while (read_size) {
                read_size =
                    fread(buffer, 1, LXLSX_BUFFER_SIZE, self->optimize_tmpfile);
                /* Ignore return value. There is no easy way to raise error. */
                (void) fwrite(buffer, 1, read_size, self->file);
            }
        }

        fclose(self->optimize_tmpfile);
        free(self->optimize_buffer);

        lxlsx_xml_end_tag(self->file, "sheetData");
    }
}

/*
 * Write the <pageMargins> element.
 */
STATIC void
_worksheet_write_page_margins(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    double left = self->margin_left;
    double right = self->margin_right;
    double top = self->margin_top;
    double bottom = self->margin_bottom;
    double header = self->margin_header;
    double footer = self->margin_footer;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_DBL("left", left);
    LXLSX_PUSH_ATTRIBUTES_DBL("right", right);
    LXLSX_PUSH_ATTRIBUTES_DBL("top", top);
    LXLSX_PUSH_ATTRIBUTES_DBL("bottom", bottom);
    LXLSX_PUSH_ATTRIBUTES_DBL("header", header);
    LXLSX_PUSH_ATTRIBUTES_DBL("footer", footer);

    lxlsx_xml_empty_tag(self->file, "pageMargins", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <pageSetup> element.
 * The following is an example taken from Excel.
 * <pageSetup
 *     paperSize="9"
 *     scale="110"
 *     fitToWidth="2"
 *     fitToHeight="2"
 *     pageOrder="overThenDown"
 *     orientation="portrait"
 *     blackAndWhite="1"
 *     draft="1"
 *     horizontalDpi="200"
 *     verticalDpi="200"
 *     r:id="rId1"
 * />
 */
STATIC void
_worksheet_write_page_setup(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (!self->page_setup_changed)
        return;

    /* Set paper size. */
    if (self->paper_size)
        LXLSX_PUSH_ATTRIBUTES_INT("paperSize", self->paper_size);

    /* Set the print_scale. */
    if (self->print_scale != 100)
        LXLSX_PUSH_ATTRIBUTES_INT("scale", self->print_scale);

    /* Set the "Fit to page" properties. */
    if (self->fit_page && self->fit_width != 1)
        LXLSX_PUSH_ATTRIBUTES_INT("fitToWidth", self->fit_width);

    if (self->fit_page && self->fit_height != 1)
        LXLSX_PUSH_ATTRIBUTES_INT("fitToHeight", self->fit_height);

    /* Set the page print direction. */
    if (self->page_order)
        LXLSX_PUSH_ATTRIBUTES_STR("pageOrder", "overThenDown");

    /* Set start page. */
    if (self->page_start > 1)
        LXLSX_PUSH_ATTRIBUTES_INT("firstPageNumber", self->page_start);

    /* Set page orientation. */
    if (self->orientation)
        LXLSX_PUSH_ATTRIBUTES_STR("orientation", "portrait");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("orientation", "landscape");

    if (self->black_white)
        LXLSX_PUSH_ATTRIBUTES_STR("blackAndWhite", "1");

    /* Set start page active flag. */
    if (self->page_start)
        LXLSX_PUSH_ATTRIBUTES_INT("useFirstPageNumber", 1);

    /* Set the DPI. Mainly only for testing. */
    if (self->horizontal_dpi)
        LXLSX_PUSH_ATTRIBUTES_INT("horizontalDpi", self->horizontal_dpi);

    if (self->vertical_dpi)
        LXLSX_PUSH_ATTRIBUTES_INT("verticalDpi", self->vertical_dpi);

    lxlsx_xml_empty_tag(self->file, "pageSetup", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <printOptions> element.
 */
STATIC void
_worksheet_write_print_options(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    if (!self->print_options_changed)
        return;

    LXLSX_INIT_ATTRIBUTES();

    /* Set horizontal centering. */
    if (self->hcenter) {
        LXLSX_PUSH_ATTRIBUTES_STR("horizontalCentered", "1");
    }

    /* Set vertical centering. */
    if (self->vcenter) {
        LXLSX_PUSH_ATTRIBUTES_STR("verticalCentered", "1");
    }

    /* Enable row and column headers. */
    if (self->print_headers) {
        LXLSX_PUSH_ATTRIBUTES_STR("headings", "1");
    }

    /* Set printed gridlines. */
    if (self->print_gridlines) {
        LXLSX_PUSH_ATTRIBUTES_STR("gridLines", "1");
    }

    lxlsx_xml_empty_tag(self->file, "printOptions", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <row> element.
 */
STATIC void
_write_row(lxlsx_worksheet *self, lxlsx_row *row, char *spans)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    int32_t xf_index = 0;
    double height;

    if (row->format) {
        xf_index = lxlsx_format_get_xf_index(row->format);
    }

    if (row->height_changed)
        height = row->height;
    else
        height = self->default_row_height;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("r", row->row_num + 1);

    if (spans)
        LXLSX_PUSH_ATTRIBUTES_STR("spans", spans);

    if (xf_index)
        LXLSX_PUSH_ATTRIBUTES_INT("s", xf_index);

    if (row->format)
        LXLSX_PUSH_ATTRIBUTES_STR("customFormat", "1");

    if (height != LXLSX_DEF_ROW_HEIGHT)
        LXLSX_PUSH_ATTRIBUTES_DBL("ht", height);

    if (row->hidden)
        LXLSX_PUSH_ATTRIBUTES_STR("hidden", "1");

    if (height != LXLSX_DEF_ROW_HEIGHT)
        LXLSX_PUSH_ATTRIBUTES_STR("customHeight", "1");

    if (row->level)
        LXLSX_PUSH_ATTRIBUTES_INT("outlineLevel", row->level);

    if (row->collapsed)
        LXLSX_PUSH_ATTRIBUTES_STR("collapsed", "1");

    if (self->excel_version == 2010)
        LXLSX_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

    if (!row->data_changed)
        lxlsx_xml_empty_tag(self->file, "row", &attributes);
    else
        lxlsx_xml_start_tag(self->file, "row", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Convert the width of a cell from user's units to pixels. Excel rounds the
 * column width to the nearest pixel. If the width hasn't been set by the user
 * we use the default value. If the column is hidden it has a value of zero.
 */
STATIC int32_t
_worksheet_size_col(lxlsx_worksheet *self, lxlsx_col_t col_num, uint8_t anchor)
{
    lxlsx_col_options *col_opt = NULL;
    uint32_t pixels;
    double width;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;
    lxlsx_col_t col_index;

    /* Search for the col number in the array of col_options. Each col_option
     * entry contains the start and end column for a range.
     */
    for (col_index = 0; col_index < self->col_options_max; col_index++) {
        col_opt = self->col_options[col_index];

        if (col_opt) {
            if (col_num >= col_opt->firstcol && col_num <= col_opt->lastcol)
                break;
            else
                col_opt = NULL;
        }
    }

    if (col_opt) {
        width = col_opt->width;

        /* Convert to pixels. */
        if (col_opt->hidden && anchor != LXLSX_OBJECT_MOVE_AND_SIZE_AFTER) {
            pixels = 0;
        }
        else if (width < 1.0) {
            pixels = (uint32_t) (width * (max_digit_width + padding) + 0.5);
        }
        else {
            pixels = (uint32_t) (width * max_digit_width + 0.5) + 5;
        }
    }
    else {
        pixels = self->default_col_pixels;
    }

    return pixels;
}

/*
 * Convert the height of a cell from user's units to pixels. If the height
 * hasn't been set by the user we use the default value. If the row is hidden
 * it has a value of zero.
 */
STATIC int32_t
_worksheet_size_row(lxlsx_worksheet *self, lxlsx_row_t row_num, uint8_t anchor)
{
    lxlsx_row *row;
    uint32_t pixels;

    row = lxlsx_worksheet_find_row(self, row_num);

    /* Note, the 0.75 below is due to the difference between 72/96 DPI. */
    if (row) {
        if (row->hidden && anchor != LXLSX_OBJECT_MOVE_AND_SIZE_AFTER)
            pixels = 0;
        else
            pixels = (uint32_t) (row->height / 0.75);
    }
    else {
        pixels = (uint32_t) (self->default_row_height / 0.75);
    }

    return pixels;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in pixels.
 *         +------------+------------+
 *         |     A      |      B     |
 *   +-----+------------+------------+
 *   |     |(x1,y1)     |            |
 *   |  1  |(A1)._______|______      |
 *   |     |    |              |     |
 *   |     |    |              |     |
 *   +-----+----|    BITMAP    |-----+
 *   |     |    |              |     |
 *   |  2  |    |______________.     |
 *   |     |            |        (B2)|
 *   |     |            |     (x2,y2)|
 *   +---- +------------+------------+
 *
 * Example of an object that covers some of the area from cell A1 to cell B2.
 * Based on the width and height of the object we need to calculate 8 vars:
 *
 *     col_start, row_start, col_end, row_end, x1, y1, x2, y2.
 *
 * We also calculate the absolute x and y position of the top left vertex of
 * the object. This is required for images:
 *
 *    x_abs, y_abs
 *
 * The width and height of the cells that the object occupies can be variable
 * and have to be taken into account.
 *
 * The values of col_start and row_start are passed in from the calling
 * function. The values of col_end and row_end are calculated by subtracting
 * the width and height of the object from the width and height of the
 * underlying cells.
 */
STATIC void
_worksheet_position_object_pixels(lxlsx_worksheet *self,
                                  lxlsx_object_properties *object_props,
                                  lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_col_t col_start;        /* Column containing upper left corner.  */
    int32_t x1;                 /* Distance to left side of object.      */

    lxlsx_row_t row_start;        /* Row containing top left corner.       */
    int32_t y1;                 /* Distance to top of object.            */

    lxlsx_col_t col_end;          /* Column containing lower right corner. */
    double x2;                  /* Distance to right side of object.     */

    lxlsx_row_t row_end;          /* Row containing bottom right corner.   */
    double y2;                  /* Distance to bottom of object.         */

    double width;               /* Width of object frame.                */
    double height;              /* Height of object frame.               */

    uint32_t x_abs = 0;         /* Abs. distance to left side of object. */
    uint32_t y_abs = 0;         /* Abs. distance to top  side of object. */

    uint32_t i;
    uint8_t anchor = lxlsx_drawing_object->anchor;
    uint8_t ignore_anchor = LXLSX_OBJECT_POSITION_DEFAULT;

    col_start = object_props->col;
    row_start = object_props->row;
    x1 = object_props->x_offset;
    y1 = object_props->y_offset;
    width = object_props->width;
    height = object_props->height;

    /* Adjust start column for negative offsets. */
    while (x1 < 0 && col_start > 0) {
        x1 += _worksheet_size_col(self, col_start - 1, ignore_anchor);
        col_start--;
    }

    /* Adjust start row for negative offsets. */
    while (y1 < 0 && row_start > 0) {
        y1 += _worksheet_size_row(self, row_start - 1, ignore_anchor);
        row_start--;
    }

    /* Ensure that the image isn't shifted off the page at top left. */
    if (x1 < 0)
        x1 = 0;

    if (y1 < 0)
        y1 = 0;

    /* Calculate the absolute x offset of the top-left vertex. */
    if (self->col_size_changed) {
        for (i = 0; i < col_start; i++)
            x_abs += _worksheet_size_col(self, i, ignore_anchor);
    }
    else {
        /* Optimization for when the column widths haven't changed. */
        x_abs += self->default_col_pixels * col_start;
    }

    x_abs += x1;

    /* Calculate the absolute y offset of the top-left vertex. */
    /* Store the column change to allow optimizations. */
    if (self->row_size_changed) {
        for (i = 0; i < row_start; i++)
            y_abs += _worksheet_size_row(self, i, ignore_anchor);
    }
    else {
        /* Optimization for when the row heights haven"t changed. */
        y_abs += self->default_row_pixels * row_start;
    }

    y_abs += y1;

    /* Adjust start col for offsets that are greater than the col width. */
    while (x1 >= _worksheet_size_col(self, col_start, anchor)) {
        x1 -= _worksheet_size_col(self, col_start, ignore_anchor);
        col_start++;
    }

    /* Adjust start row for offsets that are greater than the row height. */
    while (y1 >= _worksheet_size_row(self, row_start, anchor)) {
        y1 -= _worksheet_size_row(self, row_start, ignore_anchor);
        row_start++;
    }

    /* Initialize end cell to the same as the start cell. */
    col_end = col_start;
    row_end = row_start;

    /* Only offset the image in the cell if the row/col is hidden. */
    if (_worksheet_size_col(self, col_start, anchor) > 0)
        width = width + x1;
    if (_worksheet_size_row(self, row_start, anchor) > 0)
        height = height + y1;

    /* Subtract the underlying cell widths to find the end cell. */
    while (width >= _worksheet_size_col(self, col_end, anchor)
           && col_end < LXLSX_COL_MAX) {
        width -= _worksheet_size_col(self, col_end, anchor);
        col_end++;
    }

    /* Subtract the underlying cell heights to find the end cell. */
    while (height >= _worksheet_size_row(self, row_end, anchor)
           && row_end < LXLSX_ROW_MAX) {
        height -= _worksheet_size_row(self, row_end, anchor);
        row_end++;
    }

    /* The end vertices are whatever is left from the width and height. */
    x2 = width;
    y2 = height;

    /* Add the dimensions to the drawing object. */
    lxlsx_drawing_object->from.col = col_start;
    lxlsx_drawing_object->from.row = row_start;
    lxlsx_drawing_object->from.col_offset = x1;
    lxlsx_drawing_object->from.row_offset = y1;
    lxlsx_drawing_object->to.col = col_end;
    lxlsx_drawing_object->to.row = row_end;
    lxlsx_drawing_object->to.col_offset = x2;
    lxlsx_drawing_object->to.row_offset = y2;
    lxlsx_drawing_object->col_absolute = x_abs;
    lxlsx_drawing_object->row_absolute = y_abs;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in EMUs. The vertices are expressed as English
 * Metric Units (EMUs). There are 12,700 EMUs per point.
 * Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
 */
STATIC void
_worksheet_position_object_emus(lxlsx_worksheet *self,
                                lxlsx_object_properties *image,
                                lxlsx_drawing_object *lxlsx_drawing_object)
{

    _worksheet_position_object_pixels(self, image, lxlsx_drawing_object);

    /* Convert the pixel values to EMUs. See above. */
    lxlsx_drawing_object->from.col_offset *= 9525;
    lxlsx_drawing_object->from.row_offset *= 9525;
    lxlsx_drawing_object->to.col_offset *= 9525;
    lxlsx_drawing_object->to.row_offset *= 9525;
    lxlsx_drawing_object->to.col_offset += 0.5;
    lxlsx_drawing_object->to.row_offset += 0.5;
    lxlsx_drawing_object->col_absolute *= 9525;
    lxlsx_drawing_object->row_absolute *= 9525;
}

/*
 * This function handles the additional optional parameters to
 * lxlsx_worksheet_write_comment_opt() as well as calculating the comment object
 * position and vertices.
 */
void
_get_comment_params(lxlsx_vml_obj *comment, lxlsx_comment_options *options)
{

    lxlsx_row_t start_row;
    lxlsx_col_t start_col;
    int32_t x_offset;
    int32_t y_offset;
    uint32_t height = 74;
    uint32_t width = 128;
    double x_scale = 1.0;
    double y_scale = 1.0;
    lxlsx_row_t row = comment->row;
    lxlsx_col_t col = comment->col;;

    /* Set the default start cell and offsets for the comment. These are
     * generally fixed in relation to the parent cell. However there are some
     * edge cases for cells at the, well yes, edges. */
    if (row == 0)
        y_offset = 2;
    else if (row == LXLSX_ROW_MAX - 3)
        y_offset = 16;
    else if (row == LXLSX_ROW_MAX - 2)
        y_offset = 16;
    else if (row == LXLSX_ROW_MAX - 1)
        y_offset = 14;
    else
        y_offset = 10;

    if (col == LXLSX_COL_MAX - 3)
        x_offset = 49;
    else if (col == LXLSX_COL_MAX - 2)
        x_offset = 49;
    else if (col == LXLSX_COL_MAX - 1)
        x_offset = 49;
    else
        x_offset = 15;

    if (row == 0)
        start_row = 0;
    else if (row == LXLSX_ROW_MAX - 3)
        start_row = LXLSX_ROW_MAX - 7;
    else if (row == LXLSX_ROW_MAX - 2)
        start_row = LXLSX_ROW_MAX - 6;
    else if (row == LXLSX_ROW_MAX - 1)
        start_row = LXLSX_ROW_MAX - 5;
    else
        start_row = row - 1;

    if (col == LXLSX_COL_MAX - 3)
        start_col = LXLSX_COL_MAX - 6;
    else if (col == LXLSX_COL_MAX - 2)
        start_col = LXLSX_COL_MAX - 5;
    else if (col == LXLSX_COL_MAX - 1)
        start_col = LXLSX_COL_MAX - 4;
    else
        start_col = col + 1;

    /* Set the default font properties. */
    comment->font_size = 8;
    comment->font_family = 2;

    /* Set any user defined options. */
    if (options) {

        if (options->width > 0.0)
            width = options->width;

        if (options->height > 0.0)
            height = options->height;

        if (options->x_scale > 0.0)
            x_scale = options->x_scale;

        if (options->y_scale > 0.0)
            y_scale = options->y_scale;

        if (options->x_offset != 0)
            x_offset = options->x_offset;

        if (options->y_offset != 0)
            y_offset = options->y_offset;

        if (options->start_row > 0 || options->start_col > 0) {
            start_row = options->start_row;
            start_col = options->start_col;
        }

        if (options->font_size > 0.0)
            comment->font_size = options->font_size;

        if (options->font_family > 0)
            comment->font_family = options->font_family;

        comment->visible = options->visible;
        comment->color = options->color;
        comment->author = lxlsx_strdup(options->author);
        comment->font_name = lxlsx_strdup(options->font_name);
    }

    /* Scale the width/height to the default/user scale and round to the
     * nearest pixel. */
    width = (uint32_t) (0.5 + x_scale * width);
    height = (uint32_t) (0.5 + y_scale * height);

    comment->width = width;
    comment->height = height;
    comment->start_col = start_col;
    comment->start_row = start_row;
    comment->x_offset = x_offset;
    comment->y_offset = y_offset;
}

/*
 * This function handles the additional optional parameters to
 * lxlsx_worksheet_insert_button() as well as calculating the button object
 * position and vertices.
 */
lxlsx_error
_get_button_params(lxlsx_vml_obj *button, uint16_t button_number,
                   lxlsx_button_options *options)
{

    int32_t x_offset = 0;
    int32_t y_offset = 0;
    uint32_t height = LXLSX_DEF_ROW_HEIGHT_PIXELS;
    uint32_t width = LXLSX_DEF_COL_WIDTH_PIXELS;
    double x_scale = 1.0;
    double y_scale = 1.0;
    lxlsx_row_t row = button->row;
    lxlsx_col_t col = button->col;
    char buffer[LXLSX_ATTR_32];
    uint8_t has_caption = LXLSX_FALSE;
    uint8_t has_macro = LXLSX_FALSE;
    size_t len;

    /* Set any user defined options. */
    if (options) {

        if (options->width > 0.0)
            width = options->width;

        if (options->height > 0.0)
            height = options->height;

        if (options->x_scale > 0.0)
            x_scale = options->x_scale;

        if (options->y_scale > 0.0)
            y_scale = options->y_scale;

        if (options->x_offset != 0)
            x_offset = options->x_offset;

        if (options->y_offset != 0)
            y_offset = options->y_offset;

        if (options->caption) {
            button->name = lxlsx_strdup(options->caption);
            RETURN_ON_MEM_ERROR(button->name, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
            has_caption = LXLSX_TRUE;
        }

        if (options->macro) {
            len = sizeof("[0]!") + strlen(options->macro);
            button->macro = calloc(1, len);
            RETURN_ON_MEM_ERROR(button->macro,
                                LXLSX_ERROR_MEMORY_MALLOC_FAILED);

            if (button->macro)
                lxlsx_snprintf(button->macro, len, "[0]!%s", options->macro);

            has_macro = LXLSX_TRUE;
        }

        if (options->description) {
            button->text = lxlsx_strdup(options->description);
            RETURN_ON_MEM_ERROR(button->text, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
        }
    }

    if (!has_caption) {
        lxlsx_snprintf(buffer, LXLSX_ATTR_32, "Button %d", button_number);
        button->name = lxlsx_strdup(buffer);
        RETURN_ON_MEM_ERROR(button->name, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
    }

    if (!has_macro) {
        lxlsx_snprintf(buffer, LXLSX_ATTR_32, "[0]!Button%d_Click",
                     button_number);
        button->macro = lxlsx_strdup(buffer);
        RETURN_ON_MEM_ERROR(button->macro, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
    }

    /* Scale the width/height to the default/user scale and round to the
     * nearest pixel. */
    width = (uint32_t) (0.5 + x_scale * width);
    height = (uint32_t) (0.5 + y_scale * height);

    button->width = width;
    button->height = height;
    button->start_col = col;
    button->start_row = row;
    button->x_offset = x_offset;
    button->y_offset = y_offset;

    return LXLSX_NO_ERROR;
}

/*
 * Calculate the lxlsx_vml_obj object position and vertices.
 */
void
_worksheet_position_vml_object(lxlsx_worksheet *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    lxlsx_object_properties object_props;
    lxlsx_drawing_object lxlsx_drawing_object;

    object_props.col = lxlsx_vml_obj->start_col;
    object_props.row = lxlsx_vml_obj->start_row;
    object_props.x_offset = lxlsx_vml_obj->x_offset;
    object_props.y_offset = lxlsx_vml_obj->y_offset;
    object_props.width = lxlsx_vml_obj->width;
    object_props.height = lxlsx_vml_obj->height;

    lxlsx_drawing_object.anchor = LXLSX_OBJECT_DONT_MOVE_DONT_SIZE;

    _worksheet_position_object_pixels(self, &object_props, &lxlsx_drawing_object);

    lxlsx_vml_obj->from.col = lxlsx_drawing_object.from.col;
    lxlsx_vml_obj->from.row = lxlsx_drawing_object.from.row;
    lxlsx_vml_obj->from.col_offset = lxlsx_drawing_object.from.col_offset;
    lxlsx_vml_obj->from.row_offset = lxlsx_drawing_object.from.row_offset;
    lxlsx_vml_obj->to.col = lxlsx_drawing_object.to.col;
    lxlsx_vml_obj->to.row = lxlsx_drawing_object.to.row;
    lxlsx_vml_obj->to.col_offset = lxlsx_drawing_object.to.col_offset;
    lxlsx_vml_obj->to.row_offset = lxlsx_drawing_object.to.row_offset;
    lxlsx_vml_obj->col_absolute = lxlsx_drawing_object.col_absolute;
    lxlsx_vml_obj->row_absolute = lxlsx_drawing_object.row_absolute;
}

/*
 * Set up image/drawings.
 */
void
lxlsx_worksheet_prepare_image(lxlsx_worksheet *self,
                            uint32_t image_ref_id, uint32_t lxlsx_drawing_id,
                            lxlsx_object_properties *object_props)
{
    lxlsx_drawing_object *lxlsx_drawing_object;
    lxlsx_rel_tuple *relationship;
    double width;
    double height;
    char *url;
    char *found_string;
    size_t i;
    char filename[LXLSX_FILENAME_LENGTH];
    enum cell_types link_type = HYPERLINK_URL;

    if (!self->drawing) {
        self->drawing = lxlsx_drawing_new();
        self->drawing->embedded = LXLSX_TRUE;
        RETURN_VOID_ON_MEM_ERROR(self->drawing);

        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/drawing");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", lxlsx_drawing_id);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_drawing_links, relationship,
                           list_pointers);
    }

    lxlsx_drawing_object = calloc(1, sizeof(*lxlsx_drawing_object));
    RETURN_VOID_ON_MEM_ERROR(lxlsx_drawing_object);

    lxlsx_drawing_object->anchor = LXLSX_OBJECT_MOVE_DONT_SIZE;
    if (object_props->object_position)
        lxlsx_drawing_object->anchor = object_props->object_position;

    lxlsx_drawing_object->type = LXLSX_DRAWING_IMAGE;
    lxlsx_drawing_object->description = lxlsx_strdup(object_props->description);
    lxlsx_drawing_object->tip = lxlsx_strdup(object_props->tip);
    lxlsx_drawing_object->rel_index = 0;
    lxlsx_drawing_object->url_rel_index = 0;
    lxlsx_drawing_object->decorative = object_props->decorative;

    /* Scale to user scale. */
    width = object_props->width * object_props->x_scale;
    height = object_props->height * object_props->y_scale;

    /* Scale by non 96dpi resolutions. */
    width *= 96.0 / object_props->x_dpi;
    height *= 96.0 / object_props->y_dpi;

    object_props->width = width;
    object_props->height = height;

    _worksheet_position_object_emus(self, object_props, lxlsx_drawing_object);

    /* Convert from pixels to emus. */
    lxlsx_drawing_object->width = (uint32_t) (0.5 + width * 9525);
    lxlsx_drawing_object->height = (uint32_t) (0.5 + height * 9525);

    lxlsx_add_drawing_object(self->drawing, lxlsx_drawing_object);

    if (object_props->url) {
        url = object_props->url;

        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/hyperlink");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        /* Check the link type. Default to external hyperlinks. */
        if (strstr(url, "internal:"))
            link_type = HYPERLINK_INTERNAL;
        else if (strstr(url, "external:"))
            link_type = HYPERLINK_EXTERNAL;
        else
            link_type = HYPERLINK_URL;

        /* Set the relationship object for each type of link. */
        if (link_type == HYPERLINK_INTERNAL) {
            relationship->target_mode = NULL;
            relationship->target = lxlsx_strdup(url + sizeof("internal") - 1);
            GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

            /* We need to prefix the internal link/range with #. */
            relationship->target[0] = '#';
        }
        else if (link_type == HYPERLINK_EXTERNAL) {
            relationship->target_mode = lxlsx_strdup("External");
            GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);

            /* Look for Windows style "C:/" link or Windows share "\\" link. */
            found_string = strchr(url + sizeof("external:") - 1, ':');
            if (!found_string)
                found_string = strstr(url, "\\\\");

            if (found_string) {
                /* Copy the url with some space at the start to overwrite
                 * "external:" with "file:///". */
                relationship->target = lxlsx_escape_url_characters(url + 1,
                                                                 LXLSX_TRUE);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                /* Add the file:/// URI to the url if absolute path. */
                memcpy(relationship->target, "file:///",
                       sizeof("file:///") - 1);
            }
            else {
                /* Copy the relative url without "external:". */
                relationship->target =
                    lxlsx_escape_url_characters(url + sizeof("external:") - 1,
                                              LXLSX_TRUE);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                /* Switch backslash to forward slash. */
                for (i = 0; i <= strlen(relationship->target); i++)
                    if (relationship->target[i] == '\\')
                        relationship->target[i] = '/';
            }

        }
        else {
            relationship->target_mode = lxlsx_strdup("External");
            GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);

            relationship->target =
                lxlsx_escape_url_characters(object_props->url, LXLSX_FALSE);
            GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);
        }

        /* Check if URL exceeds Excel's length limit. */
        if (lxlsx_utf8_strlen(url) > self->max_url_length) {
            LXLSX_WARN_FORMAT2("lxlsx_worksheet_insert_image()/_opt(): URL exceeds "
                             "Excel's allowable length of %d characters: %s",
                             self->max_url_length, url);
            goto mem_error;
        }

        if (!_find_drawing_rel_index(self, url)) {
            STAILQ_INSERT_TAIL(self->lxlsx_drawing_links, relationship,
                               list_pointers);
        }
        else {
            free(relationship->type);
            free(relationship->target);
            free(relationship->target_mode);
            free(relationship);
        }

        lxlsx_drawing_object->url_rel_index = _get_drawing_rel_index(self, url);

    }

    if (!_find_drawing_rel_index(self, object_props->md5)) {
        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/image");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                     object_props->extension);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->lxlsx_drawing_links, relationship, list_pointers);
    }

    lxlsx_drawing_object->rel_index =
        _get_drawing_rel_index(self, object_props->md5);

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
}

/*
 * Set up image/drawings for header/footer images.
 */
void
lxlsx_worksheet_prepare_header_image(lxlsx_worksheet *self,
                                   uint32_t image_ref_id,
                                   lxlsx_object_properties *object_props)
{
    lxlsx_rel_tuple *relationship = NULL;
    char filename[LXLSX_FILENAME_LENGTH];
    lxlsx_vml_obj *header_image_vml;
    char *extension;

    STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);

    if (!_find_vml_drawing_rel_index(self, object_props->md5)) {
        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        RETURN_VOID_ON_MEM_ERROR(relationship);

        relationship->type = lxlsx_strdup("/image");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                     object_props->extension);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->lxlsx_vml_drawing_links, relationship,
                           list_pointers);
    }

    header_image_vml = calloc(1, sizeof(lxlsx_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(header_image_vml, mem_error);

    header_image_vml->width = (uint32_t) object_props->width;
    header_image_vml->height = (uint32_t) object_props->height;
    header_image_vml->x_dpi = object_props->x_dpi;
    header_image_vml->y_dpi = object_props->y_dpi;
    header_image_vml->rel_index = 1;

    header_image_vml->image_position =
        lxlsx_strdup(object_props->image_position);
    header_image_vml->name = lxlsx_strdup(object_props->description);

    /* Strip the extension from the filename. */
    extension = strchr(header_image_vml->name, '.');
    if (extension)
        *extension = '\0';

    header_image_vml->rel_index =
        _get_vml_drawing_rel_index(self, object_props->md5);

    STAILQ_INSERT_TAIL(self->header_image_objs, header_image_vml,
                       list_pointers);

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
}

/*
 * Set up background image.
 */
void
lxlsx_worksheet_prepare_background(lxlsx_worksheet *self,
                                 uint32_t image_ref_id,
                                 lxlsx_object_properties *object_props)
{
    lxlsx_rel_tuple *relationship = NULL;
    char filename[LXLSX_FILENAME_LENGTH];

    STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);

    relationship = calloc(1, sizeof(lxlsx_rel_tuple));
    RETURN_VOID_ON_MEM_ERROR(relationship);

    relationship->type = lxlsx_strdup("/image");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxlsx_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                 object_props->extension);

    relationship->target = lxlsx_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    self->external_background_link = relationship;

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
}

/*
 * Set up chart/drawings.
 */
void
lxlsx_worksheet_prepare_chart(lxlsx_worksheet *self,
                            uint32_t lxlsx_chart_ref_id,
                            uint32_t lxlsx_drawing_id,
                            lxlsx_object_properties *object_props,
                            uint8_t is_chartsheet)
{
    lxlsx_drawing_object *lxlsx_drawing_object;
    lxlsx_rel_tuple *relationship;
    double width;
    double height;
    char filename[LXLSX_FILENAME_LENGTH];

    if (!self->drawing) {
        self->drawing = lxlsx_drawing_new();
        RETURN_VOID_ON_MEM_ERROR(self->drawing);

        if (is_chartsheet) {
            self->drawing->embedded = LXLSX_FALSE;
            self->drawing->orientation = self->orientation;
        }
        else {
            self->drawing->embedded = LXLSX_TRUE;
        }

        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/drawing");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", lxlsx_drawing_id);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_drawing_links, relationship,
                           list_pointers);
    }

    lxlsx_drawing_object = calloc(1, sizeof(*lxlsx_drawing_object));
    RETURN_VOID_ON_MEM_ERROR(lxlsx_drawing_object);

    lxlsx_drawing_object->anchor = LXLSX_OBJECT_MOVE_AND_SIZE;
    if (object_props->object_position)
        lxlsx_drawing_object->anchor = object_props->object_position;

    lxlsx_drawing_object->type = LXLSX_DRAWING_CHART;
    lxlsx_drawing_object->description = lxlsx_strdup(object_props->description);
    lxlsx_drawing_object->tip = NULL;
    lxlsx_drawing_object->rel_index = _get_drawing_rel_index(self, NULL);
    lxlsx_drawing_object->url_rel_index = 0;
    lxlsx_drawing_object->decorative = object_props->decorative;

    /* Scale to user scale. */
    width = object_props->width * object_props->x_scale;
    height = object_props->height * object_props->y_scale;

    /* Convert to the nearest pixel. */
    object_props->width = width;
    object_props->height = height;

    _worksheet_position_object_emus(self, object_props, lxlsx_drawing_object);

    /* Convert from pixels to emus. */
    lxlsx_drawing_object->width = (uint32_t) (0.5 + width * 9525);
    lxlsx_drawing_object->height = (uint32_t) (0.5 + height * 9525);

    lxlsx_add_drawing_object(self->drawing, lxlsx_drawing_object);

    relationship = calloc(1, sizeof(lxlsx_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxlsx_strdup("/chart");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxlsx_snprintf(filename, 32, "../charts/chart%d.xml", lxlsx_chart_ref_id);

    relationship->target = lxlsx_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    STAILQ_INSERT_TAIL(self->lxlsx_drawing_links, relationship, list_pointers);

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
}

/*
 * Set up VML objects, such as comments, in the worksheet.
 */
uint32_t
lxlsx_worksheet_prepare_vml_objects(lxlsx_worksheet *self,
                                  uint32_t lxlsx_vml_data_id,
                                  uint32_t lxlsx_vml_shape_id,
                                  uint32_t lxlsx_vml_drawing_id,
                                  uint32_t comment_id)
{
    lxlsx_row *row;
    lxlsx_cell *cell;
    lxlsx_rel_tuple *relationship;
    char filename[LXLSX_FILENAME_LENGTH];
    uint32_t comment_count = 0;
    uint32_t i;
    uint32_t tmp_data_id;
    size_t data_str_len = 0;
    size_t used = 0;
    char *lxlsx_vml_data_id_str;

    RB_FOREACH(row, lxlsx_table_rows, self->comments) {

        RB_FOREACH(cell, lxlsx_table_cells, row->cells) {
            /* Calculate the worksheet position of the comment. */
            _worksheet_position_vml_object(self, cell->comment);

            /* Store comment in a simple list for use by packager. */
            STAILQ_INSERT_TAIL(self->comment_objs, cell->comment,
                               list_pointers);
            comment_count++;
        }
    }

    /* Set up the VML relationship for comments/buttons/header images. */
    relationship = calloc(1, sizeof(lxlsx_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxlsx_strdup("/vmlDrawing");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxlsx_snprintf(filename, 32, "../drawings/vmlDrawing%d.vml",
                 lxlsx_vml_drawing_id);

    relationship->target = lxlsx_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    self->external_vml_comment_link = relationship;

    if (self->has_comments) {
        /* Only need this relationship object for comment VMLs. */

        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/comments");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, 32, "../comments%d.xml", comment_id);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        self->external_comment_link = relationship;
    }

    /* The vml.c <o:idmap> element data id contains a comma separated range
     * when there is more than one 1024 block of comments, like this:
     * data="1,2,3". Since this could potentially (but unlikely) exceed
     * LXLSX_MAX_ATTRIBUTE_LENGTH we need to allocate space dynamically. */

    /* Calculate the total space required for the ID for each 1024 block. */
    for (i = 0; i <= comment_count / 1024; i++) {
        tmp_data_id = lxlsx_vml_data_id + i;

        /* Calculate the space required for the digits in the id. */
        while (tmp_data_id) {
            data_str_len++;
            tmp_data_id /= 10;
        }

        /* Add an extra char for comma separator or '\O'. */
        data_str_len++;
    };

    /* If this allocation fails it will be dealt with in packager.c. */
    lxlsx_vml_data_id_str = calloc(1, data_str_len + 2);
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_vml_data_id_str, mem_error);

    /* Create the CSV list in the allocated space. */
    for (i = 0; i <= comment_count / 1024; i++) {
        tmp_data_id = lxlsx_vml_data_id + i;
        lxlsx_snprintf(lxlsx_vml_data_id_str + used, data_str_len - used, "%d,",
                     tmp_data_id);

        used = strlen(lxlsx_vml_data_id_str);
    };

    self->lxlsx_vml_shape_id = lxlsx_vml_shape_id;
    self->lxlsx_vml_data_id_str = lxlsx_vml_data_id_str;

    return comment_count;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }

    return 0;
}

/*
 * Set up external linkage for VML header/footer images.
 */
void
lxlsx_worksheet_prepare_header_vml_objects(lxlsx_worksheet *self,
                                         uint32_t lxlsx_vml_header_id,
                                         uint32_t lxlsx_vml_drawing_id)
{

    lxlsx_rel_tuple *relationship;
    char filename[LXLSX_FILENAME_LENGTH];
    char *lxlsx_vml_data_id_str;

    self->lxlsx_vml_header_id = lxlsx_vml_header_id;

    /* Set up the VML relationship for header images. */
    relationship = calloc(1, sizeof(lxlsx_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxlsx_strdup("/vmlDrawing");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxlsx_snprintf(filename, 32, "../drawings/vmlDrawing%d.vml",
                 lxlsx_vml_drawing_id);

    relationship->target = lxlsx_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    self->external_vml_header_link = relationship;

    /* If this allocation fails it will be dealt with in packager.c. */
    lxlsx_vml_data_id_str = calloc(1, sizeof("4294967295"));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_vml_data_id_str, mem_error);

    lxlsx_snprintf(lxlsx_vml_data_id_str, sizeof("4294967295"), "%d", lxlsx_vml_header_id);

    self->lxlsx_vml_header_id_str = lxlsx_vml_data_id_str;

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }

    return;
}

/*
 * Set up external linkage for VML header/footer images.
 */
void
lxlsx_worksheet_prepare_tables(lxlsx_worksheet *self, uint32_t lxlsx_table_id)
{
    lxlsx_table_obj *lxlsx_table_obj;
    lxlsx_rel_tuple *relationship;
    char name[LXLSX_ATTR_32];
    char filename[LXLSX_FILENAME_LENGTH];

    STAILQ_FOREACH(lxlsx_table_obj, self->lxlsx_table_objs, list_pointers) {

        relationship = calloc(1, sizeof(lxlsx_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxlsx_strdup("/table");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxlsx_snprintf(filename, LXLSX_FILENAME_LENGTH,
                     "../tables/table%d.xml", lxlsx_table_id);

        relationship->target = lxlsx_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_table_links, relationship,
                           list_pointers);

        if (!lxlsx_table_obj->name) {
            lxlsx_snprintf(name, LXLSX_ATTR_32, "Table%d", lxlsx_table_id);
            lxlsx_table_obj->name = lxlsx_strdup(name);
            GOTO_LABEL_ON_MEM_ERROR(lxlsx_table_obj->name, mem_error);
        }
        lxlsx_table_obj->id = lxlsx_table_id;
        lxlsx_table_id++;
    }

    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }

    return;
}

/*
 * Extract width and height information from a PNG file.
 */
STATIC lxlsx_error
_process_png(lxlsx_object_properties *object_props)
{
    uint32_t length;
    uint32_t offset;
    char type[4];
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = object_props->stream;

    /* Skip another 4 bytes to the end of the PNG header. */
    fseek_err = fseek(stream, 4, SEEK_CUR);
    if (fseek_err)
        goto file_error;

    while (!feof(stream)) {

        /* Read the PNG length and type fields for the sub-section. */
        if (fread(&length, sizeof(length), 1, stream) < 1)
            break;

        if (fread(&type, 1, 4, stream) < 4)
            break;

        /* Convert the length to network order. */
        length = LXLSX_UINT32_NETWORK(length);

        /* The offset for next fseek() is the field length + type length. */
        offset = length + 4;

        if (memcmp(type, "IHDR", 4) == 0) {
            if (fread(&width, sizeof(width), 1, stream) < 1)
                break;

            if (fread(&height, sizeof(height), 1, stream) < 1)
                break;

            width = LXLSX_UINT32_NETWORK(width);
            height = LXLSX_UINT32_NETWORK(height);

            /* Reduce the offset by the length of previous freads(). */
            offset -= 8;
        }

        if (memcmp(type, "pHYs", 4) == 0) {
            uint32_t x_ppu = 0;
            uint32_t y_ppu = 0;
            uint8_t units = 1;

            if (fread(&x_ppu, sizeof(x_ppu), 1, stream) < 1)
                break;

            if (fread(&y_ppu, sizeof(y_ppu), 1, stream) < 1)
                break;

            if (fread(&units, sizeof(units), 1, stream) < 1)
                break;

            if (units == 1) {
                x_ppu = LXLSX_UINT32_NETWORK(x_ppu);
                y_ppu = LXLSX_UINT32_NETWORK(y_ppu);

                x_dpi = (double) x_ppu *0.0254;
                y_dpi = (double) y_ppu *0.0254;
            }

            /* Reduce the offset by the length of previous freads(). */
            offset -= 9;
        }

        if (memcmp(type, "IEND", 4) == 0)
            break;

        if (!feof(stream)) {
            fseek_err = fseek(stream, offset, SEEK_CUR);
            if (fseek_err)
                goto file_error;
        }
    }

    /* Ensure that we read some valid data from the file. */
    if (width == 0)
        goto file_error;

    /* Set the image metadata. */
    object_props->image_type = LXLSX_IMAGE_PNG;
    object_props->width = width;
    object_props->height = height;
    object_props->x_dpi = x_dpi ? x_dpi : 96;
    object_props->y_dpi = y_dpi ? y_dpi : 96;
    object_props->extension = lxlsx_strdup("png");

    return LXLSX_NO_ERROR;

file_error:
    LXLSX_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", object_props->filename);

    return LXLSX_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a JPEG file.
 */
STATIC lxlsx_error
_process_jpeg(lxlsx_object_properties *image_props)
{
    uint16_t length;
    uint16_t marker;
    uint32_t offset;
    uint16_t width = 0;
    uint16_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_props->stream;

    /* Read back 2 bytes to the end of the initial 0xFFD8 marker. */
    fseek_err = fseek(stream, -2, SEEK_CUR);
    if (fseek_err)
        goto file_error;

    /* Search through the image data and read the JPEG markers. */
    while (!feof(stream)) {

        /* Read the JPEG marker and length fields for the sub-section. */
        if (fread(&marker, sizeof(marker), 1, stream) < 1)
            break;

        if (fread(&length, sizeof(length), 1, stream) < 1)
            break;

        /* Convert the marker and length to network order. */
        marker = LXLSX_UINT16_NETWORK(marker);
        length = LXLSX_UINT16_NETWORK(length);

        /* The offset for next fseek() is the field length + type length. */
        offset = length - 2;

        /* Read the height and width in the 0xFFCn elements (except C4, C8 */
        /* and CC which aren't SOF markers). */
        if ((marker & 0xFFF0) == 0xFFC0 && marker != 0xFFC4
            && marker != 0xFFC8 && marker != 0xFFCC) {
            /* Skip 1 byte to height and width. */
            fseek_err = fseek(stream, 1, SEEK_CUR);
            if (fseek_err)
                goto file_error;

            if (fread(&height, sizeof(height), 1, stream) < 1)
                break;

            if (fread(&width, sizeof(width), 1, stream) < 1)
                break;

            height = LXLSX_UINT16_NETWORK(height);
            width = LXLSX_UINT16_NETWORK(width);

            offset -= 9;
        }

        /* Read the DPI in the 0xFFE0 element. */
        if (marker == 0xFFE0) {
            uint16_t x_density = 0;
            uint16_t y_density = 0;
            uint8_t units = 1;

            fseek_err = fseek(stream, 7, SEEK_CUR);
            if (fseek_err)
                goto file_error;

            if (fread(&units, sizeof(units), 1, stream) < 1)
                break;

            if (fread(&x_density, sizeof(x_density), 1, stream) < 1)
                break;

            if (fread(&y_density, sizeof(y_density), 1, stream) < 1)
                break;

            x_density = LXLSX_UINT16_NETWORK(x_density);
            y_density = LXLSX_UINT16_NETWORK(y_density);

            if (units == 1) {
                x_dpi = x_density;
                y_dpi = y_density;
            }

            if (units == 2) {
                x_dpi = x_density * 2.54;
                y_dpi = y_density * 2.54;
            }

            offset -= 12;
        }

        if (marker == 0xFFDA)
            break;

        if (!feof(stream)) {
            fseek_err = fseek(stream, offset, SEEK_CUR);
            if (fseek_err)
                break;
        }
    }

    /* Ensure that we read some valid data from the file. */
    if (width == 0)
        goto file_error;

    /* Set the image metadata. */
    image_props->image_type = LXLSX_IMAGE_JPEG;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi ? x_dpi : 96;
    image_props->y_dpi = y_dpi ? y_dpi : 96;
    image_props->extension = lxlsx_strdup("jpeg");

    return LXLSX_NO_ERROR;

file_error:
    LXLSX_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXLSX_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a BMP file.
 */
STATIC lxlsx_error
_process_bmp(lxlsx_object_properties *image_props)
{
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_props->stream;

    /* Skip another 14 bytes to the start of the BMP height/width. */
    fseek_err = fseek(stream, 14, SEEK_CUR);
    if (fseek_err)
        goto file_error;

    if (fread(&width, sizeof(width), 1, stream) < 1)
        width = 0;

    if (fread(&height, sizeof(height), 1, stream) < 1)
        height = 0;

    /* Ensure that we read some valid data from the file. */
    if (width == 0)
        goto file_error;

    height = LXLSX_UINT32_HOST(height);
    width = LXLSX_UINT32_HOST(width);

    /* Set the image metadata. */
    image_props->image_type = LXLSX_IMAGE_BMP;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi;
    image_props->y_dpi = y_dpi;
    image_props->extension = lxlsx_strdup("bmp");

    return LXLSX_NO_ERROR;

file_error:
    LXLSX_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXLSX_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a GIF file.
 */
STATIC lxlsx_error
_process_gif(lxlsx_object_properties *image_props)
{
    uint16_t width = 0;
    uint16_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_props->stream;

    /* Skip another 2 bytes to the start of the GIF height/width. */
    fseek_err = fseek(stream, 2, SEEK_CUR);
    if (fseek_err)
        goto file_error;

    if (fread(&width, sizeof(width), 1, stream) < 1)
        width = 0;

    if (fread(&height, sizeof(height), 1, stream) < 1)
        height = 0;

    /* Ensure that we read some valid data from the file. */
    if (width == 0)
        goto file_error;

    height = LXLSX_UINT16_HOST(height);
    width = LXLSX_UINT16_HOST(width);

    /* Set the image metadata. */
    image_props->image_type = LXLSX_IMAGE_GIF;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi;
    image_props->y_dpi = y_dpi;
    image_props->extension = lxlsx_strdup("gif");

    return LXLSX_NO_ERROR;

file_error:
    LXLSX_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXLSX_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract information from the image file such as dimension, type, filename,
 * and extension.
 */
STATIC lxlsx_error
_get_image_properties(lxlsx_object_properties *image_props)
{
    unsigned char signature[4];
#ifndef USE_NO_MD5
    uint8_t i;
    MD5_CTX md5_context;
    size_t size_read;
    char buffer[LXLSX_IMAGE_BUFFER_SIZE];
    unsigned char md5_checksum[LXLSX_MD5_SIZE];
#endif

    /* Read 4 bytes to look for the file header/signature. */
    if (fread(signature, 1, 4, image_props->stream) < 4) {
        LXLSX_WARN_FORMAT1("worksheet image insertion: "
                         "couldn't read image type for: %s.",
                         image_props->filename);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }

    if (memcmp(&signature[1], "PNG", 3) == 0) {
        if (_process_png(image_props) != LXLSX_NO_ERROR)
            return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
    else if (signature[0] == 0xFF && signature[1] == 0xD8) {
        if (_process_jpeg(image_props) != LXLSX_NO_ERROR)
            return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "BM", 2) == 0) {
        if (_process_bmp(image_props) != LXLSX_NO_ERROR)
            return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "GIF8", 4) == 0) {
        if (_process_gif(image_props) != LXLSX_NO_ERROR)
            return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
    else {
        LXLSX_WARN_FORMAT1("worksheet image insertion: "
                         "unsupported image format for: %s.",
                         image_props->filename);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }

#ifndef USE_NO_MD5
    /* Calculate an MD5 checksum for the image so that we can remove duplicate
     * images to reduce the xlsx file size.*/
    rewind(image_props->stream);

    MD5_Init(&md5_context);

    size_read = fread(buffer, 1, LXLSX_IMAGE_BUFFER_SIZE, image_props->stream);
    while (size_read) {
        MD5_Update(&md5_context, buffer, (unsigned long) size_read);
        size_read =
            fread(buffer, 1, LXLSX_IMAGE_BUFFER_SIZE, image_props->stream);
    }

    MD5_Final(md5_checksum, &md5_context);

    /* Create a 32 char hex string buffer for the MD5 checksum. */
    image_props->md5 = calloc(1, LXLSX_MD5_SIZE * 2 + 1);

    /* If this calloc fails we just return and don't remove duplicates. */
    RETURN_ON_MEM_ERROR(image_props->md5, LXLSX_NO_ERROR);

    /* Convert the 16 byte MD5 buffer to a 32 char hex string. */
    for (i = 0; i < LXLSX_MD5_SIZE; i++) {
        lxlsx_snprintf(&image_props->md5[2 * i], 3, "%02x", md5_checksum[i]);
    }
#endif

    return LXLSX_NO_ERROR;
}

/* Conditional formats that refer to the same cell sqref range, like A or
 * B1:B9, need to be written as part of one xml structure. Therefore we need
 * to store them in a RB hash/tree keyed by sqref. Within the RB hash element
 * we then store conditional formats that refer to sqref in a STAILQ list. */
lxlsx_error
_store_conditional_format_object(lxlsx_worksheet *self,
                                 lxlsx_cond_format_obj *cond_format)
{
    lxlsx_cond_format_hash_element tmp_hash_element;
    lxlsx_cond_format_hash_element *found_hash_element = NULL;
    lxlsx_cond_format_hash_element *new_hash_element = NULL;

    /* Create a temp hash element to do the lookup. */
    LXLSX_ATTRIBUTE_COPY(tmp_hash_element.sqref, cond_format->sqref);
    found_hash_element = RB_FIND(lxlsx_cond_format_hash,
                                 self->conditional_formats,
                                 &tmp_hash_element);

    if (found_hash_element) {
        /* If the RB element exists then add the conditional format to the
         * list for the sqref range.*/
        STAILQ_INSERT_TAIL(found_hash_element->cond_formats, cond_format,
                           list_pointers);
    }
    else {
        /* Create a new RB hash element. */
        new_hash_element = calloc(1, sizeof(lxlsx_cond_format_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(new_hash_element, mem_error);

        /* Use the sqref as the key. */
        LXLSX_ATTRIBUTE_COPY(new_hash_element->sqref, cond_format->sqref);

        /* Also create the list where we store the cond format objects. */
        new_hash_element->cond_formats =
            calloc(1, sizeof(struct lxlsx_cond_format_list));
        GOTO_LABEL_ON_MEM_ERROR(new_hash_element->cond_formats, mem_error);

        /* Initialize the list and add the conditional format object. */
        STAILQ_INIT(new_hash_element->cond_formats);
        STAILQ_INSERT_TAIL(new_hash_element->cond_formats, cond_format,
                           list_pointers);

        /* Now insert the RB hash element into the tree. */
        RB_INSERT(lxlsx_cond_format_hash, self->conditional_formats,
                  new_hash_element);

    }

    return LXLSX_NO_ERROR;

mem_error:
    free(new_hash_element);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write out a number worksheet cell. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
STATIC void
_write_number_cell(lxlsx_worksheet *self, char *range,
                   int32_t style_index, lxlsx_cell *cell)
{
#ifdef USE_DTOA_LIBRARY
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, cell->u.number);

    if (style_index)
        fprintf(self->file,
                "<c r=\"%s\" s=\"%d\"><v>%s</v></c>",
                range, style_index, data);
    else
        fprintf(self->file, "<c r=\"%s\"><v>%s</v></c>", range, data);
#else
    if (style_index)
        fprintf(self->file,
                "<c r=\"%s\" s=\"%d\"><v>%.16G</v></c>",
                range, style_index, cell->u.number);
    else
        fprintf(self->file,
                "<c r=\"%s\"><v>%.16G</v></c>", range, cell->u.number);

#endif
}

/*
 * Write out a string worksheet cell. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
STATIC void
_write_string_cell(lxlsx_worksheet *self, char *range,
                   int32_t style_index, lxlsx_cell *cell)
{

    if (style_index)
        fprintf(self->file,
                "<c r=\"%s\" s=\"%d\" t=\"s\"><v>%d</v></c>",
                range, style_index, cell->u.string_id);
    else
        fprintf(self->file,
                "<c r=\"%s\" t=\"s\"><v>%d</v></c>",
                range, cell->u.string_id);
}

/*
 * Write out an inline string. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
STATIC void
_write_inline_string_cell(lxlsx_worksheet *self, char *range,
                          int32_t style_index, lxlsx_cell *cell)
{
    char *string = lxlsx_escape_data(cell->u.string);

    /* Add attribute to preserve leading or trailing whitespace. */
    if (isspace((unsigned char) string[0])
        || isspace((unsigned char) string[strlen(string) - 1])) {

        if (style_index)
            fprintf(self->file,
                    "<c r=\"%s\" s=\"%d\" t=\"inlineStr\"><is>"
                    "<t xml:space=\"preserve\">%s</t></is></c>",
                    range, style_index, string);
        else
            fprintf(self->file,
                    "<c r=\"%s\" t=\"inlineStr\"><is>"
                    "<t xml:space=\"preserve\">%s</t></is></c>",
                    range, string);
    }
    else {
        if (style_index)
            fprintf(self->file,
                    "<c r=\"%s\" s=\"%d\" t=\"inlineStr\">"
                    "<is><t>%s</t></is></c>", range, style_index, string);
        else
            fprintf(self->file,
                    "<c r=\"%s\" t=\"inlineStr\">"
                    "<is><t>%s</t></is></c>", range, string);
    }

    free(string);
}

/*
 * Write out an inline rich string. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
STATIC void
_write_inline_rich_string_cell(lxlsx_worksheet *self, char *range,
                               int32_t style_index, lxlsx_cell *cell)
{
    const char *string = cell->u.string;

    if (style_index)
        fprintf(self->file,
                "<c r=\"%s\" s=\"%d\" t=\"inlineStr\">"
                "<is>%s</is></c>", range, style_index, string);
    else
        fprintf(self->file,
                "<c r=\"%s\" t=\"inlineStr\">"
                "<is>%s</is></c>", range, string);
}

/*
 * Write out a formula worksheet cell with a numeric result.
 */
STATIC void
_write_formula_num_cell(lxlsx_worksheet *self, lxlsx_cell *cell)
{
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, cell->formula_result);
    lxlsx_xml_data_element(self->file, "f", cell->u.string, NULL);
    lxlsx_xml_data_element(self->file, "v", data, NULL);
}

/*
 * Write out a formula worksheet cell with a numeric result.
 */
STATIC void
_write_formula_str_cell(lxlsx_worksheet *self, lxlsx_cell *cell)
{
    lxlsx_xml_data_element(self->file, "f", cell->u.string, NULL);
    lxlsx_xml_data_element(self->file, "v", cell->user_data2, NULL);
}

/*
 * Write out an array formula worksheet cell with a numeric result.
 */
STATIC void
_write_array_formula_num_cell(lxlsx_worksheet *self, lxlsx_cell *cell)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char data[LXLSX_ATTR_32];

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("t", "array");
    LXLSX_PUSH_ATTRIBUTES_STR("ref", cell->user_data1);

    lxlsx_sprintf_dbl(data, cell->formula_result);

    lxlsx_xml_data_element(self->file, "f", cell->u.string, &attributes);
    lxlsx_xml_data_element(self->file, "v", data, NULL);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write out a boolean worksheet cell.
 */
STATIC void
_write_boolean_cell(lxlsx_worksheet *self, lxlsx_cell *cell)
{
    char data[LXLSX_ATTR_32];

    if (cell->u.number == 0.0)
        data[0] = '0';
    else
        data[0] = '1';

    data[1] = '\0';

    lxlsx_xml_data_element(self->file, "v", data, NULL);
}

/*
 * Write out a error worksheet cell.
 */
STATIC void
_write_error_cell(lxlsx_worksheet *self)
{
    lxlsx_xml_data_element(self->file, "v", "#VALUE!", NULL);
}

/*
 * Calculate the "spans" attribute of the <row> tag. This is an XLSX
 * optimization and isn't strictly required. However, it makes comparing
 * files easier.
 *
 * The span is the same for each block of 16 rows.
 */
STATIC void
_calculate_spans(struct lxlsx_row *row, char *span, int32_t *block_num)
{
    lxlsx_cell *cell_min = RB_MIN(lxlsx_table_cells, row->cells);
    lxlsx_cell *cell_max = RB_MAX(lxlsx_table_cells, row->cells);
    lxlsx_col_t span_col_min = cell_min->col_num;
    lxlsx_col_t span_col_max = cell_max->col_num;
    lxlsx_col_t col_min;
    lxlsx_col_t col_max;
    *block_num = row->row_num / 16;

    row = RB_NEXT(lxlsx_table_rows, root, row);

    while (row && (int32_t) (row->row_num / 16) == *block_num) {

        if (!RB_EMPTY(row->cells)) {
            cell_min = RB_MIN(lxlsx_table_cells, row->cells);
            cell_max = RB_MAX(lxlsx_table_cells, row->cells);
            col_min = cell_min->col_num;
            col_max = cell_max->col_num;

            if (col_min < span_col_min)
                span_col_min = col_min;

            if (col_max > span_col_max)
                span_col_max = col_max;
        }

        row = RB_NEXT(lxlsx_table_rows, root, row);
    }

    lxlsx_snprintf(span, LXLSX_MAX_CELL_RANGE_LENGTH,
                 "%d:%d", span_col_min + 1, span_col_max + 1);
}

/*
 * Write out a generic worksheet cell.
 */
STATIC void
_write_cell(lxlsx_worksheet *self, lxlsx_cell *cell, lxlsx_format *row_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char range[LXLSX_MAX_CELL_NAME_LENGTH] = { 0 };
    lxlsx_row_t row_num = cell->row_num;
    lxlsx_col_t col_num = cell->col_num;
    int32_t style_index = 0;

    lxlsx_rowcol_to_cell(range, row_num, col_num);

    if (cell->format) {
        style_index = lxlsx_format_get_xf_index(cell->format);
    }
    else if (row_format) {
        style_index = lxlsx_format_get_xf_index(row_format);
    }
    else if (col_num < self->col_formats_max && self->col_formats[col_num]) {
        style_index = lxlsx_format_get_xf_index(self->col_formats[col_num]);
    }

    /* Unrolled optimization for most commonly written cell types. */
    if (cell->type == NUMBER_CELL) {
        _write_number_cell(self, range, style_index, cell);
        return;
    }

    if (cell->type == STRING_CELL) {
        _write_string_cell(self, range, style_index, cell);
        return;
    }

    if (cell->type == INLINE_STRING_CELL) {
        _write_inline_string_cell(self, range, style_index, cell);
        return;
    }

    if (cell->type == INLINE_RICH_STRING_CELL) {
        _write_inline_rich_string_cell(self, range, style_index, cell);
        return;
    }

    /* For other cell types use the general functions. */
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r", range);

    if (style_index)
        LXLSX_PUSH_ATTRIBUTES_INT("s", style_index);

    if (cell->type == FORMULA_CELL) {
        /* If user_data2 is set then the formula has a string result. */
        if (cell->user_data2)
            LXLSX_PUSH_ATTRIBUTES_STR("t", "str");

        lxlsx_xml_start_tag(self->file, "c", &attributes);

        if (cell->user_data2)
            _write_formula_str_cell(self, cell);
        else
            _write_formula_num_cell(self, cell);

        lxlsx_xml_end_tag(self->file, "c");
    }
    else if (cell->type == BLANK_CELL) {
        if (cell->format)
            lxlsx_xml_empty_tag(self->file, "c", &attributes);
    }
    else if (cell->type == BOOLEAN_CELL) {
        LXLSX_PUSH_ATTRIBUTES_STR("t", "b");
        lxlsx_xml_start_tag(self->file, "c", &attributes);
        _write_boolean_cell(self, cell);
        lxlsx_xml_end_tag(self->file, "c");
    }
    else if (cell->type == ARRAY_FORMULA_CELL) {
        lxlsx_xml_start_tag(self->file, "c", &attributes);
        _write_array_formula_num_cell(self, cell);
        lxlsx_xml_end_tag(self->file, "c");
    }
    else if (cell->type == DYNAMIC_ARRAY_FORMULA_CELL) {
        LXLSX_PUSH_ATTRIBUTES_STR("cm", "1");
        lxlsx_xml_start_tag(self->file, "c", &attributes);
        _write_array_formula_num_cell(self, cell);
        lxlsx_xml_end_tag(self->file, "c");
    }
    else if (cell->type == ERROR_CELL) {
        LXLSX_PUSH_ATTRIBUTES_STR("t", "e");
        LXLSX_PUSH_ATTRIBUTES_DBL("vm", cell->u.number);
        lxlsx_xml_start_tag(self->file, "c", &attributes);
        _write_error_cell(self);
        lxlsx_xml_end_tag(self->file, "c");
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write out the worksheet data as a series of rows and cells.
 */
STATIC void
_worksheet_write_rows(lxlsx_worksheet *self)
{
    lxlsx_row *row;
    lxlsx_cell *cell;
    int32_t block_num = -1;
    char spans[LXLSX_MAX_CELL_RANGE_LENGTH] = { 0 };

    RB_FOREACH(row, lxlsx_table_rows, self->table) {

        if (RB_EMPTY(row->cells)) {
            /* Row contains no cells but has height, format or other data. */

            /* Write a default span for default rows. */
            if (self->default_row_set)
                _write_row(self, row, "1:1");
            else
                _write_row(self, row, NULL);
        }
        else {
            /* Row and cell data. */
            if ((int32_t) row->row_num / 16 > block_num)
                _calculate_spans(row, spans, &block_num);

            _write_row(self, row, spans);

            if (row->data_changed) {
                RB_FOREACH(cell, lxlsx_table_cells, row->cells) {
                    _write_cell(self, cell, row->format);
                }

                lxlsx_xml_end_tag(self->file, "row");
            }
        }
    }
}

/*
 * Write out the worksheet data as a single row with cells. This method is
 * used when memory optimization is on. A single row is written and the data
 * array is reset. That way only one row of data is kept in memory at any one
 * time. We don't write span data in the optimized case since it is optional.
 */
void
lxlsx_worksheet_write_single_row(lxlsx_worksheet *self)
{
    lxlsx_row *row = self->optimize_row;
    lxlsx_col_t col;

    /* skip row if it doesn't contain row formatting, cell data or a comment. */
    if (!(row->row_changed || row->data_changed))
        return;

    /* Write the cells if the row contains data. */
    if (!row->data_changed) {
        /* Row data only. No cells. */
        _write_row(self, row, NULL);
    }
    else {
        /* Row and cell data. */
        _write_row(self, row, NULL);

        for (col = self->dim_colmin; col <= self->dim_colmax; col++) {
            if (self->array[col]) {
                _write_cell(self, self->array[col], row->format);
                _free_cell(self->array[col]);
                self->array[col] = NULL;
            }
        }

        lxlsx_xml_end_tag(self->file, "row");
    }

    /* Reset the row. */
    row->height = LXLSX_DEF_ROW_HEIGHT;
    row->format = NULL;
    row->hidden = LXLSX_FALSE;
    row->level = 0;
    row->collapsed = LXLSX_FALSE;
    row->data_changed = LXLSX_FALSE;
    row->row_changed = LXLSX_FALSE;
}

/* Process a header/footer image and store it in the correct slot. */
lxlsx_error
_worksheet_set_header_footer_image(lxlsx_worksheet *self, const char *filename,
                                   uint8_t image_position)
{
    FILE *image_stream;
    const char *description;
    lxlsx_object_properties *object_props;
    char *image_strings[] = { "LH", "CH", "RH", "LF", "CF", "RF" };

    /* Not all slots will have image files. */
    if (!filename)
        return LXLSX_NO_ERROR;

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(filename, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Use the filename as the default description, like Excel. */
    description = lxlsx_basename(filename);
    if (!description) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                         "couldn't get basename for file: %s.", filename);
        fclose(image_stream);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup(filename);
    object_props->description = lxlsx_strdup(description);
    object_props->stream = image_stream;

    /* Set VML image position string based on the header/footer/position. */
    object_props->image_position = lxlsx_strdup(image_strings[image_position]);

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        *self->header_footer_objs[image_position] = object_props;
        self->has_header_vml = LXLSX_TRUE;
        fclose(image_stream);
        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Write the <col> element.
 */
STATIC void
_worksheet_write_col_info(lxlsx_worksheet *self, lxlsx_col_options *options)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    double width = options->width;
    uint8_t has_custom_width = LXLSX_TRUE;
    int32_t xf_index = 0;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;

    /* Get the format index. */
    if (options->format) {
        xf_index = lxlsx_format_get_xf_index(options->format);
    }

    /* Check if width is the Excel default. */
    if (width == LXLSX_DEF_COL_WIDTH) {

        /* The default col width changes to 0 for hidden columns. */
        if (options->hidden)
            width = 0;
        else
            has_custom_width = LXLSX_FALSE;

    }

    /* Convert column width from user units to character width. */
    if (width > 0) {
        if (width < 1) {
            width = (uint16_t) (((uint16_t)
                                 (width * (max_digit_width + padding) + 0.5))
                                / max_digit_width * 256.0) / 256.0;
        }
        else {
            width = (uint16_t) (((uint16_t)
                                 (width * max_digit_width + 0.5) + padding)
                                / max_digit_width * 256.0) / 256.0;
        }
    }

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("min", 1 + options->firstcol);
    LXLSX_PUSH_ATTRIBUTES_INT("max", 1 + options->lastcol);
    LXLSX_PUSH_ATTRIBUTES_DBL("width", width);

    if (xf_index)
        LXLSX_PUSH_ATTRIBUTES_INT("style", xf_index);

    if (options->hidden)
        LXLSX_PUSH_ATTRIBUTES_STR("hidden", "1");

    if (has_custom_width)
        LXLSX_PUSH_ATTRIBUTES_STR("customWidth", "1");

    if (options->level)
        LXLSX_PUSH_ATTRIBUTES_INT("outlineLevel", options->level);

    if (options->collapsed)
        LXLSX_PUSH_ATTRIBUTES_STR("collapsed", "1");

    lxlsx_xml_empty_tag(self->file, "col", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cols> element and <col> sub elements.
 */
STATIC void
_worksheet_write_cols(lxlsx_worksheet *self)
{
    lxlsx_col_t col;

    if (!self->col_size_changed)
        return;

    lxlsx_xml_start_tag(self->file, "cols", NULL);

    for (col = 0; col < self->col_options_max; col++) {
        if (self->col_options[col])
            _worksheet_write_col_info(self, self->col_options[col]);
    }

    lxlsx_xml_end_tag(self->file, "cols");
}

/*
 * Write the <mergeCell> element.
 */
STATIC void
_worksheet_write_merge_cell(lxlsx_worksheet *self,
                            lxlsx_merged_range *merged_range)
{

    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char ref[LXLSX_MAX_CELL_RANGE_LENGTH];

    LXLSX_INIT_ATTRIBUTES();

    /* Convert the merge dimensions to a cell range. */
    lxlsx_rowcol_to_range(ref, merged_range->first_row, merged_range->first_col,
                        merged_range->last_row, merged_range->last_col);

    LXLSX_PUSH_ATTRIBUTES_STR("ref", ref);

    lxlsx_xml_empty_tag(self->file, "mergeCell", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <mergeCells> element.
 */
STATIC void
_worksheet_write_merge_cells(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_merged_range *merged_range;

    if (self->merged_range_count) {
        LXLSX_INIT_ATTRIBUTES();

        LXLSX_PUSH_ATTRIBUTES_INT("count", self->merged_range_count);

        lxlsx_xml_start_tag(self->file, "mergeCells", &attributes);

        STAILQ_FOREACH(merged_range, self->merged_ranges, list_pointers) {
            _worksheet_write_merge_cell(self, merged_range);
        }
        lxlsx_xml_end_tag(self->file, "mergeCells");

        LXLSX_FREE_ATTRIBUTES();
    }
}

/*
 * Write the <oddHeader> element.
 */
STATIC void
_worksheet_write_odd_header(lxlsx_worksheet *self)
{
    lxlsx_xml_data_element(self->file, "oddHeader", self->header, NULL);
}

/*
 * Write the <oddFooter> element.
 */
STATIC void
_worksheet_write_odd_footer(lxlsx_worksheet *self)
{
    lxlsx_xml_data_element(self->file, "oddFooter", self->footer, NULL);
}

/*
 * Write the <headerFooter> element.
 */
STATIC void
_worksheet_write_header_footer(lxlsx_worksheet *self)
{
    if (!self->header_footer_changed)
        return;

    lxlsx_xml_start_tag(self->file, "headerFooter", NULL);

    if (self->header)
        _worksheet_write_odd_header(self);

    if (self->footer)
        _worksheet_write_odd_footer(self);

    lxlsx_xml_end_tag(self->file, "headerFooter");
}

/*
 * Write the <pageSetUpPr> element.
 */
STATIC void
_worksheet_write_page_set_up_pr(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (!self->fit_page)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("fitToPage", "1");

    lxlsx_xml_empty_tag(self->file, "pageSetUpPr", &attributes);

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the <tabColor> element.
 */
STATIC void
_worksheet_write_tab_color(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb_str[LXLSX_ATTR_32];

    if (self->tab_color == LXLSX_COLOR_UNSET)
        return;

    lxlsx_snprintf(rgb_str, LXLSX_ATTR_32, "FF%06X",
                 self->tab_color & LXLSX_COLOR_MASK);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

    lxlsx_xml_empty_tag(self->file, "tabColor", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <outlinePr> element.
 */
STATIC void
_worksheet_write_outline_pr(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (!self->outline_changed)
        return;

    LXLSX_INIT_ATTRIBUTES();

    if (self->outline_style)
        LXLSX_PUSH_ATTRIBUTES_STR("applyStyles", "1");

    if (!self->outline_below)
        LXLSX_PUSH_ATTRIBUTES_STR("summaryBelow", "0");

    if (!self->outline_right)
        LXLSX_PUSH_ATTRIBUTES_STR("summaryRight", "0");

    if (!self->outline_on)
        LXLSX_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

    lxlsx_xml_empty_tag(self->file, "outlinePr", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetPr> element for Sheet level properties.
 */
STATIC void
_worksheet_write_sheet_pr(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (!self->fit_page
        && !self->filter_on
        && self->tab_color == LXLSX_COLOR_UNSET
        && !self->outline_changed
        && !self->vba_codename && !self->is_chartsheet) {
        return;
    }

    LXLSX_INIT_ATTRIBUTES();

    if (self->vba_codename)
        LXLSX_PUSH_ATTRIBUTES_STR("codeName", self->vba_codename);

    if (self->filter_on)
        LXLSX_PUSH_ATTRIBUTES_STR("filterMode", "1");

    if (self->fit_page || self->tab_color != LXLSX_COLOR_UNSET
        || self->outline_changed) {
        lxlsx_xml_start_tag(self->file, "sheetPr", &attributes);
        _worksheet_write_tab_color(self);
        _worksheet_write_outline_pr(self);
        _worksheet_write_page_set_up_pr(self);
        lxlsx_xml_end_tag(self->file, "sheetPr");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "sheetPr", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the <brk> element.
 */
STATIC void
_worksheet_write_brk(lxlsx_worksheet *self, uint32_t id, uint32_t max)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("id", id);
    LXLSX_PUSH_ATTRIBUTES_INT("max", max);
    LXLSX_PUSH_ATTRIBUTES_STR("man", "1");

    lxlsx_xml_empty_tag(self->file, "brk", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <rowBreaks> element.
 */
STATIC void
_worksheet_write_row_breaks(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint16_t count = self->hbreaks_count;
    uint16_t i;

    if (!count)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", count);
    LXLSX_PUSH_ATTRIBUTES_INT("manualBreakCount", count);

    lxlsx_xml_start_tag(self->file, "rowBreaks", &attributes);

    for (i = 0; i < count; i++)
        _worksheet_write_brk(self, self->hbreaks[i], LXLSX_COL_MAX - 1);

    lxlsx_xml_end_tag(self->file, "rowBreaks");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <colBreaks> element.
 */
STATIC void
_worksheet_write_col_breaks(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint16_t count = self->vbreaks_count;
    uint16_t i;

    if (!count)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", count);
    LXLSX_PUSH_ATTRIBUTES_INT("manualBreakCount", count);

    lxlsx_xml_start_tag(self->file, "colBreaks", &attributes);

    for (i = 0; i < count; i++)
        _worksheet_write_brk(self, self->vbreaks[i], LXLSX_ROW_MAX - 1);

    lxlsx_xml_end_tag(self->file, "colBreaks");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <filter> element.
 */
STATIC void
_worksheet_write_filter(lxlsx_worksheet *self, const char *str, double num,
                        uint8_t criteria)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (criteria == LXLSX_FILTER_CRITERIA_BLANKS)
        return;

    LXLSX_INIT_ATTRIBUTES();

    if (str)
        LXLSX_PUSH_ATTRIBUTES_STR("val", str);
    else
        LXLSX_PUSH_ATTRIBUTES_DBL("val", num);

    lxlsx_xml_empty_tag(self->file, "filter", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element as simple equality.
 */
STATIC void
_worksheet_write_filter_standard(lxlsx_worksheet *self,
                                 lxlsx_filter_rule_obj *filter)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (filter->has_blanks) {
        LXLSX_PUSH_ATTRIBUTES_STR("blank", "1");
    }

    if (filter->type == LXLSX_FILTER_TYPE_SINGLE && filter->has_blanks) {
        lxlsx_xml_empty_tag(self->file, "filters", &attributes);

    }
    else {
        lxlsx_xml_start_tag(self->file, "filters", &attributes);

        /* Write the filter element. */
        if (filter->type == LXLSX_FILTER_TYPE_SINGLE) {
            _worksheet_write_filter(self, filter->value1_string,
                                    filter->value1, filter->criteria1);
        }
        else if (filter->type == LXLSX_FILTER_TYPE_AND
                 || filter->type == LXLSX_FILTER_TYPE_OR) {
            _worksheet_write_filter(self, filter->value1_string,
                                    filter->value1, filter->criteria1);
            _worksheet_write_filter(self, filter->value2_string,
                                    filter->value2, filter->criteria2);
        }

        lxlsx_xml_end_tag(self->file, "filters");
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <customFilter> element.
 */
STATIC void
_worksheet_write_custom_filter(lxlsx_worksheet *self, const char *str,
                               double num, uint8_t criteria)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (criteria == LXLSX_FILTER_CRITERIA_NOT_EQUAL_TO)
        LXLSX_PUSH_ATTRIBUTES_STR("operator", "notEqual");
    if (criteria == LXLSX_FILTER_CRITERIA_GREATER_THAN)
        LXLSX_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
    else if (criteria == LXLSX_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO)
        LXLSX_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
    else if (criteria == LXLSX_FILTER_CRITERIA_LESS_THAN)
        LXLSX_PUSH_ATTRIBUTES_STR("operator", "lessThan");
    else if (criteria == LXLSX_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO)
        LXLSX_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");

    if (str)
        LXLSX_PUSH_ATTRIBUTES_STR("val", str);
    else
        LXLSX_PUSH_ATTRIBUTES_DBL("val", num);

    lxlsx_xml_empty_tag(self->file, "customFilter", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element as a list.
 */
STATIC void
_worksheet_write_filter_list(lxlsx_worksheet *self, lxlsx_filter_rule_obj *filter)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint16_t i;

    LXLSX_INIT_ATTRIBUTES();

    if (filter->has_blanks) {
        LXLSX_PUSH_ATTRIBUTES_STR("blank", "1");
    }

    lxlsx_xml_start_tag(self->file, "filters", &attributes);

    for (i = 0; i < filter->num_list_filters; i++) {
        _worksheet_write_filter(self, filter->list[i], 0, 0);
    }

    lxlsx_xml_end_tag(self->file, "filters");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element.
 */
STATIC void
_worksheet_write_filter_custom(lxlsx_worksheet *self,
                               lxlsx_filter_rule_obj *filter)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (filter->type == LXLSX_FILTER_TYPE_AND)
        LXLSX_PUSH_ATTRIBUTES_STR("and", "1");

    lxlsx_xml_start_tag(self->file, "customFilters", &attributes);

    /* Write the filter element. */
    if (filter->type == LXLSX_FILTER_TYPE_SINGLE) {
        _worksheet_write_custom_filter(self, filter->value1_string,
                                       filter->value1, filter->criteria1);
    }
    else if (filter->type == LXLSX_FILTER_TYPE_AND
             || filter->type == LXLSX_FILTER_TYPE_OR) {
        _worksheet_write_custom_filter(self, filter->value1_string,
                                       filter->value1, filter->criteria1);
        _worksheet_write_custom_filter(self, filter->value2_string,
                                       filter->value2, filter->criteria2);
    }

    lxlsx_xml_end_tag(self->file, "customFilters");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element.
 */
STATIC void
_worksheet_write_filter_column(lxlsx_worksheet *self,
                               lxlsx_filter_rule_obj *filter)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (!filter)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("colId", filter->col_num);

    lxlsx_xml_start_tag(self->file, "filterColumn", &attributes);

    if (filter->list)
        _worksheet_write_filter_list(self, filter);
    else if (filter->is_custom)
        _worksheet_write_filter_custom(self, filter);
    else
        _worksheet_write_filter_standard(self, filter);

    lxlsx_xml_end_tag(self->file, "filterColumn");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <autoFilter> element.
 */
STATIC void
_worksheet_write_auto_filter(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char range[LXLSX_MAX_CELL_RANGE_LENGTH];
    uint16_t i;

    if (!self->autofilter.in_use)
        return;

    lxlsx_rowcol_to_range(range,
                        self->autofilter.first_row,
                        self->autofilter.first_col,
                        self->autofilter.last_row, self->autofilter.last_col);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", range);

    if (self->autofilter.has_rules) {
        lxlsx_xml_start_tag(self->file, "autoFilter", &attributes);

        for (i = 0; i < self->num_filter_rules; i++)
            _worksheet_write_filter_column(self, self->filter_rules[i]);

        lxlsx_xml_end_tag(self->file, "autoFilter");

    }
    else {
        lxlsx_xml_empty_tag(self->file, "autoFilter", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <hyperlink> element for external links.
 */
STATIC void
_worksheet_write_hyperlink_external(lxlsx_worksheet *self, lxlsx_row_t row_num,
                                    lxlsx_col_t col_num, const char *location,
                                    const char *tooltip, uint16_t id)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char ref[LXLSX_MAX_CELL_NAME_LENGTH];
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_rowcol_to_cell(ref, row_num, col_num);

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", id);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", ref);
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    if (location)
        LXLSX_PUSH_ATTRIBUTES_STR("location", location);

    if (tooltip)
        LXLSX_PUSH_ATTRIBUTES_STR("tooltip", tooltip);

    lxlsx_xml_empty_tag(self->file, "hyperlink", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <hyperlink> element for internal links.
 */
STATIC void
_worksheet_write_hyperlink_internal(lxlsx_worksheet *self, lxlsx_row_t row_num,
                                    lxlsx_col_t col_num, const char *location,
                                    const char *display, const char *tooltip)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char ref[LXLSX_MAX_CELL_NAME_LENGTH];

    lxlsx_rowcol_to_cell(ref, row_num, col_num);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", ref);

    if (location)
        LXLSX_PUSH_ATTRIBUTES_STR("location", location);

    if (tooltip)
        LXLSX_PUSH_ATTRIBUTES_STR("tooltip", tooltip);

    if (display)
        LXLSX_PUSH_ATTRIBUTES_STR("display", display);

    lxlsx_xml_empty_tag(self->file, "hyperlink", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Process any stored hyperlinks in row/col order and write the <hyperlinks>
 * element. The attributes are different for internal and external links.
 */
STATIC void
_worksheet_write_hyperlinks(lxlsx_worksheet *self)
{

    lxlsx_row *row;
    lxlsx_cell *link;
    lxlsx_rel_tuple *relationship;

    if (RB_EMPTY(self->hyperlinks))
        return;

    /* Write the hyperlink elements. */
    lxlsx_xml_start_tag(self->file, "hyperlinks", NULL);

    RB_FOREACH(row, lxlsx_table_rows, self->hyperlinks) {

        RB_FOREACH(link, lxlsx_table_cells, row->cells) {

            if (link->type == HYPERLINK_URL
                || link->type == HYPERLINK_EXTERNAL) {

                self->rel_count++;

                relationship = calloc(1, sizeof(lxlsx_rel_tuple));
                GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

                relationship->type = lxlsx_strdup("/hyperlink");
                GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

                relationship->target = lxlsx_strdup(link->u.string);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                relationship->target_mode = lxlsx_strdup("External");
                GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);

                STAILQ_INSERT_TAIL(self->external_hyperlinks, relationship,
                                   list_pointers);

                _worksheet_write_hyperlink_external(self, link->row_num,
                                                    link->col_num,
                                                    link->user_data1,
                                                    link->user_data2,
                                                    self->rel_count);
            }

            if (link->type == HYPERLINK_INTERNAL) {

                _worksheet_write_hyperlink_internal(self, link->row_num,
                                                    link->col_num,
                                                    link->u.string,
                                                    link->user_data1,
                                                    link->user_data2);
            }

        }

    }

    lxlsx_xml_end_tag(self->file, "hyperlinks");
    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
    lxlsx_xml_end_tag(self->file, "hyperlinks");
}

/*
 * Write the <sheetProtection> element.
 */
STATIC void
_worksheet_write_sheet_protection(lxlsx_worksheet *self,
                                  lxlsx_protection_obj *protect)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (!protect->is_configured)
        return;

    LXLSX_INIT_ATTRIBUTES();

    if (*protect->hash)
        LXLSX_PUSH_ATTRIBUTES_STR("password", protect->hash);

    if (!protect->no_sheet)
        LXLSX_PUSH_ATTRIBUTES_INT("sheet", 1);

    if (!protect->no_content)
        LXLSX_PUSH_ATTRIBUTES_INT("content", 1);

    if (!protect->objects)
        LXLSX_PUSH_ATTRIBUTES_INT("objects", 1);

    if (!protect->scenarios)
        LXLSX_PUSH_ATTRIBUTES_INT("scenarios", 1);

    if (protect->lxlsx_format_cells)
        LXLSX_PUSH_ATTRIBUTES_INT("formatCells", 0);

    if (protect->lxlsx_format_columns)
        LXLSX_PUSH_ATTRIBUTES_INT("formatColumns", 0);

    if (protect->lxlsx_format_rows)
        LXLSX_PUSH_ATTRIBUTES_INT("formatRows", 0);

    if (protect->insert_columns)
        LXLSX_PUSH_ATTRIBUTES_INT("insertColumns", 0);

    if (protect->insert_rows)
        LXLSX_PUSH_ATTRIBUTES_INT("insertRows", 0);

    if (protect->insert_hyperlinks)
        LXLSX_PUSH_ATTRIBUTES_INT("insertHyperlinks", 0);

    if (protect->delete_columns)
        LXLSX_PUSH_ATTRIBUTES_INT("deleteColumns", 0);

    if (protect->delete_rows)
        LXLSX_PUSH_ATTRIBUTES_INT("deleteRows", 0);

    if (protect->no_select_locked_cells)
        LXLSX_PUSH_ATTRIBUTES_INT("selectLockedCells", 1);

    if (protect->sort)
        LXLSX_PUSH_ATTRIBUTES_INT("sort", 0);

    if (protect->autofilter)
        LXLSX_PUSH_ATTRIBUTES_INT("autoFilter", 0);

    if (protect->pivot_tables)
        LXLSX_PUSH_ATTRIBUTES_INT("pivotTables", 0);

    if (protect->no_select_unlocked_cells)
        LXLSX_PUSH_ATTRIBUTES_INT("selectUnlockedCells", 1);

    lxlsx_xml_empty_tag(self->file, "sheetProtection", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <legacyDrawing> element.
 */
STATIC void
_worksheet_write_legacy_drawing(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_vml)
        return;
    else
        self->rel_count++;

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", self->rel_count);
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "legacyDrawing", &attributes);

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the <legacyDrawingHF> element.
 */
STATIC void
_worksheet_write_legacy_drawing_hf(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_header_vml)
        return;
    else
        self->rel_count++;

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", self->rel_count);
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "legacyDrawingHF", &attributes);

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the <picture> element.
 */
STATIC void
_worksheet_write_picture(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_background_image)
        return;
    else
        self->rel_count++;

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", self->rel_count);
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "picture", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <drawing> element.
 */
STATIC void
_worksheet_write_drawing(lxlsx_worksheet *self, uint16_t id)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", id);
    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "drawing", &attributes);

    LXLSX_FREE_ATTRIBUTES();

}

/*
 * Write the <drawing> elements.
 */
STATIC void
_worksheet_write_drawings(lxlsx_worksheet *self)
{
    if (!self->drawing)
        return;

    self->rel_count++;

    _worksheet_write_drawing(self, self->rel_count);
}

/*
 * Write the <formula1> element for numbers.
 */
STATIC void
_worksheet_write_formula1_num(lxlsx_worksheet *self, double number)
{
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, number);

    lxlsx_xml_data_element(self->file, "formula1", data, NULL);
}

/*
 * Write the <formula1> element for strings/formulas.
 */
STATIC void
_worksheet_write_formula1_str(lxlsx_worksheet *self, char *str)
{
    lxlsx_xml_data_element(self->file, "formula1", str, NULL);
}

/*
 * Write the <formula2> element for numbers.
 */
STATIC void
_worksheet_write_formula2_num(lxlsx_worksheet *self, double number)
{
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, number);

    lxlsx_xml_data_element(self->file, "formula2", data, NULL);
}

/*
 * Write the <formula2> element for strings/formulas.
 */
STATIC void
_worksheet_write_formula2_str(lxlsx_worksheet *self, char *str)
{
    lxlsx_xml_data_element(self->file, "formula2", str, NULL);
}

/*
 * Write the <dataValidation> element.
 */
STATIC void
_worksheet_write_data_validation(lxlsx_worksheet *self,
                                 lxlsx_data_val_obj *validation)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint8_t is_between = 0;

    LXLSX_INIT_ATTRIBUTES();

    switch (validation->validate) {
        case LXLSX_VALIDATION_TYPE_INTEGER:
        case LXLSX_VALIDATION_TYPE_INTEGER_FORMULA:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "whole");
            break;
        case LXLSX_VALIDATION_TYPE_DECIMAL:
        case LXLSX_VALIDATION_TYPE_DECIMAL_FORMULA:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "decimal");
            break;
        case LXLSX_VALIDATION_TYPE_LIST:
        case LXLSX_VALIDATION_TYPE_LIST_FORMULA:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "list");
            break;
        case LXLSX_VALIDATION_TYPE_DATE:
        case LXLSX_VALIDATION_TYPE_DATE_FORMULA:
        case LXLSX_VALIDATION_TYPE_DATE_NUMBER:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "date");
            break;
        case LXLSX_VALIDATION_TYPE_TIME:
        case LXLSX_VALIDATION_TYPE_TIME_FORMULA:
        case LXLSX_VALIDATION_TYPE_TIME_NUMBER:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "time");
            break;
        case LXLSX_VALIDATION_TYPE_LENGTH:
        case LXLSX_VALIDATION_TYPE_LENGTH_FORMULA:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "textLength");
            break;
        case LXLSX_VALIDATION_TYPE_CUSTOM_FORMULA:
            LXLSX_PUSH_ATTRIBUTES_STR("type", "custom");
            break;
    }

    switch (validation->criteria) {
        case LXLSX_VALIDATION_CRITERIA_EQUAL_TO:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "equal");
            break;
        case LXLSX_VALIDATION_CRITERIA_NOT_EQUAL_TO:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "notEqual");
            break;
        case LXLSX_VALIDATION_CRITERIA_LESS_THAN:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "lessThan");
            break;
        case LXLSX_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");
            break;
        case LXLSX_VALIDATION_CRITERIA_GREATER_THAN:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
            break;
        case LXLSX_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
            break;
        case LXLSX_VALIDATION_CRITERIA_BETWEEN:
            /* Between is the default for 2 formulas and isn't added. */
            is_between = 1;
            break;
        case LXLSX_VALIDATION_CRITERIA_NOT_BETWEEN:
            is_between = 1;
            LXLSX_PUSH_ATTRIBUTES_STR("operator", "notBetween");
            break;
    }

    if (validation->error_type == LXLSX_VALIDATION_ERROR_TYPE_WARNING)
        LXLSX_PUSH_ATTRIBUTES_STR("errorStyle", "warning");

    if (validation->error_type == LXLSX_VALIDATION_ERROR_TYPE_INFORMATION)
        LXLSX_PUSH_ATTRIBUTES_STR("errorStyle", "information");

    if (validation->ignore_blank)
        LXLSX_PUSH_ATTRIBUTES_INT("allowBlank", 1);

    if (validation->dropdown == LXLSX_VALIDATION_OFF)
        LXLSX_PUSH_ATTRIBUTES_INT("showDropDown", 1);

    if (validation->show_input)
        LXLSX_PUSH_ATTRIBUTES_INT("showInputMessage", 1);

    if (validation->show_error)
        LXLSX_PUSH_ATTRIBUTES_INT("showErrorMessage", 1);

    if (validation->error_title)
        LXLSX_PUSH_ATTRIBUTES_STR("errorTitle", validation->error_title);

    if (validation->error_message)
        LXLSX_PUSH_ATTRIBUTES_STR("error", validation->error_message);

    if (validation->input_title)
        LXLSX_PUSH_ATTRIBUTES_STR("promptTitle", validation->input_title);

    if (validation->input_message)
        LXLSX_PUSH_ATTRIBUTES_STR("prompt", validation->input_message);

    LXLSX_PUSH_ATTRIBUTES_STR("sqref", validation->sqref);

    if (validation->validate == LXLSX_VALIDATION_TYPE_ANY)
        lxlsx_xml_empty_tag(self->file, "dataValidation", &attributes);
    else
        lxlsx_xml_start_tag(self->file, "dataValidation", &attributes);

    /* Write the formula1 and formula2 elements. */
    switch (validation->validate) {
        case LXLSX_VALIDATION_TYPE_INTEGER:
        case LXLSX_VALIDATION_TYPE_DECIMAL:
        case LXLSX_VALIDATION_TYPE_LENGTH:
        case LXLSX_VALIDATION_TYPE_DATE:
        case LXLSX_VALIDATION_TYPE_TIME:
        case LXLSX_VALIDATION_TYPE_DATE_NUMBER:
        case LXLSX_VALIDATION_TYPE_TIME_NUMBER:
            _worksheet_write_formula1_num(self, validation->value_number);
            if (is_between)
                _worksheet_write_formula2_num(self,
                                              validation->maximum_number);
            break;
        case LXLSX_VALIDATION_TYPE_INTEGER_FORMULA:
        case LXLSX_VALIDATION_TYPE_DECIMAL_FORMULA:
        case LXLSX_VALIDATION_TYPE_LENGTH_FORMULA:
        case LXLSX_VALIDATION_TYPE_DATE_FORMULA:
        case LXLSX_VALIDATION_TYPE_TIME_FORMULA:
        case LXLSX_VALIDATION_TYPE_LIST:
        case LXLSX_VALIDATION_TYPE_LIST_FORMULA:
        case LXLSX_VALIDATION_TYPE_CUSTOM_FORMULA:
            _worksheet_write_formula1_str(self, validation->value_formula);
            if (is_between)
                _worksheet_write_formula2_str(self,
                                              validation->maximum_formula);
            break;
    }

    if (validation->validate != LXLSX_VALIDATION_TYPE_ANY)
        lxlsx_xml_end_tag(self->file, "dataValidation");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <dataValidations> element.
 */
STATIC void
_worksheet_write_data_validations(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_data_val_obj *data_validation;

    if (self->num_validations == 0)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->num_validations);

    lxlsx_xml_start_tag(self->file, "dataValidations", &attributes);

    STAILQ_FOREACH(data_validation, self->data_validations, list_pointers) {
        /* Write the dataValidation element. */
        _worksheet_write_data_validation(self, data_validation);
    }

    lxlsx_xml_end_tag(self->file, "dataValidations");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <formula> element for strings.
 */
STATIC void
_worksheet_write_formula_str(lxlsx_worksheet *self, char *data)
{
    lxlsx_xml_data_element(self->file, "formula", data, NULL);
}

/*
 * Write the <formula> element for numbers.
 */
STATIC void
_worksheet_write_formula_num(lxlsx_worksheet *self, double num)
{
    char data[LXLSX_ATTR_32];

    lxlsx_sprintf_dbl(data, num);
    lxlsx_xml_data_element(self->file, "formula", data, NULL);
}

/*
 * Write the <ext> element.
 */
STATIC void
_worksheet_write_ext(lxlsx_worksheet *self, char *uri)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_x_14[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:x14", xmlns_x_14);
    LXLSX_PUSH_ATTRIBUTES_STR("uri", uri);

    lxlsx_xml_start_tag(self->file, "ext", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <extLst> dataBar extension element.
 */
STATIC void
_worksheet_write_data_bar_ext(lxlsx_worksheet *self,
                              lxlsx_cond_format_obj *cond_format)
{
    /* Create a pseudo GUID for each unique Excel 2010 data bar. */
    cond_format->guid = calloc(1, LXLSX_GUID_LENGTH);
    lxlsx_snprintf(cond_format->guid, LXLSX_GUID_LENGTH,
                 "{DA7ABA51-AAAA-BBBB-%04X-%012X}",
                 self->index + 1, ++self->data_bar_2010_index);

    lxlsx_xml_start_tag(self->file, "extLst", NULL);

    _worksheet_write_ext(self, "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}");

    lxlsx_xml_data_element(self->file, "x14:id", cond_format->guid, NULL);

    lxlsx_xml_end_tag(self->file, "ext");
    lxlsx_xml_end_tag(self->file, "extLst");
}

/*
 * Write the <color> element.
 */
STATIC void
_worksheet_write_color(lxlsx_worksheet *self, lxlsx_color_t color)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb[LXLSX_ATTR_32];

    lxlsx_snprintf(rgb, LXLSX_ATTR_32, "FF%06X", color & LXLSX_COLOR_MASK);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb);

    lxlsx_xml_empty_tag(self->file, "color", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfvo> element for strings.
 */
STATIC void
_worksheet_write_cfvo_str(lxlsx_worksheet *self, uint8_t rule_type,
                          char *value, uint8_t data_bar_2010)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "min");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_NUMBER)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "num");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENT)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percent");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENTILE)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percentile");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_FORMULA)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "formula");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "max");

    if (!data_bar_2010 || (rule_type != LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
                           && rule_type != LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM))
        LXLSX_PUSH_ATTRIBUTES_STR("val", value);

    lxlsx_xml_empty_tag(self->file, "cfvo", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfvo> element for numbers.
 */
STATIC void
_worksheet_write_cfvo_num(lxlsx_worksheet *self, uint8_t rule_type,
                          double value, uint8_t data_bar_2010)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "min");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_NUMBER)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "num");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENT)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percent");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENTILE)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percentile");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_FORMULA)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "formula");
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "max");

    if (!data_bar_2010 || (rule_type != LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
                           && rule_type != LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM))
        LXLSX_PUSH_ATTRIBUTES_DBL("val", value);

    lxlsx_xml_empty_tag(self->file, "cfvo", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <iconSet> element.
 */
STATIC void
_worksheet_write_icon_set(lxlsx_worksheet *self,
                          lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char *icon_set[] = {
        "3Arrows",
        "3ArrowsGray",
        "3Flags",
        "3TrafficLights",
        "3TrafficLights2",
        "3Signs",
        "3Symbols",
        "3Symbols2",
        "4Arrows",
        "4ArrowsGray",
        "4RedToBlack",
        "4Rating",
        "4TrafficLights",
        "5Arrows",
        "5ArrowsGray",
        "5Rating",
        "5Quarters",
    };
    uint8_t percent = LXLSX_CONDITIONAL_RULE_TYPE_PERCENT;
    uint8_t style = cond_format->icon_style;

    LXLSX_INIT_ATTRIBUTES();

    if (style != LXLSX_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED)
        LXLSX_PUSH_ATTRIBUTES_STR("iconSet", icon_set[style]);

    if (cond_format->reverse_icons == LXLSX_TRUE)
        LXLSX_PUSH_ATTRIBUTES_STR("reverse", "1");

    if (cond_format->icons_only == LXLSX_TRUE)
        LXLSX_PUSH_ATTRIBUTES_STR("showValue", "0");

    lxlsx_xml_start_tag(self->file, "iconSet", &attributes);

    if (style < LXLSX_CONDITIONAL_ICONS_4_ARROWS_COLORED) {
        _worksheet_write_cfvo_num(self, percent, 0, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 33, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 67, LXLSX_FALSE);
    }

    if (style >= LXLSX_CONDITIONAL_ICONS_4_ARROWS_COLORED
        && style < LXLSX_CONDITIONAL_ICONS_5_ARROWS_COLORED) {
        _worksheet_write_cfvo_num(self, percent, 0, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 25, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 50, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 75, LXLSX_FALSE);
    }

    if (style >= LXLSX_CONDITIONAL_ICONS_5_ARROWS_COLORED
        && style <= LXLSX_CONDITIONAL_ICONS_5_QUARTERS) {
        _worksheet_write_cfvo_num(self, percent, 0, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 20, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 40, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 60, LXLSX_FALSE);
        _worksheet_write_cfvo_num(self, percent, 80, LXLSX_FALSE);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for data bar rules.
 */
STATIC void
_worksheet_write_cf_rule_icons(lxlsx_worksheet *self,
                               lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    _worksheet_write_icon_set(self, cond_format);

    lxlsx_xml_end_tag(self->file, "iconSet");
    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <dataBar> element.
 */
STATIC void
_worksheet_write_data_bar(lxlsx_worksheet *self,
                          lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (cond_format->bar_only)
        LXLSX_PUSH_ATTRIBUTES_STR("showValue", "0");

    lxlsx_xml_start_tag(self->file, "dataBar", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for data bar rules.
 */
STATIC void
_worksheet_write_cf_rule_data_bar(lxlsx_worksheet *self,
                                  lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    _worksheet_write_data_bar(self, cond_format);

    if (cond_format->min_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->min_rule_type,
                                  cond_format->min_value_string,
                                  cond_format->data_bar_2010);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->min_rule_type,
                                  cond_format->min_value,
                                  cond_format->data_bar_2010);
    }

    if (cond_format->max_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->max_rule_type,
                                  cond_format->max_value_string,
                                  cond_format->data_bar_2010);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->max_rule_type,
                                  cond_format->max_value,
                                  cond_format->data_bar_2010);
    }

    _worksheet_write_color(self, cond_format->bar_color);

    lxlsx_xml_end_tag(self->file, "dataBar");

    if (cond_format->data_bar_2010)
        _worksheet_write_data_bar_ext(self, cond_format);

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for 2 and 3 color scale rules.
 */
STATIC void
_worksheet_write_cf_rule_color_scale(lxlsx_worksheet *self,
                                     lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    lxlsx_xml_start_tag(self->file, "colorScale", NULL);

    if (cond_format->min_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->min_rule_type,
                                  cond_format->min_value_string, LXLSX_FALSE);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->min_rule_type,
                                  cond_format->min_value, LXLSX_FALSE);
    }

    if (cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE) {
        if (cond_format->mid_value_string) {
            _worksheet_write_cfvo_str(self, cond_format->mid_rule_type,
                                      cond_format->mid_value_string,
                                      LXLSX_FALSE);
        }
        else {
            _worksheet_write_cfvo_num(self, cond_format->mid_rule_type,
                                      cond_format->mid_value, LXLSX_FALSE);
        }
    }

    if (cond_format->max_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->max_rule_type,
                                  cond_format->max_value_string, LXLSX_FALSE);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->max_rule_type,
                                  cond_format->max_value, LXLSX_FALSE);
    }

    _worksheet_write_color(self, cond_format->min_color);

    if (cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE)
        _worksheet_write_color(self, cond_format->mid_color);

    _worksheet_write_color(self, cond_format->max_color);

    lxlsx_xml_end_tag(self->file, "colorScale");
    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for formula rules.
 */
STATIC void
_worksheet_write_cf_rule_formula(lxlsx_worksheet *self,
                                 lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    _worksheet_write_formula_str(self, cond_format->min_value_string);

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for top and bottom rules.
 */
STATIC void
_worksheet_write_cf_rule_top(lxlsx_worksheet *self,
                             lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    if (cond_format->criteria ==
        LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT)
        LXLSX_PUSH_ATTRIBUTES_INT("percent", 1);

    if (cond_format->type == LXLSX_CONDITIONAL_TYPE_BOTTOM)
        LXLSX_PUSH_ATTRIBUTES_INT("bottom", 1);

    /* Rank must be an int in the range 1-1000 . */
    if (cond_format->min_value < 1.0 || cond_format->min_value > 1000.0)
        LXLSX_PUSH_ATTRIBUTES_DBL("rank", 10);
    else
        LXLSX_PUSH_ATTRIBUTES_DBL("rank", (uint16_t) cond_format->min_value);

    lxlsx_xml_empty_tag(self->file, "cfRule", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for unique/duplicate rules.
 */
STATIC void
_worksheet_write_cf_rule_duplicate(lxlsx_worksheet *self,
                                   lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    /* Set the attributes common to all rule types. */
    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    lxlsx_xml_empty_tag(self->file, "cfRule", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for averages rules.
 */
STATIC void
_worksheet_write_cf_rule_average(lxlsx_worksheet *self,
                                 lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint8_t criteria = cond_format->criteria;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
        LXLSX_PUSH_ATTRIBUTES_INT("aboveAverage", 0);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL)
        LXLSX_PUSH_ATTRIBUTES_INT("equalAverage", 1);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW)
        LXLSX_PUSH_ATTRIBUTES_INT("stdDev", 1);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW)
        LXLSX_PUSH_ATTRIBUTES_INT("stdDev", 2);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE
        || criteria == LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
        LXLSX_PUSH_ATTRIBUTES_INT("stdDev", 3);

    lxlsx_xml_empty_tag(self->file, "cfRule", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for time_period rules.
 */
STATIC void
_worksheet_write_cf_rule_time_period(lxlsx_worksheet *self,
                                     lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char formula[LXLSX_MAX_ATTRIBUTE_LENGTH];
    uint8_t pos;
    uint8_t criteria = cond_format->criteria;
    char *first_cell = cond_format->first_cell;
    char *time_periods[] = {
        "yesterday",
        "today",
        "tomorrow",
        "last7Days",
        "lastWeek",
        "thisWeek",
        "nextWeek",
        "lastMonth",
        "thisMonth",
        "nextMonth",
    };

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    pos = criteria - LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY;
    LXLSX_PUSH_ATTRIBUTES_STR("timePeriod", time_periods[pos]);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()-1", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()+1", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),"
                     "TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,"
                     "ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),"
                     "ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY())-1,OR(YEAR(%s)=YEAR("
                     "TODAY()),AND(MONTH(%s)=1,YEAR(A1)=YEAR(TODAY())-1)))",
                     first_cell, first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {
        lxlsx_snprintf(formula, LXLSX_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY())+1,OR(YEAR(%s)=YEAR("
                     "TODAY()),AND(MONTH(%s)=12,YEAR(%s)=YEAR(TODAY())+1)))",
                     first_cell, first_cell, first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for blanks/no_blanks, errors/no_errors rules.
 */
STATIC void
_worksheet_write_cf_rule_blanks(lxlsx_worksheet *self,
                                lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char formula[LXLSX_ATTR_32];
    uint8_t type = cond_format->type;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    if (type == LXLSX_CONDITIONAL_TYPE_BLANKS) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32, "LEN(TRIM(%s))=0",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXLSX_CONDITIONAL_TYPE_NO_BLANKS) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32, "LEN(TRIM(%s))>0",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXLSX_CONDITIONAL_TYPE_ERRORS) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32, "ISERROR(%s)",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXLSX_CONDITIONAL_TYPE_NO_ERRORS) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32, "NOT(ISERROR(%s))",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for text rules.
 */
STATIC void
_worksheet_write_cf_rule_text(lxlsx_worksheet *self,
                              lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint8_t pos;
    char formula[LXLSX_ATTR_32 * 2];
    char *operators[] = {
        "containsText",
        "notContains",
        "beginsWith",
        "endsWith",
    };
    uint8_t criteria = cond_format->criteria;

    LXLSX_INIT_ATTRIBUTES();

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "containsText");
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "notContainsText");
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "beginsWith");
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH)
        LXLSX_PUSH_ATTRIBUTES_STR("type", "endsWith");

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    pos = criteria - LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING;
    LXLSX_PUSH_ATTRIBUTES_STR("operator", operators[pos]);

    LXLSX_PUSH_ATTRIBUTES_STR("text", cond_format->min_value_string);

    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32 * 2,
                     "NOT(ISERROR(SEARCH(\"%s\",%s)))",
                     cond_format->min_value_string, cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32 * 2,
                     "ISERROR(SEARCH(\"%s\",%s))",
                     cond_format->min_value_string, cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32 * 2,
                     "LEFT(%s,%d)=\"%s\"",
                     cond_format->first_cell,
                     (uint16_t) strlen(cond_format->min_value_string),
                     cond_format->min_value_string);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXLSX_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {
        lxlsx_snprintf(formula, LXLSX_ATTR_32 * 2,
                     "RIGHT(%s,%d)=\"%s\"",
                     cond_format->first_cell,
                     (uint16_t) strlen(cond_format->min_value_string),
                     cond_format->min_value_string);
        _worksheet_write_formula_str(self, formula);
    }

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for cell rules.
 */
STATIC void
_worksheet_write_cf_rule_cell(lxlsx_worksheet *self,
                              lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char *operators[] = {
        "none",
        "equal",
        "notEqual",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqual",
        "lessThanOrEqual",
        "between",
        "notBetween",
    };

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXLSX_PROPERTY_UNSET)
        LXLSX_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXLSX_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXLSX_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    LXLSX_PUSH_ATTRIBUTES_STR("operator", operators[cond_format->criteria]);

    lxlsx_xml_start_tag(self->file, "cfRule", &attributes);

    if (cond_format->min_value_string)
        _worksheet_write_formula_str(self, cond_format->min_value_string);
    else
        _worksheet_write_formula_num(self, cond_format->min_value);

    if (cond_format->has_max) {
        if (cond_format->max_value_string)
            _worksheet_write_formula_str(self, cond_format->max_value_string);
        else
            _worksheet_write_formula_num(self, cond_format->max_value);
    }

    lxlsx_xml_end_tag(self->file, "cfRule");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element.
 */
STATIC void
_worksheet_write_cf_rule(lxlsx_worksheet *self,
                         lxlsx_cond_format_obj *cond_format)
{
    if (cond_format->type == LXLSX_CONDITIONAL_TYPE_CELL) {

        _worksheet_write_cf_rule_cell(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TEXT) {

        _worksheet_write_cf_rule_text(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TIME_PERIOD) {

        _worksheet_write_cf_rule_time_period(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_DUPLICATE
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_UNIQUE) {

        _worksheet_write_cf_rule_duplicate(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_AVERAGE) {

        _worksheet_write_cf_rule_average(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_BOTTOM) {

        _worksheet_write_cf_rule_top(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_BLANKS
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_NO_BLANKS
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_ERRORS
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_NO_ERRORS) {

        _worksheet_write_cf_rule_blanks(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_FORMULA) {

        _worksheet_write_cf_rule_formula(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_2_COLOR_SCALE
             || cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE) {

        _worksheet_write_cf_rule_color_scale(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_DATA_BAR) {

        _worksheet_write_cf_rule_data_bar(self, cond_format);
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_ICON_SETS) {

        _worksheet_write_cf_rule_icons(self, cond_format);
    }

}

/*
 * Write the <conditionalFormatting> element.
 */
STATIC void
_worksheet_write_conditional_formatting(lxlsx_worksheet *self,
                                        lxlsx_cond_format_hash_element *element)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_cond_format_obj *cond_format;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("sqref", element->sqref);

    lxlsx_xml_start_tag(self->file, "conditionalFormatting", &attributes);

    STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
        /* Write the cfRule element. */
        _worksheet_write_cf_rule(self, cond_format);
    }

    lxlsx_xml_end_tag(self->file, "conditionalFormatting");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the conditional formatting> elements.
 */
STATIC void
_worksheet_write_conditional_formats(lxlsx_worksheet *self)
{
    lxlsx_cond_format_hash_element *element;
    lxlsx_cond_format_hash_element *next_element;

    for (element = RB_MIN(lxlsx_cond_format_hash, self->conditional_formats);
         element; element = next_element) {

        _worksheet_write_conditional_formatting(self, element);

        next_element =
            RB_NEXT(lxlsx_cond_format_hash, self->conditional_formats, element);
    }
}

/*
 * Write the <x14:xxxColor> elements for data bar conditional formats.
 */
STATIC void
_worksheet_write_x14_color(lxlsx_worksheet *self, char *type, lxlsx_color_t color)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rgb[LXLSX_ATTR_32];

    lxlsx_snprintf(rgb, LXLSX_ATTR_32, "FF%06X", color & LXLSX_COLOR_MASK);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("rgb", rgb);
    lxlsx_xml_empty_tag(self->file, type, &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:cfvo> element.
 */
STATIC void
_worksheet_write_x14_cfvo(lxlsx_worksheet *self, uint8_t rule_type,
                          double number, char *string)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char data[LXLSX_ATTR_32];
    uint8_t has_value = LXLSX_FALSE;

    LXLSX_INIT_ATTRIBUTES();

    if (!string)
        lxlsx_sprintf_dbl(data, number);

    if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MIN) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "autoMin");
        has_value = LXLSX_FALSE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "min");
        has_value = LXLSX_FALSE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_NUMBER) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "num");
        has_value = LXLSX_TRUE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENT) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percent");
        has_value = LXLSX_TRUE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_PERCENTILE) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "percentile");
        has_value = LXLSX_TRUE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_FORMULA) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "formula");
        has_value = LXLSX_TRUE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "max");
        has_value = LXLSX_FALSE;
    }
    else if (rule_type == LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MAX) {
        LXLSX_PUSH_ATTRIBUTES_STR("type", "autoMax");
        has_value = LXLSX_FALSE;
    }

    if (has_value) {
        lxlsx_xml_start_tag(self->file, "x14:cfvo", &attributes);

        if (string)
            lxlsx_xml_data_element(self->file, "xm:f", string, NULL);
        else
            lxlsx_xml_data_element(self->file, "xm:f", data, NULL);

        lxlsx_xml_end_tag(self->file, "x14:cfvo");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "x14:cfvo", &attributes);
    }
    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:dataBar> element.
 */
STATIC void
_worksheet_write_x14_data_bar(lxlsx_worksheet *self,
                              lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char min_length[] = "0";
    char max_length[] = "100";
    char border[] = "1";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("minLength", min_length);
    LXLSX_PUSH_ATTRIBUTES_STR("maxLength", max_length);

    if (!cond_format->bar_no_border)
        LXLSX_PUSH_ATTRIBUTES_STR("border", border);

    if (cond_format->bar_solid)
        LXLSX_PUSH_ATTRIBUTES_STR("gradient", "0");

    if (cond_format->bar_direction ==
        LXLSX_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT)
        LXLSX_PUSH_ATTRIBUTES_STR("direction", "rightToLeft");

    if (cond_format->bar_direction ==
        LXLSX_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT)
        LXLSX_PUSH_ATTRIBUTES_STR("direction", "leftToRight");

    if (cond_format->bar_negative_color_same)
        LXLSX_PUSH_ATTRIBUTES_STR("negativeBarColorSameAsPositive", "1");

    if (!cond_format->bar_no_border
        && !cond_format->bar_negative_border_color_same)
        LXLSX_PUSH_ATTRIBUTES_STR("negativeBarBorderColorSameAsPositive", "0");

    if (cond_format->bar_axis_position == LXLSX_CONDITIONAL_BAR_AXIS_MIDPOINT)
        LXLSX_PUSH_ATTRIBUTES_STR("axisPosition", "middle");

    if (cond_format->bar_axis_position == LXLSX_CONDITIONAL_BAR_AXIS_NONE)
        LXLSX_PUSH_ATTRIBUTES_STR("axisPosition", "none");

    lxlsx_xml_start_tag(self->file, "x14:dataBar", &attributes);

    if (cond_format->auto_min)
        cond_format->min_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MIN;

    _worksheet_write_x14_cfvo(self, cond_format->min_rule_type,
                              cond_format->min_value,
                              cond_format->min_value_string);

    if (cond_format->auto_max)
        cond_format->max_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_AUTO_MAX;

    _worksheet_write_x14_cfvo(self, cond_format->max_rule_type,
                              cond_format->max_value,
                              cond_format->max_value_string);

    if (!cond_format->bar_no_border)
        _worksheet_write_x14_color(self, "x14:borderColor",
                                   cond_format->bar_border_color);

    if (!cond_format->bar_negative_color_same)
        _worksheet_write_x14_color(self, "x14:negativeFillColor",
                                   cond_format->bar_negative_color);

    if (!cond_format->bar_no_border
        && !cond_format->bar_negative_border_color_same)
        _worksheet_write_x14_color(self, "x14:negativeBorderColor",
                                   cond_format->bar_negative_border_color);

    if (cond_format->bar_axis_position != LXLSX_CONDITIONAL_BAR_AXIS_NONE)
        _worksheet_write_x14_color(self, "x14:axisColor",
                                   cond_format->bar_axis_color);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:cfRule> element.
 */
STATIC void
_worksheet_write_x14_cf_rule(lxlsx_worksheet *self,
                             lxlsx_cond_format_obj *cond_format)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("type", "dataBar");
    LXLSX_PUSH_ATTRIBUTES_STR("id", cond_format->guid);

    lxlsx_xml_start_tag(self->file, "x14:cfRule", &attributes);

    /* Write the x14:dataBar element. */
    _worksheet_write_x14_data_bar(self, cond_format);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xm:sqref> element.
 */
STATIC void
_worksheet_write_xm_sqref(lxlsx_worksheet *self,
                          lxlsx_cond_format_obj *cond_format)
{
    lxlsx_xml_data_element(self->file, "xm:sqref", cond_format->sqref, NULL);
}

/*
 * Write the <conditionalFormatting> element.
 */
STATIC void
_worksheet_write_conditional_formatting_2010(lxlsx_worksheet *self, lxlsx_cond_format_hash_element
                                             *element)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_cond_format_obj *cond_format;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:xm",
                            "http://schemas.microsoft.com/office/excel/2006/main");

    STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
        if (!cond_format->data_bar_2010)
            continue;

        lxlsx_xml_start_tag(self->file, "x14:conditionalFormatting",
                          &attributes);

        _worksheet_write_x14_cf_rule(self, cond_format);

        lxlsx_xml_end_tag(self->file, "x14:dataBar");
        lxlsx_xml_end_tag(self->file, "x14:cfRule");
        _worksheet_write_xm_sqref(self, cond_format);
        lxlsx_xml_end_tag(self->file, "x14:conditionalFormatting");
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <extLst> element for Excel 2010 conditional formatting data bars.
 */
STATIC void
_worksheet_write_ext_list_data_bars(lxlsx_worksheet *self)
{
    lxlsx_cond_format_hash_element *element;
    lxlsx_cond_format_hash_element *next_element;

    _worksheet_write_ext(self, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
    lxlsx_xml_start_tag(self->file, "x14:conditionalFormattings", NULL);

    for (element = RB_MIN(lxlsx_cond_format_hash, self->conditional_formats);
         element; element = next_element) {

        _worksheet_write_conditional_formatting_2010(self, element);

        next_element =
            RB_NEXT(lxlsx_cond_format_hash, self->conditional_formats, element);
    }

    lxlsx_xml_end_tag(self->file, "x14:conditionalFormattings");
    lxlsx_xml_end_tag(self->file, "ext");
}

/*
 * Write the <extLst> element.
 */
STATIC void
_worksheet_write_ext_list(lxlsx_worksheet *self)
{
    if (self->data_bar_2010_index == 0)
        return;

    lxlsx_xml_start_tag(self->file, "extLst", NULL);

    _worksheet_write_ext_list_data_bars(self);

    lxlsx_xml_end_tag(self->file, "extLst");
}

/*
 * Write the <ignoredError> element.
 */
STATIC void
_worksheet_write_ignored_error(lxlsx_worksheet *self, char *ignore_error,
                               char *range)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("sqref", range);
    LXLSX_PUSH_ATTRIBUTES_STR(ignore_error, "1");

    lxlsx_xml_empty_tag(self->file, "ignoredError", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

lxlsx_error
_validate_conditional_icons(lxlsx_conditional_format *user)
{
    if (user->icon_style > LXLSX_CONDITIONAL_ICONS_5_QUARTERS) {

        LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "For type = LXLSX_CONDITIONAL_TYPE_ICON_SETS, "
                         "invalid icon_style (%d).", user->icon_style);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXLSX_NO_ERROR;
    }
}

lxlsx_error
_validate_conditional_data_bar(lxlsx_worksheet *self,
                               lxlsx_cond_format_obj *cond_format,
                               lxlsx_conditional_format *user_options)
{
    uint8_t min_rule_type = user_options->min_rule_type;
    uint8_t max_rule_type = user_options->max_rule_type;

    if (user_options->data_bar_2010
        || user_options->bar_solid
        || user_options->bar_no_border
        || user_options->bar_direction
        || user_options->bar_axis_position
        || user_options->bar_negative_color_same
        || user_options->bar_negative_border_color_same
        || user_options->bar_negative_color
        || user_options->bar_border_color
        || user_options->bar_negative_border_color
        || user_options->bar_axis_color) {

        cond_format->data_bar_2010 = LXLSX_TRUE;
        self->excel_version = 2010;
    }

    if (min_rule_type > LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
        && min_rule_type < LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->min_rule_type = min_rule_type;
        cond_format->min_value = user_options->min_value;
        cond_format->min_value_string =
            lxlsx_strdup_formula(user_options->min_value_string);
    }
    else {
        cond_format->min_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM;
        cond_format->min_value = 0;
    }

    if (max_rule_type > LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
        && max_rule_type < LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->max_rule_type = max_rule_type;
        cond_format->max_value = user_options->max_value;
        cond_format->max_value_string =
            lxlsx_strdup_formula(user_options->max_value_string);
    }
    else {
        cond_format->max_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM;
        cond_format->max_value = 0;
    }

    if (cond_format->data_bar_2010) {
        if (min_rule_type == LXLSX_CONDITIONAL_RULE_TYPE_NONE)
            cond_format->auto_min = LXLSX_TRUE;
        if (max_rule_type == LXLSX_CONDITIONAL_RULE_TYPE_NONE)
            cond_format->auto_max = LXLSX_TRUE;
    }

    cond_format->bar_only = user_options->bar_only;
    cond_format->bar_solid = user_options->bar_solid;
    cond_format->bar_no_border = user_options->bar_no_border;
    cond_format->bar_direction = user_options->bar_direction;
    cond_format->bar_axis_position = user_options->bar_axis_position;
    cond_format->bar_negative_color_same =
        user_options->bar_negative_color_same;
    cond_format->bar_negative_border_color_same =
        user_options->bar_negative_border_color_same;

    if (user_options->bar_color != LXLSX_COLOR_UNSET)
        cond_format->bar_color = user_options->bar_color;
    else
        cond_format->bar_color = 0x638EC6;

    if (user_options->bar_negative_color != LXLSX_COLOR_UNSET)
        cond_format->bar_negative_color = user_options->bar_negative_color;
    else
        cond_format->bar_negative_color = 0xFF0000;

    if (user_options->bar_border_color != LXLSX_COLOR_UNSET)
        cond_format->bar_border_color = user_options->bar_border_color;
    else
        cond_format->bar_border_color = cond_format->bar_color;

    if (user_options->bar_negative_border_color != LXLSX_COLOR_UNSET)
        cond_format->bar_negative_border_color =
            user_options->bar_negative_border_color;
    else
        cond_format->bar_negative_border_color = 0xFF0000;

    if (user_options->bar_axis_color != LXLSX_COLOR_UNSET)
        cond_format->bar_axis_color = user_options->bar_axis_color;
    else
        cond_format->bar_axis_color = 0x000000;

    return LXLSX_NO_ERROR;
}

lxlsx_error
_validate_conditional_scale(lxlsx_cond_format_obj *cond_format,
                            lxlsx_conditional_format *user_options)
{
    uint8_t min_rule_type = user_options->min_rule_type;
    uint8_t mid_rule_type = user_options->mid_rule_type;
    uint8_t max_rule_type = user_options->max_rule_type;

    if (min_rule_type > LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
        && min_rule_type < LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->min_rule_type = min_rule_type;
        cond_format->min_value = user_options->min_value;
        cond_format->min_value_string =
            lxlsx_strdup_formula(user_options->min_value_string);
    }
    else {
        cond_format->min_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM;
        cond_format->min_value = 0;
    }

    if (max_rule_type > LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
        && max_rule_type < LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->max_rule_type = max_rule_type;
        cond_format->max_value = user_options->max_value;
        cond_format->max_value_string =
            lxlsx_strdup_formula(user_options->max_value_string);
    }
    else {
        cond_format->max_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM;
        cond_format->max_value = 0;
    }

    if (cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE) {
        if (mid_rule_type > LXLSX_CONDITIONAL_RULE_TYPE_MINIMUM
            && mid_rule_type < LXLSX_CONDITIONAL_RULE_TYPE_MAXIMUM) {
            cond_format->mid_rule_type = mid_rule_type;
            cond_format->mid_value = user_options->mid_value;
            cond_format->mid_value_string =
                lxlsx_strdup_formula(user_options->mid_value_string);
        }
        else {
            cond_format->mid_rule_type = LXLSX_CONDITIONAL_RULE_TYPE_PERCENTILE;
            cond_format->mid_value = 50;
        }
    }

    if (user_options->min_color != LXLSX_COLOR_UNSET)
        cond_format->min_color = user_options->min_color;
    else
        cond_format->min_color = 0xFF7128;

    if (user_options->max_color != LXLSX_COLOR_UNSET)
        cond_format->max_color = user_options->max_color;
    else
        cond_format->max_color = 0xFFEF9C;

    if (cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE) {
        if (user_options->min_color == LXLSX_COLOR_UNSET)
            cond_format->min_color = 0xF8696B;

        if (user_options->mid_color != LXLSX_COLOR_UNSET)
            cond_format->mid_color = user_options->mid_color;
        else
            cond_format->mid_color = 0xFFEB84;

        if (user_options->max_color == LXLSX_COLOR_UNSET)
            cond_format->max_color = 0x63BE7B;
    }

    return LXLSX_NO_ERROR;
}

lxlsx_error
_validate_conditional_top(lxlsx_cond_format_obj *cond_format,
                          lxlsx_conditional_format *user_options)
{
    /* Restrict the range of rank values to Excel's allowed range. */
    if (user_options->criteria ==
        LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT) {
        if (user_options->value < 0.0 || user_options->value > 100.0) {

            LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                             "For type = LXLSX_CONDITIONAL_TYPE_TOP/BOTTOM, "
                             "top/bottom percent (%g%%) must by in range 0-100",
                             user_options->value);

            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }
    }
    else {
        if (user_options->value < 1.0 || user_options->value > 1000.0) {

            LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                             "For type = LXLSX_CONDITIONAL_TYPE_TOP/BOTTOM, "
                             "top/bottom items (%g) must by in range 1-1000",
                             user_options->value);

            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }
    }

    cond_format->min_value = (uint16_t) user_options->value;

    return LXLSX_NO_ERROR;
}

lxlsx_error
_validate_conditional_average(lxlsx_conditional_format *user)
{
    if (user->criteria < LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE ||
        user->criteria > LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW) {

        LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "For type = LXLSX_CONDITIONAL_TYPE_AVERAGE, "
                         "invalid criteria value (%d).", user->criteria);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXLSX_NO_ERROR;
    }
}

lxlsx_error
_validate_conditional_time_period(lxlsx_conditional_format *user)
{
    if (user->criteria < LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY ||
        user->criteria > LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {

        LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "For type = LXLSX_CONDITIONAL_TYPE_TIME_PERIOD, "
                         "invalid criteria value (%d).", user->criteria);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXLSX_NO_ERROR;
    }
}

lxlsx_error
_validate_conditional_text(lxlsx_cond_format_obj *cond_format,
                           lxlsx_conditional_format *user_options)
{
    if (!user_options->value_string) {

        LXLSX_WARN_FORMAT("lxlsx_worksheet_conditional_format_cell()/_range(): "
                        "For type = LXLSX_CONDITIONAL_TYPE_TEXT, "
                        "value_string can not be NULL. "
                        "Text must be specified.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (strlen(user_options->value_string) >= LXLSX_MAX_ATTRIBUTE_LENGTH) {

        LXLSX_WARN_FORMAT2("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "For type = LXLSX_CONDITIONAL_TYPE_TEXT, "
                         "value_string length (%d) must be less than %d.",
                         (uint16_t) strlen(user_options->value_string),
                         LXLSX_MAX_ATTRIBUTE_LENGTH);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (user_options->criteria < LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING ||
        user_options->criteria > LXLSX_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {

        LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "For type = LXLSX_CONDITIONAL_TYPE_TEXT, "
                         "invalid criteria value (%d).",
                         user_options->criteria);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    cond_format->min_value_string =
        lxlsx_strdup_formula(user_options->value_string);

    return LXLSX_NO_ERROR;
}

lxlsx_error
_validate_conditional_formula(lxlsx_cond_format_obj *cond_format,
                              lxlsx_conditional_format *user_options)
{
    if (!user_options->value_string) {

        LXLSX_WARN_FORMAT("lxlsx_worksheet_conditional_format_cell()/_range(): "
                        "For type = LXLSX_CONDITIONAL_TYPE_FORMULA, "
                        "value_string can not be NULL. "
                        "Formula must be specified.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    cond_format->min_value_string =
        lxlsx_strdup_formula(user_options->value_string);

    return LXLSX_NO_ERROR;
}

lxlsx_error
_validate_conditional_cell(lxlsx_cond_format_obj *cond_format,
                           lxlsx_conditional_format *user_options)
{
    cond_format->min_value = user_options->value;
    cond_format->min_value_string =
        lxlsx_strdup_formula(user_options->value_string);

    if (cond_format->criteria == LXLSX_CONDITIONAL_CRITERIA_BETWEEN
        || cond_format->criteria == LXLSX_CONDITIONAL_CRITERIA_NOT_BETWEEN) {
        cond_format->has_max = LXLSX_TRUE;
        cond_format->min_value = user_options->min_value;
        cond_format->max_value = user_options->max_value;
        cond_format->min_value_string =
            lxlsx_strdup_formula(user_options->min_value_string);
        cond_format->max_value_string =
            lxlsx_strdup_formula(user_options->max_value_string);
    }

    return LXLSX_NO_ERROR;
}

/* Check that the correct criteria and used with the correct conditional format. */
lxlsx_error
_validate_conditional_criteria(lxlsx_cond_format_obj *cond_format)
{
    uint8_t criteria_mismatch = LXLSX_FALSE;

    if (cond_format->type == LXLSX_CONDITIONAL_TYPE_CELL) {
        switch (cond_format->criteria) {
            case LXLSX_CONDITIONAL_CRITERIA_EQUAL_TO:
            case LXLSX_CONDITIONAL_CRITERIA_NOT_EQUAL_TO:
            case LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN:
            case LXLSX_CONDITIONAL_CRITERIA_LESS_THAN:
            case LXLSX_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
            case LXLSX_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO:
            case LXLSX_CONDITIONAL_CRITERIA_BETWEEN:
            case LXLSX_CONDITIONAL_CRITERIA_NOT_BETWEEN:
                criteria_mismatch = LXLSX_FALSE;
                break;
            default:
                criteria_mismatch = LXLSX_TRUE;
        }
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TIME_PERIOD) {
        switch (cond_format->criteria) {
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH:
            case LXLSX_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH:
                criteria_mismatch = LXLSX_FALSE;
                break;
            default:
                criteria_mismatch = LXLSX_TRUE;
        }
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TEXT) {
        switch (cond_format->criteria) {
            case LXLSX_CONDITIONAL_CRITERIA_TEXT_CONTAINING:
            case LXLSX_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING:
            case LXLSX_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH:
            case LXLSX_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH:
                criteria_mismatch = LXLSX_FALSE;
                break;
            default:
                criteria_mismatch = LXLSX_TRUE;
        }
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_AVERAGE) {
        switch (cond_format->criteria) {
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE:
            case LXLSX_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW:
                criteria_mismatch = LXLSX_FALSE;
                break;
            default:
                criteria_mismatch = LXLSX_TRUE;
        }
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_BOTTOM) {
        switch (cond_format->criteria) {
            case LXLSX_CONDITIONAL_CRITERIA_NONE:
            case LXLSX_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT:
                criteria_mismatch = LXLSX_FALSE;
                break;
            default:
                criteria_mismatch = LXLSX_TRUE;
        }
    }
    else {
        /* Any other conditional type should have a zero criteria. */
        cond_format->criteria = LXLSX_CONDITIONAL_CRITERIA_NONE;
    }

    if (criteria_mismatch) {
        LXLSX_WARN_FORMAT2("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "LXLSX_CONDITIONAL_CRITERIA_* = %d is not valid for "
                         "LXLSX_CONDITIONAL_TYPE_* = %d", cond_format->criteria,
                         cond_format->type);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXLSX_NO_ERROR;
    }
}

/*
 * Write the <ignoredErrors> element.
 */
STATIC void
_worksheet_write_ignored_errors(lxlsx_worksheet *self)
{
    if (!self->has_ignore_errors)
        return;

    lxlsx_xml_start_tag(self->file, "ignoredErrors", NULL);

    if (self->ignore_number_stored_as_text) {
        _worksheet_write_ignored_error(self, "numberStoredAsText",
                                       self->ignore_number_stored_as_text);
    }

    if (self->ignore_eval_error) {
        _worksheet_write_ignored_error(self, "evalError",
                                       self->ignore_eval_error);
    }

    if (self->ignore_formula_differs) {
        _worksheet_write_ignored_error(self, "formula",
                                       self->ignore_formula_differs);
    }

    if (self->ignore_formula_range) {
        _worksheet_write_ignored_error(self, "formulaRange",
                                       self->ignore_formula_range);
    }

    if (self->ignore_formula_unlocked) {
        _worksheet_write_ignored_error(self, "unlockedFormula",
                                       self->ignore_formula_unlocked);
    }

    if (self->ignore_empty_cell_reference) {
        _worksheet_write_ignored_error(self, "emptyCellReference",
                                       self->ignore_empty_cell_reference);
    }

    if (self->ignore_list_data_validation) {
        _worksheet_write_ignored_error(self, "listDataValidation",
                                       self->ignore_list_data_validation);
    }

    if (self->ignore_calculated_column) {
        _worksheet_write_ignored_error(self, "calculatedColumn",
                                       self->ignore_calculated_column);
    }

    if (self->ignore_two_digit_text_year) {
        _worksheet_write_ignored_error(self, "twoDigitTextYear",
                                       self->ignore_two_digit_text_year);
    }

    lxlsx_xml_end_tag(self->file, "ignoredErrors");
}

/*
 * Write the <tablePart> element.
 */
STATIC void
_worksheet_write_table_part(lxlsx_worksheet *self, uint16_t id)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", id);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "tablePart", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <tableParts> element.
 */
STATIC void
_worksheet_write_table_parts(lxlsx_worksheet *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    lxlsx_table_obj *lxlsx_table_obj;

    if (!self->lxlsx_table_count)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", self->lxlsx_table_count);

    lxlsx_xml_start_tag(self->file, "tableParts", &attributes);

    STAILQ_FOREACH(lxlsx_table_obj, self->lxlsx_table_objs, list_pointers) {
        self->rel_count++;

        /* Write the tablePart element. */
        _worksheet_write_table_part(self, self->rel_count);
    }

    lxlsx_xml_end_tag(self->file, "tableParts");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * External functions to call intern XML methods shared with chartsheet.
 */
void
lxlsx_worksheet_write_sheet_views(lxlsx_worksheet *self)
{
    _worksheet_write_sheet_views(self);
}

void
lxlsx_worksheet_write_page_margins(lxlsx_worksheet *self)
{
    _worksheet_write_page_margins(self);
}

void
lxlsx_worksheet_write_drawings(lxlsx_worksheet *self)
{
    _worksheet_write_drawings(self);
}

void
lxlsx_worksheet_write_sheet_protection(lxlsx_worksheet *self,
                                     lxlsx_protection_obj *protect)
{
    _worksheet_write_sheet_protection(self, protect);
}

void
lxlsx_worksheet_write_sheet_pr(lxlsx_worksheet *self)
{
    _worksheet_write_sheet_pr(self);
}

void
lxlsx_worksheet_write_page_setup(lxlsx_worksheet *self)
{
    _worksheet_write_page_setup(self);
}

void
lxlsx_worksheet_write_header_footer(lxlsx_worksheet *self)
{
    _worksheet_write_header_footer(self);
}

/*
 * Assemble and write the XML file.
 */
void
lxlsx_worksheet_assemble_xml_file(lxlsx_worksheet *self)
{
    /* Write the XML declaration. */
    _worksheet_xml_declaration(self);

    /* Write the worksheet element. */
    _worksheet_write_worksheet(self);

    /* Write the worksheet properties. */
    _worksheet_write_sheet_pr(self);

    /* Write the worksheet dimensions. */
    _worksheet_write_dimension(self);

    /* Write the sheet view properties. */
    _worksheet_write_sheet_views(self);

    /* Write the sheet format properties. */
    _worksheet_write_sheet_format_pr(self);

    /* Write the sheet column info. */
    _worksheet_write_cols(self);

    /* Write the sheetData element. */
    if (!self->optimize)
        _worksheet_write_sheet_data(self);
    else
        _worksheet_write_optimized_sheet_data(self);

    /* Write the sheetProtection element. */
    _worksheet_write_sheet_protection(self, &self->protection);

    /* Write the autoFilter element. */
    _worksheet_write_auto_filter(self);

    /* Write the mergeCells element. */
    _worksheet_write_merge_cells(self);

    /* Write the conditionalFormatting elements. */
    _worksheet_write_conditional_formats(self);

    /* Write the dataValidations element. */
    _worksheet_write_data_validations(self);

    /* Write the hyperlink element. */
    _worksheet_write_hyperlinks(self);

    /* Write the printOptions element. */
    _worksheet_write_print_options(self);

    /* Write the worksheet page_margins. */
    _worksheet_write_page_margins(self);

    /* Write the worksheet page setup. */
    _worksheet_write_page_setup(self);

    /* Write the headerFooter element. */
    _worksheet_write_header_footer(self);

    /* Write the rowBreaks element. */
    _worksheet_write_row_breaks(self);

    /* Write the colBreaks element. */
    _worksheet_write_col_breaks(self);

    /* Write the ignoredErrors element. */
    _worksheet_write_ignored_errors(self);

    /* Write the drawing element. */
    _worksheet_write_drawings(self);

    /* Write the legacyDrawing element. */
    _worksheet_write_legacy_drawing(self);

    /* Write the legacyDrawingHF element. */
    _worksheet_write_legacy_drawing_hf(self);

    /* Write the picture element. */
    _worksheet_write_picture(self);

    /* Write the tableParts element. */
    _worksheet_write_table_parts(self);

    /* Write the extLst element. */
    _worksheet_write_ext_list(self);

    /* Close the worksheet tag. */
    lxlsx_xml_end_tag(self->file, "worksheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

static uint8_t
_worksheet_is_edit(lxlsx_worksheet *self)
{
    return self && self->is_edit && self->edit_session;
}

static lxlsx_error
_worksheet_edit_reject_format(lxlsx_format *format)
{
    return format ? LXLSX_ERROR_FEATURE_NOT_SUPPORTED : LXLSX_NO_ERROR;
}

/*
 * Write a number to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_number(lxlsx_worksheet *self,
                       lxlsx_row_t row_num,
                       lxlsx_col_t col_num, double value, lxlsx_format *format)
{
    lxlsx_cell *cell;
    lxlsx_error err;

    if (_worksheet_is_edit(self)) {
        err = _worksheet_edit_reject_format(format);
        if (err)
            return err;

        err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
        if (err)
            return err;

        return lxlsx_edit_set_number(self->edit_session, self->edit_sheet_name,
                                     row_num, col_num, value);
    }

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    cell = _new_number_cell(row_num, col_num, value, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a string to an Excel file.
 */
lxlsx_error
lxlsx_worksheet_write_string(lxlsx_worksheet *self,
                       lxlsx_row_t row_num,
                       lxlsx_col_t col_num, const char *string,
                       lxlsx_format *format)
{
    lxlsx_cell *cell;
    int32_t string_id;
    char *string_copy;
    struct lxlsx_sst_element *lxlsx_sst_element;
    lxlsx_error err;

    if (_worksheet_is_edit(self)) {
        if (!string || !*string)
            return format ? LXLSX_ERROR_FEATURE_NOT_SUPPORTED : LXLSX_NO_ERROR;

        err = _worksheet_edit_reject_format(format);
        if (err)
            return err;

        err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
        if (err)
            return err;

        if (lxlsx_utf8_strlen(string) > LXLSX_STR_MAX)
            return LXLSX_ERROR_MAX_STRING_LENGTH_EXCEEDED;

        return lxlsx_edit_set_string(self->edit_session, self->edit_sheet_name,
                                     row_num, col_num, string);
    }

    if (!string || !*string) {
        /* Treat a NULL or empty string with formatting as a blank cell. */
        /* Null strings without formats should be ignored.      */
        if (format)
            return lxlsx_worksheet_write_blank(self, row_num, col_num, format);
        else
            return LXLSX_NO_ERROR;
    }

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    if (lxlsx_utf8_strlen(string) > LXLSX_STR_MAX)
        return LXLSX_ERROR_MAX_STRING_LENGTH_EXCEEDED;

    if (!self->optimize) {
        /* Get the SST element and string id. */
        lxlsx_sst_element = lxlsx_get_sst_index(self->sst, string, LXLSX_FALSE);

        if (!lxlsx_sst_element)
            return LXLSX_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = lxlsx_sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id,
                                lxlsx_sst_element->string, format);
    }
    else {
        /* Look for and escape control chars in the string. */
        if (lxlsx_has_control_characters(string)) {
            string_copy = lxlsx_escape_control_characters(string);
        }
        else {
            string_copy = lxlsx_strdup(string);
        }
        cell = _new_inline_string_cell(row_num, col_num, string_copy, format);
    }

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a formula with a numerical result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_formula_num(lxlsx_worksheet *self,
                            lxlsx_row_t row_num,
                            lxlsx_col_t col_num,
                            const char *formula,
                            lxlsx_format *format, double result)
{
    lxlsx_cell *cell;
    char *formula_copy;
    lxlsx_error err;
    char result_buf[64];

    if (_worksheet_is_edit(self)) {
        err = _worksheet_edit_reject_format(format);
        if (err)
            return err;

        if (!formula)
            return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

        if (lxlsx_str_is_empty(formula))
            return LXLSX_ERROR_PARAMETER_IS_EMPTY;

        err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
        if (err)
            return err;

        snprintf(result_buf, sizeof(result_buf), "%.15g", result);
        return lxlsx_edit_set_formula(self->edit_session, self->edit_sheet_name,
                                      row_num, col_num, formula, result_buf);
    }

    if (!formula)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_str_is_empty(formula))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Strip leading "=" from formula. */
    if (formula[0] == '=')
        formula_copy = lxlsx_strdup(formula + 1);
    else
        formula_copy = lxlsx_strdup(formula);

    cell = _new_formula_cell(row_num, col_num, formula_copy, format);
    cell->formula_result = result;

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a formula with a string result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_formula_str(lxlsx_worksheet *self,
                            lxlsx_row_t row_num,
                            lxlsx_col_t col_num,
                            const char *formula,
                            lxlsx_format *format, const char *result)
{
    lxlsx_cell *cell;
    char *formula_copy;
    lxlsx_error err;

    if (_worksheet_is_edit(self)) {
        err = _worksheet_edit_reject_format(format);
        if (err)
            return err;

        if (!formula)
            return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

        if (lxlsx_str_is_empty(formula))
            return LXLSX_ERROR_PARAMETER_IS_EMPTY;

        err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
        if (err)
            return err;

        return lxlsx_edit_set_formula(self->edit_session, self->edit_sheet_name,
                                      row_num, col_num, formula, result);
    }

    if (!formula)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_str_is_empty(formula))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Strip leading "=" from formula. */
    if (formula[0] == '=')
        formula_copy = lxlsx_strdup(formula + 1);
    else
        formula_copy = lxlsx_strdup(formula);

    cell = _new_formula_cell(row_num, col_num, formula_copy, format);
    cell->user_data2 = lxlsx_strdup(result);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a formula with a default result to a cell in Excel .
 */
lxlsx_error
lxlsx_worksheet_write_formula(lxlsx_worksheet *self,
                        lxlsx_row_t row_num,
                        lxlsx_col_t col_num, const char *formula,
                        lxlsx_format *format)
{
    return lxlsx_worksheet_write_formula_num(self, row_num, col_num, formula,
                                       format, 0);
}

/*
 * Internal shared function for various array formula functions.
 */
lxlsx_error
_store_array_formula(lxlsx_worksheet *self,
                     lxlsx_row_t first_row,
                     lxlsx_col_t first_col,
                     lxlsx_row_t last_row,
                     lxlsx_col_t last_col,
                     const char *formula, lxlsx_format *format, double result,
                     uint8_t is_dynamic)
{
    lxlsx_cell *cell;
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    char *formula_copy;
    char *range;
    lxlsx_error err;

    if (_worksheet_is_edit(self))
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    if (!formula)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_str_is_empty(formula))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    /* Check that row and col are valid and store max and min values. */
    err = _check_dimensions(self, first_row, first_col, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    err = _check_dimensions(self, last_row, last_col, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Define the array range. */
    range = calloc(1, LXLSX_MAX_CELL_RANGE_LENGTH);
    RETURN_ON_MEM_ERROR(range, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    if (first_row == last_row && first_col == last_col)
        lxlsx_rowcol_to_cell(range, first_row, first_col);
    else
        lxlsx_rowcol_to_range(range, first_row, first_col, last_row, last_col);

    /* Copy and trip leading "{=" from formula. */
    if (formula[0] == '{')
        if (strlen(formula) >= 2 && formula[1] == '=')
            formula_copy = lxlsx_strdup(formula + 2);
        else
            formula_copy = lxlsx_strdup(formula + 1);
    else
        formula_copy = lxlsx_strdup_formula(formula);

    /* Strip trailing "}" from formula. */
    if (strlen(formula_copy) > 0
        && formula_copy[strlen(formula_copy) - 1] == '}') {
        formula_copy[strlen(formula_copy) - 1] = '\0';
    }

    /* Check for empty formula that started as {=}. */
    if (lxlsx_str_is_empty(formula_copy)) {
        free(formula_copy);
        free(range);
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;
    }

    /* Create a new array formula cell object. */
    cell = _new_array_formula_cell(first_row, first_col,
                                   formula_copy, range, format, is_dynamic);

    cell->formula_result = result;

    _insert_cell(self, first_row, first_col, cell);

    if (is_dynamic)
        self->has_dynamic_functions = LXLSX_TRUE;

    /* Pad out the rest of the area with formatted zeroes. */
    if (!self->optimize) {
        for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
            for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
                if (tmp_row == first_row && tmp_col == first_col)
                    continue;

                lxlsx_worksheet_write_number(self, tmp_row, tmp_col, 0, format);
            }
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Write an array formula with a numerical result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_array_formula_num(lxlsx_worksheet *self,
                                  lxlsx_row_t first_row,
                                  lxlsx_col_t first_col,
                                  lxlsx_row_t last_row,
                                  lxlsx_col_t last_col,
                                  const char *formula,
                                  lxlsx_format *format, double result)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, result,
                                LXLSX_FALSE);
}

/*
 * Write an array formula with a default result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_array_formula(lxlsx_worksheet *self,
                              lxlsx_row_t first_row,
                              lxlsx_col_t first_col,
                              lxlsx_row_t last_row,
                              lxlsx_col_t last_col,
                              const char *formula, lxlsx_format *format)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, 0,
                                LXLSX_FALSE);
}

/*
 * Write a single cell dynamic array formula with a default result to a cell.
 */
lxlsx_error
lxlsx_worksheet_write_dynamic_formula(lxlsx_worksheet *self,
                                lxlsx_row_t row,
                                lxlsx_col_t col,
                                const char *formula, lxlsx_format *format)
{
    return _store_array_formula(self, row, col, row, col, formula, format, 0,
                                LXLSX_TRUE);
}

/*
 * Write a single cell dynamic array formula with a numerical result to a cell.
 */
lxlsx_error
lxlsx_worksheet_write_dynamic_formula_num(lxlsx_worksheet *self,
                                    lxlsx_row_t row,
                                    lxlsx_col_t col,
                                    const char *formula,
                                    lxlsx_format *format, double result)
{
    return _store_array_formula(self, row, col, row, col, formula, format,
                                result, LXLSX_TRUE);
}

/*
 * Write a dynamic array formula with a numerical result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_dynamic_array_formula_num(lxlsx_worksheet *self,
                                          lxlsx_row_t first_row,
                                          lxlsx_col_t first_col,
                                          lxlsx_row_t last_row,
                                          lxlsx_col_t last_col,
                                          const char *formula,
                                          lxlsx_format *format, double result)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, result,
                                LXLSX_TRUE);
}

/*
 * Write a dynamic array formula with a default result to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_dynamic_array_formula(lxlsx_worksheet *self,
                                      lxlsx_row_t first_row,
                                      lxlsx_col_t first_col,
                                      lxlsx_row_t last_row,
                                      lxlsx_col_t last_col,
                                      const char *formula, lxlsx_format *format)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, 0,
                                LXLSX_TRUE);
}

/*
 * Write a blank cell with a format to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_blank(lxlsx_worksheet *self,
                      lxlsx_row_t row_num, lxlsx_col_t col_num,
                      lxlsx_format *format)
{
    lxlsx_cell *cell;
    lxlsx_error err;

    if (_worksheet_is_edit(self))
        return format ? LXLSX_ERROR_FEATURE_NOT_SUPPORTED : LXLSX_NO_ERROR;

    /* Blank cells without formatting are ignored by Excel. */
    if (!format)
        return LXLSX_NO_ERROR;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    cell = _new_blank_cell(row_num, col_num, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a boolean cell with a format to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_boolean(lxlsx_worksheet *self,
                        lxlsx_row_t row_num, lxlsx_col_t col_num,
                        int value, lxlsx_format *format)
{
    lxlsx_cell *cell;
    lxlsx_error err;

    if (_worksheet_is_edit(self)) {
        err = _worksheet_edit_reject_format(format);
        if (err)
            return err;

        err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
        if (err)
            return err;

        return lxlsx_edit_set_boolean(self->edit_session, self->edit_sheet_name,
                                      row_num, col_num, value);
    }

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    cell = _new_boolean_cell(row_num, col_num, value, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a date and or time to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_datetime(lxlsx_worksheet *self,
                         lxlsx_row_t row_num,
                         lxlsx_col_t col_num, lxlsx_datetime *datetime,
                         lxlsx_format *format)
{
    lxlsx_cell *cell;
    double excel_date;
    lxlsx_error err;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    err = lxlsx_datetime_validate(datetime);
    if (err)
        return err;

    excel_date =
        lxlsx_datetime_to_excel_date_with_epoch(datetime, self->use_1904_epoch);

    cell = _new_number_cell(row_num, col_num, excel_date, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a date and or time to a cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_unixtime(lxlsx_worksheet *self,
                         lxlsx_row_t row_num,
                         lxlsx_col_t col_num,
                         int64_t unixtime, lxlsx_format *format)
{
    lxlsx_cell *cell;
    double excel_date;
    lxlsx_error err;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    excel_date =
        lxlsx_unixtime_to_excel_date_with_epoch(unixtime, self->use_1904_epoch);

    cell = _new_number_cell(row_num, col_num, excel_date, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;
}

/*
 * Write a hyperlink/url to an Excel file.
 */
lxlsx_error
lxlsx_worksheet_write_url_opt(lxlsx_worksheet *self,
                        lxlsx_row_t row_num,
                        lxlsx_col_t col_num, const char *url,
                        lxlsx_format *user_format, const char *string,
                        const char *tooltip)
{
    lxlsx_cell *link;
    char *string_copy = NULL;
    char *url_copy = NULL;
    char *url_external = NULL;
    char *url_string = NULL;
    char *tooltip_copy = NULL;
    char *found_string;
    char *tmp_string = NULL;
    lxlsx_format *format = NULL;
    size_t string_size;
    size_t i;
    lxlsx_error err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    enum cell_types link_type = HYPERLINK_URL;

    if (!url || !*url)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    /* Check the Excel limit of URLS per worksheet. */
    if (self->hlink_count > LXLSX_MAX_NUMBER_URLS) {
        LXLSX_WARN("lxlsx_worksheet_write_url()/_opt(): URL ignored since it exceeds "
                 "the maximum number of allowed worksheet URLs (65530).");
        return LXLSX_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED;
    }

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Reset default error condition. */
    err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    /* Set the URI scheme from internal links. */
    found_string = strstr(url, "internal:");
    if (found_string)
        link_type = HYPERLINK_INTERNAL;

    /* Set the URI scheme from external links. */
    found_string = strstr(url, "external:");
    if (found_string)
        link_type = HYPERLINK_EXTERNAL;

    if (string) {
        string_copy = lxlsx_strdup(string);
        GOTO_LABEL_ON_MEM_ERROR(string_copy, mem_error);
    }
    else {
        if (link_type == HYPERLINK_URL) {
            /* Strip the mailto header. */
            found_string = strstr(url, "mailto:");
            if (found_string)
                string_copy = lxlsx_strdup(url + sizeof("mailto"));
            else
                string_copy = lxlsx_strdup(url);
        }
        else {
            string_copy = lxlsx_strdup(url + sizeof("__ternal"));
        }
        GOTO_LABEL_ON_MEM_ERROR(string_copy, mem_error);
    }

    if (url) {
        if (link_type == HYPERLINK_URL)
            url_copy = lxlsx_strdup(url);
        else
            url_copy = lxlsx_strdup(url + sizeof("__ternal"));

        GOTO_LABEL_ON_MEM_ERROR(url_copy, mem_error);
    }

    if (tooltip) {
        tooltip_copy = lxlsx_strdup(tooltip);
        GOTO_LABEL_ON_MEM_ERROR(tooltip_copy, mem_error);
    }

    if (link_type == HYPERLINK_INTERNAL) {
        url_string = lxlsx_strdup(string_copy);
        GOTO_LABEL_ON_MEM_ERROR(url_string, mem_error);
    }

    /* Split url into the link and optional anchor/location. */
    found_string = strchr(url_copy, '#');

    if (found_string) {
        free(url_string);
        url_string = lxlsx_strdup(found_string + 1);
        GOTO_LABEL_ON_MEM_ERROR(url_string, mem_error);

        *found_string = '\0';
    }

    /* Escape the URL. */
    if (link_type == HYPERLINK_URL || link_type == HYPERLINK_EXTERNAL) {
        tmp_string = lxlsx_escape_url_characters(url_copy, LXLSX_FALSE);
        GOTO_LABEL_ON_MEM_ERROR(tmp_string, mem_error);

        free(url_copy);
        url_copy = tmp_string;
    }

    if (link_type == HYPERLINK_EXTERNAL) {
        /* External Workbook links need to be modified into the right format.
         * The URL will look something like "c:\temp\file.xlsx#Sheet!A1". */

        /* For external links change the dir separator from Unix to DOS. */
        for (i = 0; i <= strlen(url_copy); i++)
            if (url_copy[i] == '/')
                url_copy[i] = '\\';

        for (i = 0; i <= strlen(string_copy); i++)
            if (string_copy[i] == '/')
                string_copy[i] = '\\';

        /* Look for Windows style "C:/" link or Windows share "\\" link. */
        found_string = strchr(url_copy, ':');
        if (!found_string)
            found_string = strstr(url_copy, "\\\\");

        if (found_string) {
            /* Add the file:/// URI to the url if non-local. */
            string_size = sizeof("file:///") + strlen(url_copy);
            url_external = calloc(1, string_size);
            GOTO_LABEL_ON_MEM_ERROR(url_external, mem_error);

            lxlsx_snprintf(url_external, string_size, "file:///%s", url_copy);

        }

        /* Convert a ./dir/file.xlsx link to dir/file.xlsx. */
        found_string = strstr(url_copy, ".\\");
        if (found_string == url_copy)
            memmove(url_copy, url_copy + 2, strlen(url_copy) - 1);

        if (url_external) {
            free(url_copy);
            url_copy = lxlsx_strdup(url_external);
            GOTO_LABEL_ON_MEM_ERROR(url_copy, mem_error);

            free(url_external);
            url_external = NULL;
        }

    }

    /* Check if URL exceeds Excel's length limit. */
    if (lxlsx_utf8_strlen(url_copy) > self->max_url_length) {
        LXLSX_WARN_FORMAT2("lxlsx_worksheet_write_url()/_opt(): URL exceeds "
                         "Excel's allowable length of %d characters: %s",
                         self->max_url_length, url_copy);
        err = LXLSX_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED;
        goto mem_error;
    }

    /* Use the default URL format if none is specified. */
    if (!user_format)
        format = self->default_url_format;
    else
        format = user_format;

    if (!self->storing_embedded_image) {
        err =
            lxlsx_worksheet_write_string(self, row_num, col_num, string_copy,
                                   format);
        if (err)
            goto mem_error;
    }

    /* Reset default error condition. */
    err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    link = _new_hyperlink_cell(row_num, col_num, link_type, url_copy,
                               url_string, tooltip_copy);
    GOTO_LABEL_ON_MEM_ERROR(link, mem_error);

    _insert_hyperlink(self, row_num, col_num, link);

    free(string_copy);
    self->hlink_count++;
    return LXLSX_NO_ERROR;

mem_error:
    free(string_copy);
    free(url_copy);
    free(url_external);
    free(url_string);
    free(tooltip_copy);
    return err;
}

/*
 * Write a hyperlink/url to an Excel file.
 */
lxlsx_error
lxlsx_worksheet_write_url(lxlsx_worksheet *self,
                    lxlsx_row_t row_num,
                    lxlsx_col_t col_num, const char *url, lxlsx_format *format)
{
    return lxlsx_worksheet_write_url_opt(self, row_num, col_num, url, format, NULL,
                                   NULL);
}

/*
 * Write a rich string to an Excel file.
 *
 * Rather than duplicate several of the styles.c font xml methods of styles.c
 * and write the data to a memory buffer this function creates a temporary
 * styles object and uses it to write the data to a file. It then reads that
 * data back into memory and closes the file.
 */
lxlsx_error
lxlsx_worksheet_write_rich_string(lxlsx_worksheet *self,
                            lxlsx_row_t row_num,
                            lxlsx_col_t col_num,
                            lxlsx_rich_string_tuple *rich_strings[],
                            lxlsx_format *format)
{
    lxlsx_cell *cell;
    int32_t string_id;
    struct lxlsx_sst_element *lxlsx_sst_element;
    lxlsx_error err;
    uint8_t i;
    long file_size;
    char *rich_string = NULL;
    const char *string_copy = NULL;
    lxlsx_styles *styles = NULL;
    lxlsx_format *default_format = NULL;
    lxlsx_rich_string_tuple *rich_string_tuple = NULL;
    FILE *tmpfile;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Iterate through rich string fragments to check for input errors. */
    i = 0;
    err = LXLSX_NO_ERROR;
    while ((rich_string_tuple = rich_strings[i++]) != NULL) {

        /* Check for NULL or empty strings. */
        if (!rich_string_tuple->string || !*rich_string_tuple->string) {
            err = LXLSX_ERROR_PARAMETER_VALIDATION;
        }
    }

    /* If there are less than 2 fragments it isn't a rich string. */
    if (i <= 2)
        err = LXLSX_ERROR_PARAMETER_VALIDATION;

    if (err)
        return err;

    /* Create a tmp file for the styles object. */
    tmpfile = lxlsx_get_filehandle(&rich_string, NULL, self->tmpdir);
    if (!tmpfile)
        return LXLSX_ERROR_CREATING_TMPFILE;

    /* Create a temp styles object for writing the font data. */
    styles = lxlsx_styles_new();
    GOTO_LABEL_ON_MEM_ERROR(styles, mem_error);
    styles->file = tmpfile;

    /* Create a default format for non-formatted text. */
    default_format = lxlsx_format_new();
    GOTO_LABEL_ON_MEM_ERROR(default_format, mem_error);

    /* Iterate through the rich string fragments and write each one out. */
    i = 0;
    while ((rich_string_tuple = rich_strings[i++]) != NULL) {
        lxlsx_xml_start_tag(tmpfile, "r", NULL);

        if (rich_string_tuple->format) {
            /* Write the user defined font format. */
            lxlsx_styles_write_rich_font(styles, rich_string_tuple->format);
        }
        else {
            /* Write a default font format. Except for the first fragment. */
            if (i > 1)
                lxlsx_styles_write_rich_font(styles, default_format);
        }

        lxlsx_styles_write_string_fragment(styles, rich_string_tuple->string);
        lxlsx_xml_end_tag(tmpfile, "r");
    }

    /* Free the temp objects. */
    lxlsx_styles_free(styles);
    lxlsx_format_free(default_format);

    /* Flush the file. */
    fflush(tmpfile);

    if (!rich_string) {
        /* Read the size to calculate the required memory. */
        file_size = ftell(tmpfile);
        /* Allocate a buffer for the rich string xml data. */
        rich_string = calloc(file_size + 1, 1);
        GOTO_LABEL_ON_MEM_ERROR(rich_string, mem_error);

        /* Rewind the file and read the data into the memory buffer. */
        rewind(tmpfile);
        if (fread((void *) rich_string, file_size, 1, tmpfile) < 1) {
            fclose(tmpfile);
            free((void *) rich_string);
            return LXLSX_ERROR_READING_TMPFILE;
        }
    }

    /* Close the temp file. */
    fclose(tmpfile);

    if (lxlsx_utf8_strlen(rich_string) > LXLSX_STR_MAX) {
        free((void *) rich_string);
        return LXLSX_ERROR_MAX_STRING_LENGTH_EXCEEDED;
    }

    if (!self->optimize) {
        /* Get the SST element and string id. */
        lxlsx_sst_element = lxlsx_get_sst_index(self->sst, rich_string, LXLSX_TRUE);
        free((void *) rich_string);

        if (!lxlsx_sst_element)
            return LXLSX_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = lxlsx_sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id,
                                lxlsx_sst_element->string, format);
    }
    else {
        /* Look for and escape control chars in the string. */
        if (lxlsx_has_control_characters(rich_string)) {
            string_copy = lxlsx_escape_control_characters(rich_string);
            free((void *) rich_string);
        }
        else {
            string_copy = rich_string;
        }
        cell = _new_inline_rich_string_cell(row_num, col_num, string_copy,
                                            format);
    }

    _insert_cell(self, row_num, col_num, cell);

    return LXLSX_NO_ERROR;

mem_error:
    lxlsx_styles_free(styles);
    lxlsx_format_free(default_format);
    fclose(tmpfile);

    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Write a comment to a worksheet cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_comment_opt(lxlsx_worksheet *self,
                            lxlsx_row_t row_num, lxlsx_col_t col_num,
                            const char *text, lxlsx_comment_options *options)
{
    lxlsx_cell *cell;
    lxlsx_error err;
    lxlsx_vml_obj *comment;

    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    if (!text)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    if (lxlsx_str_is_empty(text))
        return LXLSX_ERROR_PARAMETER_IS_EMPTY;

    if (lxlsx_utf8_strlen(text) > LXLSX_STR_MAX)
        return LXLSX_ERROR_MAX_STRING_LENGTH_EXCEEDED;

    comment = calloc(1, sizeof(lxlsx_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(comment, mem_error);

    comment->text = lxlsx_strdup(text);
    GOTO_LABEL_ON_MEM_ERROR(comment->text, mem_error);

    comment->row = row_num;
    comment->col = col_num;

    cell = _new_comment_cell(row_num, col_num, comment);
    GOTO_LABEL_ON_MEM_ERROR(cell, mem_error);

    _insert_comment(self, row_num, col_num, cell);

    /* Set user and default parameters for the comment. */
    _get_comment_params(comment, options);

    self->has_vml = LXLSX_TRUE;
    self->has_comments = LXLSX_TRUE;

    /* Insert a placeholder in the cell RB table in the same position so
     * that the worksheet row "spans" calculations are correct. */
    _insert_cell_placeholder(self, row_num, col_num);

    return LXLSX_NO_ERROR;

mem_error:
    if (comment)
        _free_vml_object(comment);

    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Write a comment to a worksheet cell in Excel.
 */
lxlsx_error
lxlsx_worksheet_write_comment(lxlsx_worksheet *self,
                        lxlsx_row_t row_num, lxlsx_col_t col_num,
                        const char *string)
{
    return lxlsx_worksheet_write_comment_opt(self, row_num, col_num, string, NULL);
}

/*
 * Set the properties of a single column or a range of columns with options.
 */
lxlsx_error
lxlsx_worksheet_set_column_opt(lxlsx_worksheet *self,
                         lxlsx_col_t firstcol,
                         lxlsx_col_t lastcol,
                         double width,
                         lxlsx_format *format,
                         lxlsx_row_col_options *user_options)
{
    lxlsx_col_options *copied_options;
    uint8_t ignore_row = LXLSX_TRUE;
    uint8_t ignore_col = LXLSX_TRUE;
    uint8_t hidden = LXLSX_FALSE;
    uint8_t level = 0;
    uint8_t collapsed = LXLSX_FALSE;
    lxlsx_col_t col;
    lxlsx_error err;

    if (_worksheet_is_edit(self))
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    if (user_options) {
        hidden = user_options->hidden;
        level = user_options->level;
        collapsed = user_options->collapsed;
    }

    /* Ensure second col is larger than first. */
    if (firstcol > lastcol) {
        lxlsx_col_t tmp = firstcol;
        firstcol = lastcol;
        lastcol = tmp;
    }

    /* Ensure that the cols are valid and store max and min values.
     * NOTE: The check shouldn't modify the row dimensions and should only
     *       modify the column dimensions in certain cases. */
    if (format != NULL || (width != LXLSX_DEF_COL_WIDTH && hidden))
        ignore_col = LXLSX_FALSE;

    err = _check_dimensions(self, 0, firstcol, ignore_row, ignore_col);

    if (!err)
        err = _check_dimensions(self, 0, lastcol, ignore_row, ignore_col);

    if (err)
        return err;

    /* Resize the col_options array if required. */
    if (firstcol >= self->col_options_max) {
        lxlsx_col_t col_tmp;
        lxlsx_col_t old_size = self->col_options_max;
        lxlsx_col_t new_size = _next_power_of_two(firstcol + 1);
        lxlsx_col_options **new_ptr = realloc(self->col_options,
                                            new_size *
                                            sizeof(lxlsx_col_options *));

        if (new_ptr) {
            for (col_tmp = old_size; col_tmp < new_size; col_tmp++)
                new_ptr[col_tmp] = NULL;

            self->col_options = new_ptr;
            self->col_options_max = new_size;
        }
        else {
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Resize the col_formats array if required. */
    if (lastcol >= self->col_formats_max) {
        lxlsx_col_t col;
        lxlsx_col_t old_size = self->col_formats_max;
        lxlsx_col_t new_size = _next_power_of_two(lastcol + 1);
        lxlsx_format **new_ptr = realloc(self->col_formats,
                                       new_size * sizeof(lxlsx_format *));

        if (new_ptr) {
            for (col = old_size; col < new_size; col++)
                new_ptr[col] = NULL;

            self->col_formats = new_ptr;
            self->col_formats_max = new_size;
        }
        else {
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Store the column options. */
    copied_options = calloc(1, sizeof(lxlsx_col_options));
    RETURN_ON_MEM_ERROR(copied_options, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Ensure the level is <= 7). */
    if (level > 7)
        level = 7;

    if (level > self->outline_col_level)
        self->outline_col_level = level;

    /* Set the column properties. */
    copied_options->firstcol = firstcol;
    copied_options->lastcol = lastcol;
    copied_options->width = width;
    copied_options->format = format;
    copied_options->hidden = hidden;
    copied_options->level = level;
    copied_options->collapsed = collapsed;

    free(self->col_options[firstcol]);
    self->col_options[firstcol] = copied_options;

    /* Store the column formats for use when writing cell data. */
    for (col = firstcol; col <= lastcol; col++) {
        self->col_formats[col] = format;
    }

    /* Store the column change to allow optimizations. */
    self->col_size_changed = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set the properties of a single column or a range of columns.
 */
lxlsx_error
lxlsx_worksheet_set_column(lxlsx_worksheet *self,
                     lxlsx_col_t firstcol,
                     lxlsx_col_t lastcol, double width, lxlsx_format *format)
{
    return lxlsx_worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    NULL);
}

/*
 * Set the properties of a single column or a range of columns, with the
 * width in pixels.
 */
lxlsx_error
lxlsx_worksheet_set_column_pixels(lxlsx_worksheet *self,
                            lxlsx_col_t firstcol,
                            lxlsx_col_t lastcol,
                            uint32_t pixels, lxlsx_format *format)
{
    double width = _pixels_to_width(pixels);

    return lxlsx_worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    NULL);
}

/*
 * Set the properties of a single column or a range of columns with options,
 * with the width in pixels.
 */
lxlsx_error
lxlsx_worksheet_set_column_pixels_opt(lxlsx_worksheet *self,
                                lxlsx_col_t firstcol,
                                lxlsx_col_t lastcol,
                                uint32_t pixels,
                                lxlsx_format *format,
                                lxlsx_row_col_options *user_options)
{
    double width = _pixels_to_width(pixels);

    return lxlsx_worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    user_options);
}

/*
 * Set the properties of a row with options.
 */
lxlsx_error
lxlsx_worksheet_set_row_opt(lxlsx_worksheet *self,
                      lxlsx_row_t row_num,
                      double height,
                      lxlsx_format *format, lxlsx_row_col_options *user_options)
{

    lxlsx_col_t min_col;
    uint8_t hidden = LXLSX_FALSE;
    uint8_t level = 0;
    uint8_t collapsed = LXLSX_FALSE;
    lxlsx_row *row;
    lxlsx_error err;

    if (_worksheet_is_edit(self))
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    if (user_options) {
        hidden = user_options->hidden;
        level = user_options->level;
        collapsed = user_options->collapsed;
    }

    /* Use minimum col in _check_dimensions(). */
    if (self->dim_colmin != LXLSX_COL_MAX)
        min_col = self->dim_colmin;
    else
        min_col = 0;

    err = _check_dimensions(self, row_num, min_col, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* If the height is 0 the row is hidden and the height is the default. */
    if (height == 0) {
        hidden = LXLSX_TRUE;
        height = self->default_row_height;
    }

    /* Ensure the level is <= 7). */
    if (level > 7)
        level = 7;

    if (level > self->outline_row_level)
        self->outline_row_level = level;

    /* Store the row properties. */
    row = _get_row(self, row_num);

    row->height = height;
    row->format = format;
    row->hidden = hidden;
    row->level = level;
    row->collapsed = collapsed;
    row->row_changed = LXLSX_TRUE;

    if (height != self->default_row_height)
        row->height_changed = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set the properties of a row.
 */
lxlsx_error
lxlsx_worksheet_set_row(lxlsx_worksheet *self,
                  lxlsx_row_t row_num, double height, lxlsx_format *format)
{
    return lxlsx_worksheet_set_row_opt(self, row_num, height, format, NULL);
}

/*
 * Set the properties of a row, with the height in pixels.
 */
lxlsx_error
lxlsx_worksheet_set_row_pixels(lxlsx_worksheet *self,
                         lxlsx_row_t row_num, uint32_t pixels,
                         lxlsx_format *format)
{
    double height = _pixels_to_height(pixels);

    return lxlsx_worksheet_set_row_opt(self, row_num, height, format, NULL);
}

/*
 * Set the properties of a row with options, with the height in pixels.
 */
lxlsx_error
lxlsx_worksheet_set_row_pixels_opt(lxlsx_worksheet *self,
                             lxlsx_row_t row_num,
                             uint32_t pixels,
                             lxlsx_format *format,
                             lxlsx_row_col_options *user_options)
{
    double height = _pixels_to_height(pixels);

    return lxlsx_worksheet_set_row_opt(self, row_num, height, format, user_options);
}

/*
 * Merge a range of cells. The first cell should contain the data and the others
 * should be blank. All cells should contain the same format.
 */
lxlsx_error
lxlsx_worksheet_merge_range(lxlsx_worksheet *self, lxlsx_row_t first_row,
                      lxlsx_col_t first_col, lxlsx_row_t last_row,
                      lxlsx_col_t last_col, const char *string,
                      lxlsx_format *format)
{
    lxlsx_merged_range *merged_range;
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_error err;

    if (_worksheet_is_edit(self))
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    /* Excel doesn't allow a single cell to be merged */
    if (first_row == last_row && first_col == last_col)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(self, last_row, last_col, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Store the merge range. */
    merged_range = calloc(1, sizeof(lxlsx_merged_range));
    RETURN_ON_MEM_ERROR(merged_range, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    merged_range->first_row = first_row;
    merged_range->first_col = first_col;
    merged_range->last_row = last_row;
    merged_range->last_col = last_col;

    STAILQ_INSERT_TAIL(self->merged_ranges, merged_range, list_pointers);
    self->merged_range_count++;

    /* Write the first cell */
    lxlsx_worksheet_write_string(self, first_row, first_col, string, format);

    /* Pad out the rest of the area with formatted blank cells. */
    for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
        for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
            if (tmp_row == first_row && tmp_col == first_col)
                continue;

            lxlsx_worksheet_write_blank(self, tmp_row, tmp_col, format);
        }
    }

    return LXLSX_NO_ERROR;
}

/*
 * Set the autofilter area in the worksheet.
 */
lxlsx_error
lxlsx_worksheet_autofilter(lxlsx_worksheet *self, lxlsx_row_t first_row,
                     lxlsx_col_t first_col, lxlsx_row_t last_row,
                     lxlsx_col_t last_col)
{
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_error err;
    lxlsx_filter_rule_obj **filter_rules;
    lxlsx_col_t num_filter_rules;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(self, last_row, last_col, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Create a array to hold filter rules. */
    self->autofilter.in_use = LXLSX_FALSE;
    self->autofilter.has_rules = LXLSX_FALSE;
    _free_filter_rules(self);
    num_filter_rules = last_col - first_col + 1;
    filter_rules = calloc(num_filter_rules, sizeof(lxlsx_filter_rule_obj *));
    RETURN_ON_MEM_ERROR(filter_rules, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    self->autofilter.in_use = LXLSX_TRUE;
    self->autofilter.first_row = first_row;
    self->autofilter.first_col = first_col;
    self->autofilter.last_row = last_row;
    self->autofilter.last_col = last_col;

    self->filter_rules = filter_rules;
    self->num_filter_rules = num_filter_rules;

    return LXLSX_NO_ERROR;
}

/*
 * Set a autofilter rule for a filter column.
 */
lxlsx_error
lxlsx_worksheet_filter_column(lxlsx_worksheet *self, lxlsx_col_t col,
                        lxlsx_filter_rule *rule)
{
    lxlsx_filter_rule_obj *rule_obj;
    uint16_t rule_index;

    if (!rule) {
        LXLSX_WARN("lxlsx_worksheet_filter_column(): rule parameter cannot be NULL");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXLSX_FALSE) {
        LXLSX_WARN("lxlsx_worksheet_filter_column(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use lxlsx_worksheet_autofilter() first.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXLSX_WARN_FORMAT3("lxlsx_worksheet_filter_column(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxlsx_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    rule_obj->col_num = rule_index;
    rule_obj->type = LXLSX_FILTER_TYPE_SINGLE;
    rule_obj->criteria1 = rule->criteria;
    rule_obj->value1 = rule->value;

    if (rule_obj->criteria1 != LXLSX_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value1_string = lxlsx_strdup(rule->value_string);
    }
    else {
        rule_obj->criteria1 = LXLSX_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value1_string = lxlsx_strdup(" ");
    }

    if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXLSX_TRUE;

    _set_custom_filter(rule_obj);

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXLSX_TRUE;
    self->autofilter.has_rules = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set two autofilter rules for a filter column.
 */
lxlsx_error
lxlsx_worksheet_filter_column2(lxlsx_worksheet *self, lxlsx_col_t col,
                         lxlsx_filter_rule *rule1, lxlsx_filter_rule *rule2,
                         uint8_t and_or)
{
    lxlsx_filter_rule_obj *rule_obj;
    uint16_t rule_index;

    if (!rule1 || !rule2) {
        LXLSX_WARN("lxlsx_worksheet_filter_column2(): rule parameter cannot be NULL");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXLSX_FALSE) {
        LXLSX_WARN("lxlsx_worksheet_filter_column2(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use lxlsx_worksheet_autofilter() first.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXLSX_WARN_FORMAT3("lxlsx_worksheet_filter_column2(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxlsx_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    if (and_or == LXLSX_FILTER_AND)
        rule_obj->type = LXLSX_FILTER_TYPE_AND;
    else
        rule_obj->type = LXLSX_FILTER_TYPE_OR;

    rule_obj->col_num = rule_index;

    rule_obj->criteria1 = rule1->criteria;
    rule_obj->value1 = rule1->value;

    rule_obj->criteria2 = rule2->criteria;
    rule_obj->value2 = rule2->value;

    if (rule_obj->criteria1 != LXLSX_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value1_string = lxlsx_strdup(rule1->value_string);
    }
    else {
        rule_obj->criteria1 = LXLSX_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value1_string = lxlsx_strdup(" ");
    }

    if (rule_obj->criteria2 != LXLSX_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value2_string = lxlsx_strdup(rule2->value_string);
    }
    else {
        rule_obj->criteria2 = LXLSX_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value2_string = lxlsx_strdup(" ");
    }

    if (rule_obj->criteria1 == LXLSX_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXLSX_TRUE;

    if (rule_obj->criteria2 == LXLSX_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXLSX_TRUE;

    _set_custom_filter(rule_obj);

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXLSX_TRUE;
    self->autofilter.has_rules = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set two autofilter rules for a filter column.
 */
lxlsx_error
lxlsx_worksheet_filter_list(lxlsx_worksheet *self, lxlsx_col_t col, const char **list)
{
    lxlsx_filter_rule_obj *rule_obj;
    uint16_t rule_index;
    uint8_t has_blanks = LXLSX_FALSE;
    uint16_t num_filters = 0;
    uint16_t input_list_index;
    uint16_t rule_obj_list_index;
    const char *str;
    char **tmp_list;

    if (!list) {
        LXLSX_WARN("lxlsx_worksheet_filter_list(): list parameter cannot be NULL");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXLSX_FALSE) {
        LXLSX_WARN("lxlsx_worksheet_filter_list(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use lxlsx_worksheet_autofilter() first.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXLSX_WARN_FORMAT3("lxlsx_worksheet_filter_list(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Count the number of non "Blanks" strings in the input list. */
    input_list_index = 0;
    while ((str = list[input_list_index]) != NULL) {
        if (strncmp(str, "Blanks", 6) == 0)
            has_blanks = LXLSX_TRUE;
        else
            num_filters++;

        input_list_index++;
    }

    /* There should be at least one filter string. */
    if (num_filters == 0) {
        LXLSX_WARN("lxlsx_worksheet_filter_list(): "
                 "list must have at least 1 non-blanks item.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxlsx_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    tmp_list = calloc(num_filters + 1, sizeof(char *));
    GOTO_LABEL_ON_MEM_ERROR(tmp_list, mem_error);

    /* Copy input list (without any "Blanks" command) to an internal list. */
    input_list_index = 0;
    rule_obj_list_index = 0;
    while ((str = list[input_list_index]) != NULL) {
        if (strncmp(str, "Blanks", 6) != 0) {
            tmp_list[rule_obj_list_index] = lxlsx_strdup(str);
            rule_obj_list_index++;
        }

        input_list_index++;
    }

    rule_obj->list = tmp_list;
    rule_obj->num_list_filters = num_filters;
    rule_obj->is_custom = LXLSX_FALSE;
    rule_obj->col_num = rule_index;
    rule_obj->type = LXLSX_FILTER_TYPE_STRING_LIST;
    rule_obj->has_blanks = has_blanks;

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXLSX_TRUE;
    self->autofilter.has_rules = LXLSX_TRUE;

    return LXLSX_NO_ERROR;

mem_error:
    free(rule_obj);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

}

/*
 * Add an Excel table to the worksheet.
 */
lxlsx_error
lxlsx_worksheet_add_table(lxlsx_worksheet *self, lxlsx_row_t first_row,
                    lxlsx_col_t first_col, lxlsx_row_t last_row,
                    lxlsx_col_t last_col, lxlsx_table_options *user_options)
{
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_col_t num_cols;
    lxlsx_error err;
    lxlsx_table_obj *lxlsx_table_obj;
    lxlsx_table_column **columns;

    if (self->optimize) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_add_table(): "
                        "worksheet tables aren't supported in "
                        "'constant_memory' mode");
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    }

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(self, last_row, last_col, LXLSX_TRUE, LXLSX_TRUE);
    if (err)
        return err;

    num_cols = last_col - first_col + 1;

    /* Check that there are sufficient data rows. */
    err = _check_table_rows(first_row, last_row, user_options);
    if (err)
        return err;

    /* Check that the the table name is valid. */
    err = _check_table_name(user_options);
    if (err)
        return err;

    /* Create a table object to copy from the user options. */
    lxlsx_table_obj = calloc(1, sizeof(*lxlsx_table_obj));
    RETURN_ON_MEM_ERROR(lxlsx_table_obj, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    columns = calloc(num_cols, sizeof(lxlsx_table_column *));
    GOTO_LABEL_ON_MEM_ERROR(columns, error);

    lxlsx_table_obj->columns = columns;
    lxlsx_table_obj->num_cols = num_cols;
    lxlsx_table_obj->first_row = first_row;
    lxlsx_table_obj->first_col = first_col;
    lxlsx_table_obj->last_row = last_row;
    lxlsx_table_obj->last_col = last_col;

    err = _set_default_table_columns(lxlsx_table_obj);
    if (err)
        goto error;

    /* Create the table range. */
    lxlsx_rowcol_to_range(lxlsx_table_obj->sqref,
                        first_row, first_col, last_row, last_col);
    lxlsx_rowcol_to_range(lxlsx_table_obj->filter_sqref,
                        first_row, first_col, last_row, last_col);

    /* Validate and copy user options to an internal object. */
    if (user_options) {

        _check_and_copy_table_style(lxlsx_table_obj, user_options);

        lxlsx_table_obj->total_row = user_options->total_row;
        lxlsx_table_obj->last_column = user_options->last_column;
        lxlsx_table_obj->first_column = user_options->first_column;
        lxlsx_table_obj->no_autofilter = user_options->no_autofilter;
        lxlsx_table_obj->no_header_row = user_options->no_header_row;
        lxlsx_table_obj->no_banded_rows = user_options->no_banded_rows;
        lxlsx_table_obj->banded_columns = user_options->banded_columns;

        if (user_options->no_header_row)
            lxlsx_table_obj->no_autofilter = LXLSX_TRUE;

        if (user_options->columns) {
            err = _set_custom_table_columns(lxlsx_table_obj, user_options);
            if (err)
                goto error;
        }

        if (user_options->total_row) {
            lxlsx_rowcol_to_range(lxlsx_table_obj->filter_sqref,
                                first_row, first_col, last_row - 1, last_col);
        }

        if (user_options->name) {
            lxlsx_table_obj->name = lxlsx_strdup(user_options->name);
            if (!lxlsx_table_obj->name) {
                err = LXLSX_ERROR_MEMORY_MALLOC_FAILED;
                goto error;
            }
        }
    }

    _write_table_column_data(self, lxlsx_table_obj);

    STAILQ_INSERT_TAIL(self->lxlsx_table_objs, lxlsx_table_obj, list_pointers);
    self->lxlsx_table_count++;

    return LXLSX_NO_ERROR;

error:
    _free_worksheet_table(lxlsx_table_obj);
    return err;

}

/*
 * Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
void
lxlsx_worksheet_select(lxlsx_worksheet *self)
{
    self->selected = LXLSX_TRUE;

    /* Selected worksheet can't be hidden. */
    self->hidden = LXLSX_FALSE;
}

/*
 * Set this worksheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
void
lxlsx_worksheet_activate(lxlsx_worksheet *self)
{
    self->selected = LXLSX_TRUE;
    self->active = LXLSX_TRUE;

    /* Active worksheet can't be hidden. */
    self->hidden = LXLSX_FALSE;

    *self->active_sheet = self->index;
}

/*
 * Set this worksheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
void
lxlsx_worksheet_set_first_sheet(lxlsx_worksheet *self)
{
    /* Active worksheet can't be hidden. */
    self->hidden = LXLSX_FALSE;

    *self->first_sheet = self->index;
}

/*
 * Hide this worksheet.
 */
void
lxlsx_worksheet_hide(lxlsx_worksheet *self)
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
 * Set which cell or cells are selected in a worksheet.
 */
lxlsx_error
lxlsx_worksheet_set_selection(lxlsx_worksheet *self,
                        lxlsx_row_t first_row, lxlsx_col_t first_col,
                        lxlsx_row_t last_row, lxlsx_col_t last_col)
{
    lxlsx_selection *selection;
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_error err;
    char active_cell[LXLSX_MAX_CELL_RANGE_LENGTH];
    char sqref[LXLSX_MAX_CELL_RANGE_LENGTH];

    /* Only allow selection to be set once to avoid freeing/re-creating it. */
    if (!STAILQ_EMPTY(self->selections))
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    /* Excel doesn't set a selection for cell A1 since it is the default. */
    if (first_row == 0 && first_col == 0 && last_row == 0 && last_col == 0)
        return LXLSX_NO_ERROR;

    selection = calloc(1, sizeof(lxlsx_selection));
    RETURN_ON_MEM_ERROR(selection, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Check that row and col are valid without storing. */
    err = _check_dimensions(self, first_row, first_col, LXLSX_TRUE, LXLSX_TRUE);
    if (err) {
        free(selection);
        return err;
    }

    err = _check_dimensions(self, last_row, last_col, LXLSX_TRUE, LXLSX_TRUE);
    if (err) {
        free(selection);
        return err;
    }

    /* Set the cell range selection. Do this before swapping max/min to  */
    /* allow the selection direction to be reversed. */
    lxlsx_rowcol_to_cell(active_cell, first_row, first_col);

    /* Swap last row/col for first row/col if necessary. */
    if (first_row > last_row) {
        tmp_row = first_row;
        first_row = last_row;
        last_row = tmp_row;
    }

    if (first_col > last_col) {
        tmp_col = first_col;
        first_col = last_col;
        last_col = tmp_col;
    }

    /* If the first and last cell are the same write a single cell. */
    if ((first_row == last_row) && (first_col == last_col))
        lxlsx_rowcol_to_cell(sqref, first_row, first_col);
    else
        lxlsx_rowcol_to_range(sqref, first_row, first_col, last_row, last_col);

    lxlsx_strcpy(selection->pane, "");
    lxlsx_strcpy(selection->active_cell, active_cell);
    lxlsx_strcpy(selection->sqref, sqref);

    STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);

    return LXLSX_NO_ERROR;
}

/*
 * Set the first visible cell at the top left of the worksheet.
 */
void
lxlsx_worksheet_set_top_left_cell(lxlsx_worksheet *self, lxlsx_row_t row, lxlsx_col_t col)
{
    if (row == 0 && col == 0)
        return;

    lxlsx_rowcol_to_cell(self->top_left_cell, row, col);
}

/*
 * Set panes and mark them as frozen. With extra options.
 */
void
lxlsx_worksheet_freeze_panes_opt(lxlsx_worksheet *self,
                           lxlsx_row_t first_row, lxlsx_col_t first_col,
                           lxlsx_row_t top_row, lxlsx_col_t left_col,
                           uint8_t type)
{
    self->panes.first_row = first_row;
    self->panes.first_col = first_col;
    self->panes.top_row = top_row;
    self->panes.left_col = left_col;
    self->panes.x_split = 0.0;
    self->panes.y_split = 0.0;

    if (type)
        self->panes.type = FREEZE_SPLIT_PANES;
    else
        self->panes.type = FREEZE_PANES;
}

/*
 * Set panes and mark them as frozen.
 */
void
lxlsx_worksheet_freeze_panes(lxlsx_worksheet *self,
                       lxlsx_row_t first_row, lxlsx_col_t first_col)
{
    lxlsx_worksheet_freeze_panes_opt(self, first_row, first_col,
                               first_row, first_col, 0);
}

/*
 * Set panes and mark them as split.With extra options.
 */
void
lxlsx_worksheet_split_panes_opt(lxlsx_worksheet *self,
                          double y_split, double x_split,
                          lxlsx_row_t top_row, lxlsx_col_t left_col)
{
    self->panes.first_row = 0;
    self->panes.first_col = 0;
    self->panes.top_row = top_row;
    self->panes.left_col = left_col;
    self->panes.x_split = x_split;
    self->panes.y_split = y_split;
    self->panes.type = SPLIT_PANES;
}

/*
 * Set panes and mark them as split.
 */
void
lxlsx_worksheet_split_panes(lxlsx_worksheet *self, double y_split, double x_split)
{
    lxlsx_worksheet_split_panes_opt(self, y_split, x_split, 0, 0);
}

/*
 * Set the page orientation as portrait.
 */
void
lxlsx_worksheet_set_portrait(lxlsx_worksheet *self)
{
    self->orientation = LXLSX_PORTRAIT;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Set the page orientation as landscape.
 */
void
lxlsx_worksheet_set_landscape(lxlsx_worksheet *self)
{
    self->orientation = LXLSX_LANDSCAPE;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Set the page view mode for Mac Excel.
 */
void
lxlsx_worksheet_set_page_view(lxlsx_worksheet *self)
{
    self->page_view = LXLSX_TRUE;
}

/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
void
lxlsx_worksheet_set_paper(lxlsx_worksheet *self, uint8_t paper_size)
{
    if (paper_size > 118) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_paper(): invalid paper size: %d. "
                         "Valid range is 0-118", paper_size);
        return;
    }

    self->paper_size = paper_size;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Set the order in which pages are printed.
 */
void
lxlsx_worksheet_print_across(lxlsx_worksheet *self)
{
    self->page_order = LXLSX_PRINT_ACROSS;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Set all the page margins in inches.
 */
void
lxlsx_worksheet_set_margins(lxlsx_worksheet *self, double left, double right,
                      double top, double bottom)
{

    if (left >= 0)
        self->margin_left = left;

    if (right >= 0)
        self->margin_right = right;

    if (top >= 0)
        self->margin_top = top;

    if (bottom >= 0)
        self->margin_bottom = bottom;
}

/*
 * Set the page header caption and options.
 */
lxlsx_error
lxlsx_worksheet_set_header_opt(lxlsx_worksheet *self, const char *string,
                         lxlsx_header_footer_options *options)
{
    lxlsx_error err;
    char *tmp_header;
    char *found_string;
    char *offset_string;
    uint8_t placeholder_count = 0;
    uint8_t image_count = 0;

    if (!string) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_set_header_opt/footer_opt(): "
                        "header/footer string cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_utf8_strlen(string) > LXLSX_HEADER_FOOTER_MAX) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_set_header_opt/footer_opt(): "
                        "header/footer string exceeds Excel's limit of "
                        "255 characters.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    tmp_header = lxlsx_strdup(string);
    RETURN_ON_MEM_ERROR(tmp_header, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Replace &[Picture] with &G which is used internally by Excel. */
    while ((found_string = strstr(tmp_header, "&[Picture]"))) {
        found_string++;
        *found_string = 'G';

        do {
            found_string++;
            offset_string = found_string + sizeof("Picture");
            *found_string = *offset_string;
        } while (*offset_string);
    }

    /* Count &G placeholders and ensure there are sufficient images. */
    found_string = tmp_header;
    while (*found_string) {
        if (*found_string == '&' && *(found_string + 1) == 'G')
            placeholder_count++;
        found_string++;
    }

    if (placeholder_count > 0 && !options) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                         "the number of &G/&[Picture] placeholders in option "
                         "string \"%s\" does not match the number of supplied "
                         "images.", string);

        free(tmp_header);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous header string so we can overwrite it. */
    free(self->header);
    self->header = NULL;

    if (options) {
        /* Ensure there are enough images to match the placeholders. There is
         * a potential bug where there are sufficient images but in the wrong
         * positions but we don't currently try to deal with that.*/
        if (options->image_left)
            image_count++;
        if (options->image_center)
            image_count++;
        if (options->image_right)
            image_count++;

        if (placeholder_count != image_count) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                             "the number of &G/&[Picture] placeholders in option "
                             "string \"%s\" does not match the number of supplied "
                             "images.", string);

            free(tmp_header);
            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }

        /* Free any existing header image objects. */
        _free_object_properties(self->header_left_object_props);
        _free_object_properties(self->header_center_object_props);
        _free_object_properties(self->header_right_object_props);

        if (options->margin > 0.0)
            self->margin_header = options->margin;

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_left,
                                                 HEADER_LEFT);
        if (err) {
            free(tmp_header);
            return err;
        }

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_center,
                                                 HEADER_CENTER);
        if (err) {
            free(tmp_header);
            return err;
        }

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_right,
                                                 HEADER_RIGHT);
        if (err) {
            free(tmp_header);
            return err;
        }
    }

    self->header = tmp_header;
    self->header_footer_changed = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set the page footer caption and options.
 */
lxlsx_error
lxlsx_worksheet_set_footer_opt(lxlsx_worksheet *self, const char *string,
                         lxlsx_header_footer_options *options)
{
    lxlsx_error err;
    char *tmp_footer;
    char *found_string;
    char *offset_string;
    uint8_t placeholder_count = 0;
    uint8_t image_count = 0;

    if (!string) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_set_header_opt/footer_opt(): "
                        "header/footer string cannot be NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxlsx_utf8_strlen(string) > LXLSX_HEADER_FOOTER_MAX) {
        LXLSX_WARN_FORMAT("lxlsx_worksheet_set_header_opt/footer_opt(): "
                        "header/footer string exceeds Excel's limit of "
                        "255 characters.");
        return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    tmp_footer = lxlsx_strdup(string);
    RETURN_ON_MEM_ERROR(tmp_footer, LXLSX_ERROR_MEMORY_MALLOC_FAILED);

    /* Replace &[Picture] with &G which is used internally by Excel. */
    while ((found_string = strstr(tmp_footer, "&[Picture]"))) {
        found_string++;
        *found_string = 'G';

        do {
            found_string++;
            offset_string = found_string + sizeof("Picture");
            *found_string = *offset_string;
        } while (*offset_string);
    }

    /* Count &G placeholders and ensure there are sufficient images. */
    found_string = tmp_footer;
    while (*found_string) {
        if (*found_string == '&' && *(found_string + 1) == 'G')
            placeholder_count++;
        found_string++;
    }

    if (placeholder_count > 0 && !options) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                         "the number of &G/&[Picture] placeholders in option "
                         "string \"%s\" does not match the number of supplied "
                         "images.", string);

        free(tmp_footer);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous footer string so we can overwrite it. */
    free(self->footer);
    self->footer = NULL;

    if (options) {
        /* Ensure there are enough images to match the placeholders. There is
         * a potential bug where there are sufficient images but in the wrong
         * positions but we don't currently try to deal with that.*/
        if (options->image_left)
            image_count++;
        if (options->image_center)
            image_count++;
        if (options->image_right)
            image_count++;

        if (placeholder_count != image_count) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_header_opt/footer_opt(): "
                             "the number of &G/&[Picture] placeholders in option "
                             "string \"%s\" does not match the number of supplied "
                             "images.", string);

            free(tmp_footer);
            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }

        /* Free any existing footer image objects. */
        _free_object_properties(self->footer_left_object_props);
        _free_object_properties(self->footer_center_object_props);
        _free_object_properties(self->footer_right_object_props);

        if (options->margin > 0.0)
            self->margin_footer = options->margin;

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_left,
                                                 FOOTER_LEFT);
        if (err) {
            free(tmp_footer);
            return err;
        }

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_center,
                                                 FOOTER_CENTER);
        if (err) {
            free(tmp_footer);
            return err;
        }

        err = _worksheet_set_header_footer_image(self,
                                                 options->image_right,
                                                 FOOTER_RIGHT);
        if (err) {
            free(tmp_footer);
            return err;
        }
    }

    self->footer = tmp_footer;
    self->header_footer_changed = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Set the page header caption.
 */
lxlsx_error
lxlsx_worksheet_set_header(lxlsx_worksheet *self, const char *string)
{
    return lxlsx_worksheet_set_header_opt(self, string, NULL);
}

/*
 * Set the page footer caption.
 */
lxlsx_error
lxlsx_worksheet_set_footer(lxlsx_worksheet *self, const char *string)
{
    return lxlsx_worksheet_set_footer_opt(self, string, NULL);
}

/*
 * Set the option to show/hide gridlines on the screen and the printed page.
 */
void
lxlsx_worksheet_gridlines(lxlsx_worksheet *self, uint8_t option)
{
    if (option == LXLSX_HIDE_ALL_GRIDLINES) {
        self->print_gridlines = 0;
        self->screen_gridlines = 0;
    }

    if (option & LXLSX_SHOW_SCREEN_GRIDLINES) {
        self->screen_gridlines = 1;
    }

    if (option & LXLSX_SHOW_PRINT_GRIDLINES) {
        self->print_gridlines = 1;
        self->print_options_changed = 1;
    }
}

/*
 * Center the page horizontally.
 */
void
lxlsx_worksheet_center_horizontally(lxlsx_worksheet *self)
{
    self->print_options_changed = 1;
    self->hcenter = 1;
}

/*
 * Center the page horizontally.
 */
void
lxlsx_worksheet_center_vertically(lxlsx_worksheet *self)
{
    self->print_options_changed = 1;
    self->vcenter = 1;
}

/*
 * Set the option to print the row and column headers on the printed page.
 */
void
lxlsx_worksheet_print_row_col_headers(lxlsx_worksheet *self)
{
    self->print_headers = 1;
    self->print_options_changed = 1;
}

/*
 * Set the rows to repeat at the top of each printed page.
 */
lxlsx_error
lxlsx_worksheet_repeat_rows(lxlsx_worksheet *self, lxlsx_row_t first_row,
                      lxlsx_row_t last_row)
{
    lxlsx_row_t tmp_row;
    lxlsx_error err;

    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }

    err = _check_dimensions(self, last_row, 0, LXLSX_IGNORE, LXLSX_IGNORE);
    if (err)
        return err;

    self->repeat_rows.in_use = LXLSX_TRUE;
    self->repeat_rows.first_row = first_row;
    self->repeat_rows.last_row = last_row;

    return LXLSX_NO_ERROR;
}

/*
 * Set the columns to repeat at the left hand side of each printed page.
 */
lxlsx_error
lxlsx_worksheet_repeat_columns(lxlsx_worksheet *self, lxlsx_col_t first_col,
                         lxlsx_col_t last_col)
{
    lxlsx_col_t tmp_col;
    lxlsx_error err;

    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    err = _check_dimensions(self, last_col, 0, LXLSX_IGNORE, LXLSX_IGNORE);
    if (err)
        return err;

    self->repeat_cols.in_use = LXLSX_TRUE;
    self->repeat_cols.first_col = first_col;
    self->repeat_cols.last_col = last_col;

    return LXLSX_NO_ERROR;
}

/*
 * Set the print area in the current worksheet.
 */
lxlsx_error
lxlsx_worksheet_print_area(lxlsx_worksheet *self, lxlsx_row_t first_row,
                     lxlsx_col_t first_col, lxlsx_row_t last_row,
                     lxlsx_col_t last_col)
{
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_error err;

    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }

    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    err = _check_dimensions(self, last_row, last_col, LXLSX_IGNORE, LXLSX_IGNORE);
    if (err)
        return err;

    /* Ignore max area since it is the same as no print area in Excel. */
    if (first_row == 0 && first_col == 0 && last_row == LXLSX_ROW_MAX - 1
        && last_col == LXLSX_COL_MAX - 1) {
        return LXLSX_NO_ERROR;
    }

    self->print_area.in_use = LXLSX_TRUE;
    self->print_area.first_row = first_row;
    self->print_area.last_row = last_row;
    self->print_area.first_col = first_col;
    self->print_area.last_col = last_col;

    return LXLSX_NO_ERROR;
}

/* Store the vertical and horizontal number of pages that will define the
 * maximum area printed.
 */
void
lxlsx_worksheet_fit_to_pages(lxlsx_worksheet *self, uint16_t width, uint16_t height)
{
    self->fit_page = 1;
    self->fit_width = width;
    self->fit_height = height;
    self->page_setup_changed = 1;
}

/*
 * Set the start page number.
 */
void
lxlsx_worksheet_set_start_page(lxlsx_worksheet *self, uint16_t start_page)
{
    self->page_start = start_page;
}

/*
 * Set the scale factor for the printed page.
 */
void
lxlsx_worksheet_set_print_scale(lxlsx_worksheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400)
        return;

    /* Turn off "fit to page" option. */
    self->fit_page = LXLSX_FALSE;

    self->print_scale = scale;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Set the print in black and white option.
 */
void
lxlsx_worksheet_print_black_and_white(lxlsx_worksheet *self)
{
    self->black_white = LXLSX_TRUE;
    self->page_setup_changed = LXLSX_TRUE;
}

/*
 * Store the horizontal page breaks on a worksheet.
 */
lxlsx_error
lxlsx_worksheet_set_h_pagebreaks(lxlsx_worksheet *self, lxlsx_row_t hbreaks[])
{
    uint16_t count = 0;

    if (hbreaks == NULL)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    while (hbreaks[count])
        count++;

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXLSX_BREAKS_MAX)
        count = LXLSX_BREAKS_MAX;

    self->hbreaks = calloc(count, sizeof(lxlsx_row_t));
    RETURN_ON_MEM_ERROR(self->hbreaks, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
    memcpy(self->hbreaks, hbreaks, count * sizeof(lxlsx_row_t));
    self->hbreaks_count = count;

    return LXLSX_NO_ERROR;
}

/*
 * Store the vertical page breaks on a worksheet.
 */
lxlsx_error
lxlsx_worksheet_set_v_pagebreaks(lxlsx_worksheet *self, lxlsx_col_t vbreaks[])
{
    uint16_t count = 0;

    if (vbreaks == NULL)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    while (vbreaks[count])
        count++;

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXLSX_BREAKS_MAX)
        count = LXLSX_BREAKS_MAX;

    self->vbreaks = calloc(count, sizeof(lxlsx_col_t));
    RETURN_ON_MEM_ERROR(self->vbreaks, LXLSX_ERROR_MEMORY_MALLOC_FAILED);
    memcpy(self->vbreaks, vbreaks, count * sizeof(lxlsx_col_t));
    self->vbreaks_count = count;

    return LXLSX_NO_ERROR;
}

/*
 * Set the worksheet zoom factor.
 */
void
lxlsx_worksheet_set_zoom(lxlsx_worksheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400) {
        LXLSX_WARN("lxlsx_worksheet_set_zoom(): "
                 "Zoom factor scale outside range: 10 <= zoom <= 400.");
        return;
    }

    self->zoom = scale;
}

/*
 * Hide cell zero values.
 */
void
lxlsx_worksheet_hide_zero(lxlsx_worksheet *self)
{
    self->show_zeros = LXLSX_FALSE;
}

/*
 * Display the worksheet right to left for some eastern versions of Excel.
 */
void
lxlsx_worksheet_right_to_left(lxlsx_worksheet *self)
{
    self->right_to_left = LXLSX_TRUE;
}

/*
 * Set the color of the worksheet tab.
 */
void
lxlsx_worksheet_set_tab_color(lxlsx_worksheet *self, lxlsx_color_t color)
{
    self->tab_color = color;
}

/*
 * Set the worksheet protection flags to prevent modification of worksheet
 * objects.
 */
void
lxlsx_worksheet_protect(lxlsx_worksheet *self, const char *password,
                  lxlsx_protection *options)
{
    struct lxlsx_protection_obj *protect = &self->protection;

    /* Copy any user parameters to the internal structure. */
    if (options) {
        protect->no_select_locked_cells = options->no_select_locked_cells;
        protect->no_select_unlocked_cells = options->no_select_unlocked_cells;
        protect->lxlsx_format_cells = options->lxlsx_format_cells;
        protect->lxlsx_format_columns = options->lxlsx_format_columns;
        protect->lxlsx_format_rows = options->lxlsx_format_rows;
        protect->insert_columns = options->insert_columns;
        protect->insert_rows = options->insert_rows;
        protect->insert_hyperlinks = options->insert_hyperlinks;
        protect->delete_columns = options->delete_columns;
        protect->delete_rows = options->delete_rows;
        protect->sort = options->sort;
        protect->autofilter = options->autofilter;
        protect->pivot_tables = options->pivot_tables;
        protect->scenarios = options->scenarios;
        protect->objects = options->objects;
    }

    if (password) {
        uint16_t hash = lxlsx_hash_password(password);
        lxlsx_snprintf(protect->hash, 5, "%X", hash);
    }

    protect->no_sheet = LXLSX_FALSE;
    protect->no_content = LXLSX_TRUE;
    protect->is_configured = LXLSX_TRUE;
}

/*
 * Set the worksheet properties for outlines and grouping.
 */
void
lxlsx_worksheet_outline_settings(lxlsx_worksheet *self,
                           uint8_t visible,
                           uint8_t symbols_below,
                           uint8_t symbols_right, uint8_t auto_style)
{
    self->outline_on = visible;
    self->outline_below = symbols_below;
    self->outline_right = symbols_right;
    self->outline_style = auto_style;

    self->outline_changed = LXLSX_TRUE;
}

/*
 * Set the default row properties
 */
void
lxlsx_worksheet_set_default_row(lxlsx_worksheet *self, double height,
                          uint8_t hide_unused_rows)
{
    if (height < 0)
        height = self->default_row_height;

    if (height != self->default_row_height) {
        self->default_row_height = height;
        self->row_size_changed = LXLSX_TRUE;
    }

    if (hide_unused_rows)
        self->default_row_zeroed = LXLSX_TRUE;

    self->default_row_set = LXLSX_TRUE;
}

/*
 * Insert an image with options into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_image_opt(lxlsx_worksheet *self,
                           lxlsx_row_t row_num, lxlsx_col_t col_num,
                           const char *filename,
                           lxlsx_image_options *user_options)
{
    FILE *image_stream;
    const char *description;
    lxlsx_object_properties *object_props;

    if (!filename) {
        LXLSX_WARN("lxlsx_worksheet_insert_image()/_opt(): "
                 "filename must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(filename, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_insert_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Use the filename as the default description, like Excel. */
    description = lxlsx_basename(filename);
    if (!description) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_insert_image()/_opt(): "
                         "couldn't get basename for file: %s.", filename);
        fclose(image_stream);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (user_options) {
        object_props->x_offset = user_options->x_offset;
        object_props->y_offset = user_options->y_offset;
        object_props->x_scale = user_options->x_scale;
        object_props->y_scale = user_options->y_scale;
        object_props->url = lxlsx_strdup(user_options->url);
        object_props->tip = lxlsx_strdup(user_options->tip);
        object_props->object_position = user_options->object_position;
        object_props->decorative = user_options->decorative;

        if (user_options->description)
            description = user_options->description;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup(filename);
    object_props->description = lxlsx_strdup(description);
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);
        fclose(image_stream);
        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_image(lxlsx_worksheet *self,
                       lxlsx_row_t row_num, lxlsx_col_t col_num,
                       const char *filename)
{
    return lxlsx_worksheet_insert_image_opt(self, row_num, col_num, filename, NULL);
}

/*
 * Insert an image buffer, with options, into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_image_buffer_opt(lxlsx_worksheet *self,
                                  lxlsx_row_t row_num,
                                  lxlsx_col_t col_num,
                                  const unsigned char *image_buffer,
                                  size_t image_size,
                                  lxlsx_image_options *user_options)
{
    FILE *image_stream;
    lxlsx_object_properties *object_props;

    if (!image_size) {
        LXLSX_WARN("lxlsx_worksheet_insert_image_buffer()/_opt(): "
                 "size must be non-zero.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxlsx_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXLSX_TRUE;
    }

    if (user_options) {
        object_props->x_offset = user_options->x_offset;
        object_props->y_offset = user_options->y_offset;
        object_props->x_scale = user_options->x_scale;
        object_props->y_scale = user_options->y_scale;
        object_props->url = lxlsx_strdup(user_options->url);
        object_props->tip = lxlsx_strdup(user_options->tip);
        object_props->object_position = user_options->object_position;
        object_props->description = lxlsx_strdup(user_options->description);
        object_props->decorative = user_options->decorative;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);
        fclose(image_stream);
        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image buffer into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_image_buffer(lxlsx_worksheet *self,
                              lxlsx_row_t row_num,
                              lxlsx_col_t col_num,
                              const unsigned char *image_buffer,
                              size_t image_size)
{
    return lxlsx_worksheet_insert_image_buffer_opt(self, row_num, col_num,
                                             image_buffer, image_size, NULL);
}

/*
 * Embed an image with options into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_embed_image_opt(lxlsx_worksheet *self,
                          lxlsx_row_t row_num, lxlsx_col_t col_num,
                          const char *filename,
                          lxlsx_image_options *user_options)
{
    FILE *image_stream;
    lxlsx_object_properties *object_props;
    lxlsx_error err;

    if (!filename) {
        LXLSX_WARN("lxlsx_worksheet_embed_image()/_opt(): "
                 "filename must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(filename, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_embed_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check and store the cell dimensions. */
    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err) {
        fclose(image_stream);
        return err;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* We only copy/use a limited number of options for embedded images. */
    if (user_options) {
        if (user_options->cell_format)
            object_props->format = user_options->cell_format;

        /* The url for embedded images is written as a cell url. */
        if (user_options->url) {
            if (!user_options->cell_format)
                object_props->format = self->default_url_format;

            self->storing_embedded_image = LXLSX_TRUE;
            err = lxlsx_worksheet_write_url(self,
                                      row_num,
                                      col_num,
                                      user_options->url,
                                      object_props->format);
            if (err) {
                _free_object_properties(object_props);
                fclose(image_stream);
                return err;
            }

            self->storing_embedded_image = LXLSX_FALSE;
        }

        object_props->decorative = user_options->decorative;
        if (user_options->description)
            object_props->description = lxlsx_strdup(user_options->description);
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup(filename);
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->embedded_image_props, object_props,
                           list_pointers);
        fclose(image_stream);

        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Embed an image into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_embed_image(lxlsx_worksheet *self,
                      lxlsx_row_t row_num, lxlsx_col_t col_num,
                      const char *filename)
{
    return lxlsx_worksheet_embed_image_opt(self, row_num, col_num, filename, NULL);
}

/*
 * Embed an image buffer, with options, into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_embed_image_buffer_opt(lxlsx_worksheet *self,
                                 lxlsx_row_t row_num,
                                 lxlsx_col_t col_num,
                                 const unsigned char *image_buffer,
                                 size_t image_size,
                                 lxlsx_image_options *user_options)
{
    FILE *image_stream;
    lxlsx_object_properties *object_props;
    lxlsx_error err;

    if (!image_size) {
        LXLSX_WARN("lxlsx_worksheet_embed_image_buffer()/_opt(): "
                 "size must be non-zero.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. For embedded images we really only
     * need the image type. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxlsx_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Check and store the cell dimensions. */
    err = _check_dimensions(self, row_num, col_num, LXLSX_FALSE, LXLSX_FALSE);
    if (err)
        return err;

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXLSX_TRUE;
    }

    /* We only copy/use a limited number of options for embedded images. */
    if (user_options) {
        if (user_options->cell_format)
            object_props->format = user_options->cell_format;

        /* The url for embedded images is written as a cell url. */
        if (user_options->url) {
            if (!user_options->cell_format)
                object_props->format = self->default_url_format;

            self->storing_embedded_image = LXLSX_TRUE;
            err = lxlsx_worksheet_write_url(self,
                                      row_num,
                                      col_num,
                                      user_options->url,
                                      object_props->format);
            if (err) {
                _free_object_properties(object_props);
                fclose(image_stream);
                return err;
            }

            self->storing_embedded_image = LXLSX_FALSE;
        }

        object_props->decorative = user_options->decorative;
        if (user_options->description)
            object_props->description = lxlsx_strdup(user_options->description);
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->embedded_image_props, object_props,
                           list_pointers);
        fclose(image_stream);

        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image buffer into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_embed_image_buffer(lxlsx_worksheet *self,
                             lxlsx_row_t row_num,
                             lxlsx_col_t col_num,
                             const unsigned char *image_buffer,
                             size_t image_size)
{
    return lxlsx_worksheet_embed_image_buffer_opt(self, row_num, col_num,
                                            image_buffer, image_size, NULL);
}

/*
 * Set an image as a worksheet background.
 */
lxlsx_error
lxlsx_worksheet_set_background(lxlsx_worksheet *self, const char *filename)
{
    FILE *image_stream;
    lxlsx_object_properties *object_props;

    if (!filename) {
        LXLSX_WARN("lxlsx_worksheet_set_background(): "
                 "filename must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxlsx_fopen(filename, "rb");
    if (!image_stream) {
        LXLSX_WARN_FORMAT1("lxlsx_worksheet_set_background(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup(filename);
    object_props->stream = image_stream;
    object_props->is_background = LXLSX_TRUE;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        _free_object_properties(self->background_image);
        self->background_image = object_props;
        self->has_background_image = LXLSX_TRUE;
        fclose(image_stream);
        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Set an image buffer as a worksheet background.
 */
lxlsx_error
lxlsx_worksheet_set_background_buffer(lxlsx_worksheet *self,
                                const unsigned char *image_buffer,
                                size_t image_size)
{
    FILE *image_stream;
    lxlsx_object_properties *object_props;

    if (!image_size) {
        LXLSX_WARN("lxlsx_worksheet_set_background(): " "size must be non-zero.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxlsx_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXLSX_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXLSX_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxlsx_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXLSX_TRUE;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxlsx_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->is_background = LXLSX_TRUE;

    if (_get_image_properties(object_props) == LXLSX_NO_ERROR) {
        _free_object_properties(self->background_image);
        self->background_image = object_props;
        self->has_background_image = LXLSX_TRUE;
        fclose(image_stream);
        return LXLSX_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXLSX_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an chart into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_chart_opt(lxlsx_worksheet *self,
                           lxlsx_row_t row_num, lxlsx_col_t col_num,
                           lxlsx_chart *chart, lxlsx_chart_options *user_options)
{
    lxlsx_object_properties *object_props;
    lxlsx_chart_series *series;

    if (!chart) {
        LXLSX_WARN("lxlsx_worksheet_insert_chart()/_opt(): chart must be non-NULL.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the chart isn't being used more than once. */
    if (chart->in_use) {
        LXLSX_WARN("lxlsx_worksheet_insert_chart()/_opt(): the same chart object "
                 "cannot be inserted in a worksheet more than once.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a data series. */
    if (STAILQ_EMPTY(chart->series_list)) {
        LXLSX_WARN
            ("lxlsx_worksheet_insert_chart()/_opt(): chart must have a series.");

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a 'values' series. */
    STAILQ_FOREACH(series, chart->series_list, list_pointers) {
        if (!series->values->formula && !series->values->sheetname) {
            LXLSX_WARN("lxlsx_worksheet_insert_chart()/_opt(): chart must have a "
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
        object_props->object_position = user_options->object_position;
        object_props->description = lxlsx_strdup(user_options->description);
        object_props->decorative = user_options->decorative;
    }

    /* Copy other options or set defaults. */
    object_props->row = row_num;
    object_props->col = col_num;

    object_props->width = 480;
    object_props->height = 288;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    /* Store chart references so they can be ordered in the workbook. */
    object_props->chart = chart;

    STAILQ_INSERT_TAIL(self->lxlsx_chart_data, object_props, list_pointers);

    chart->in_use = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Insert an image into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_chart(lxlsx_worksheet *self,
                       lxlsx_row_t row_num, lxlsx_col_t col_num, lxlsx_chart *chart)
{
    return lxlsx_worksheet_insert_chart_opt(self, row_num, col_num, chart, NULL);
}

/*
 * Add a data validation to a worksheet, for a range. Ironically this requires
 * a lot of validation of the user input.
 */
lxlsx_error
lxlsx_worksheet_data_validation_range(lxlsx_worksheet *self, lxlsx_row_t first_row,
                                lxlsx_col_t first_col,
                                lxlsx_row_t last_row,
                                lxlsx_col_t last_col,
                                lxlsx_data_validation *validation)
{
    lxlsx_data_val_obj *copy;
    uint8_t is_between = LXLSX_FALSE;
    uint8_t is_formula = LXLSX_FALSE;
    uint8_t has_criteria = LXLSX_TRUE;
    lxlsx_error err;
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    size_t length;

    /* No action is required for validation type 'any' unless there are
     * input messages to display.*/
    if (validation->validate == LXLSX_VALIDATION_TYPE_ANY
        && !(validation->input_title || validation->input_message)) {

        return LXLSX_NO_ERROR;
    }

    /* Check for formula types. */
    switch (validation->validate) {
        case LXLSX_VALIDATION_TYPE_INTEGER_FORMULA:
        case LXLSX_VALIDATION_TYPE_DECIMAL_FORMULA:
        case LXLSX_VALIDATION_TYPE_LIST_FORMULA:
        case LXLSX_VALIDATION_TYPE_LENGTH_FORMULA:
        case LXLSX_VALIDATION_TYPE_DATE_FORMULA:
        case LXLSX_VALIDATION_TYPE_TIME_FORMULA:
        case LXLSX_VALIDATION_TYPE_CUSTOM_FORMULA:
            is_formula = LXLSX_TRUE;
            break;
    }

    /* Check for types without a criteria. */
    switch (validation->validate) {
        case LXLSX_VALIDATION_TYPE_LIST:
        case LXLSX_VALIDATION_TYPE_LIST_FORMULA:
        case LXLSX_VALIDATION_TYPE_ANY:
        case LXLSX_VALIDATION_TYPE_CUSTOM_FORMULA:
            has_criteria = LXLSX_FALSE;
            break;
    }

    /* Check that a validation parameter has been specified
     * except for 'list', 'any' and 'custom'. */
    if (has_criteria && validation->criteria == LXLSX_VALIDATION_CRITERIA_NONE) {

        LXLSX_WARN_FORMAT("lxlsx_worksheet_data_validation_cell()/_range(): "
                        "criteria parameter must be specified.");
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Check for "between" criteria so we can do additional checks. */
    if (has_criteria
        && (validation->criteria == LXLSX_VALIDATION_CRITERIA_BETWEEN
            || validation->criteria == LXLSX_VALIDATION_CRITERIA_NOT_BETWEEN)) {

        is_between = LXLSX_TRUE;
    }

    /* Check that formula values are non NULL. */
    if (is_formula) {
        if (is_between) {
            if (!validation->minimum_formula) {
                LXLSX_WARN_FORMAT("lxlsx_worksheet_data_validation_cell()/_range(): "
                                "minimum_formula parameter cannot be NULL.");
                return LXLSX_ERROR_PARAMETER_VALIDATION;
            }
            if (!validation->maximum_formula) {
                LXLSX_WARN_FORMAT("lxlsx_worksheet_data_validation_cell()/_range(): "
                                "maximum_formula parameter cannot be NULL.");
                return LXLSX_ERROR_PARAMETER_VALIDATION;
            }
        }
        else {
            if (!validation->value_formula) {
                LXLSX_WARN_FORMAT("lxlsx_worksheet_data_validation_cell()/_range(): "
                                "formula parameter cannot be NULL.");
                return LXLSX_ERROR_PARAMETER_VALIDATION;
            }
        }
    }

    /* Check Excel limitations on input strings. */
    if (validation->input_title) {
        length = lxlsx_utf8_strlen(validation->input_title);
        if (length > LXLSX_VALIDATION_MAX_TITLE_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_data_validation_cell()/_range(): "
                             "input_title length > Excel limit of %d.",
                             LXLSX_VALIDATION_MAX_TITLE_LENGTH);
            return LXLSX_ERROR_32_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->error_title) {
        length = lxlsx_utf8_strlen(validation->error_title);
        if (length > LXLSX_VALIDATION_MAX_TITLE_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_data_validation_cell()/_range(): "
                             "error_title length > Excel limit of %d.",
                             LXLSX_VALIDATION_MAX_TITLE_LENGTH);
            return LXLSX_ERROR_32_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->input_message) {
        length = lxlsx_utf8_strlen(validation->input_message);
        if (length > LXLSX_VALIDATION_MAX_STRING_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_data_validation_cell()/_range(): "
                             "input_message length > Excel limit of %d.",
                             LXLSX_VALIDATION_MAX_STRING_LENGTH);
            return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->error_message) {
        length = lxlsx_utf8_strlen(validation->error_message);
        if (length > LXLSX_VALIDATION_MAX_STRING_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_data_validation_cell()/_range(): "
                             "error_message length > Excel limit of %d.",
                             LXLSX_VALIDATION_MAX_STRING_LENGTH);
            return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->validate == LXLSX_VALIDATION_TYPE_LIST) {
        length = _validation_list_length(validation->value_list);

        if (length == 0) {
            LXLSX_WARN_FORMAT("lxlsx_worksheet_data_validation_cell()/_range(): "
                            "list parameters cannot be zero.");
            return LXLSX_ERROR_PARAMETER_VALIDATION;
        }

        if (length > LXLSX_VALIDATION_MAX_STRING_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_data_validation_cell()/_range(): "
                             "list length with commas > Excel limit of %d.",
                             LXLSX_VALIDATION_MAX_STRING_LENGTH);
            return LXLSX_ERROR_255_STRING_LENGTH_EXCEEDED;
        }
    }

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that dimensions are valid but don't store them. */
    err = _check_dimensions(self, last_row, last_col, LXLSX_TRUE, LXLSX_TRUE);
    if (err)
        return err;

    /* Create a copy of the parameters from the user data validation. */
    copy = calloc(1, sizeof(lxlsx_data_val_obj));
    GOTO_LABEL_ON_MEM_ERROR(copy, mem_error);

    /* Create the data validation range. */
    if (first_row == last_row && first_col == last_col)
        lxlsx_rowcol_to_cell(copy->sqref, first_row, first_col);
    else
        lxlsx_rowcol_to_range(copy->sqref, first_row, first_col, last_row,
                            last_col);

    /* Copy the parameters from the user data validation. */
    copy->validate = validation->validate;
    copy->value_number = validation->value_number;
    copy->error_type = validation->error_type;
    copy->dropdown = validation->dropdown;

    if (has_criteria)
        copy->criteria = validation->criteria;

    if (is_between) {
        copy->value_number = validation->minimum_number;
        copy->maximum_number = validation->maximum_number;
    }

    /* Copy the input/error titles and messages. */
    if (validation->input_title) {
        copy->input_title = lxlsx_strdup_formula(validation->input_title);
        GOTO_LABEL_ON_MEM_ERROR(copy->input_title, mem_error);
    }

    if (validation->input_message) {
        copy->input_message = lxlsx_strdup_formula(validation->input_message);
        GOTO_LABEL_ON_MEM_ERROR(copy->input_message, mem_error);
    }

    if (validation->error_title) {
        copy->error_title = lxlsx_strdup_formula(validation->error_title);
        GOTO_LABEL_ON_MEM_ERROR(copy->error_title, mem_error);
    }

    if (validation->error_message) {
        copy->error_message = lxlsx_strdup_formula(validation->error_message);
        GOTO_LABEL_ON_MEM_ERROR(copy->error_message, mem_error);
    }

    /* Copy the formula strings. */
    if (is_formula) {
        if (is_between) {
            copy->value_formula =
                lxlsx_strdup_formula(validation->minimum_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
            copy->maximum_formula =
                lxlsx_strdup_formula(validation->maximum_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->maximum_formula, mem_error);
        }
        else {
            copy->value_formula =
                lxlsx_strdup_formula(validation->value_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
        }
    }

    /* Copy the validation list as a csv string. */
    if (validation->validate == LXLSX_VALIDATION_TYPE_LIST) {
        copy->value_formula = _validation_list_to_csv(validation->value_list);
        GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
    }

    if (validation->validate == LXLSX_VALIDATION_TYPE_DATE
        || validation->validate == LXLSX_VALIDATION_TYPE_TIME) {
        if (is_between) {
            copy->value_number =
                lxlsx_datetime_to_excel_date_with_epoch
                (&validation->minimum_datetime, self->use_1904_epoch);
            copy->maximum_number =
                lxlsx_datetime_to_excel_date_with_epoch
                (&validation->maximum_datetime, self->use_1904_epoch);
        }
        else {
            copy->value_number =
                lxlsx_datetime_to_excel_date_with_epoch
                (&validation->value_datetime, self->use_1904_epoch);
        }
    }

    /* These options are on by default so we can't take plain booleans. */
    copy->ignore_blank = validation->ignore_blank ^ 1;
    copy->show_input = validation->show_input ^ 1;
    copy->show_error = validation->show_error ^ 1;

    STAILQ_INSERT_TAIL(self->data_validations, copy, list_pointers);

    self->num_validations++;

    return LXLSX_NO_ERROR;

mem_error:
    _free_data_validation(copy);
    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Add a data validation to a worksheet, for a cell.
 */
lxlsx_error
lxlsx_worksheet_data_validation_cell(lxlsx_worksheet *self, lxlsx_row_t row,
                               lxlsx_col_t col, lxlsx_data_validation *validation)
{
    return lxlsx_worksheet_data_validation_range(self, row, col,
                                           row, col, validation);
}

/*
 * Add a conditional format to a worksheet, for a range.
 */
lxlsx_error
lxlsx_worksheet_conditional_format_range(lxlsx_worksheet *self, lxlsx_row_t first_row,
                                   lxlsx_col_t first_col,
                                   lxlsx_row_t last_row,
                                   lxlsx_col_t last_col,
                                   lxlsx_conditional_format *user_options)
{
    lxlsx_cond_format_obj *cond_format;
    lxlsx_row_t tmp_row;
    lxlsx_col_t tmp_col;
    lxlsx_error err = LXLSX_NO_ERROR;
    char *type_strings[] = {
        "none",
        "cellIs",
        "containsText",
        "timePeriod",
        "aboveAverage",
        "duplicateValues",
        "uniqueValues",
        "top10",
        "top10",
        "containsBlanks",
        "notContainsBlanks",
        "containsErrors",
        "notContainsErrors",
        "expression",
        "colorScale",
        "colorScale",
        "dataBar",
        "iconSet",
    };

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that dimensions are valid but don't store them. */
    err = _check_dimensions(self, last_row, last_col, LXLSX_TRUE, LXLSX_TRUE);
    if (err)
        return err;

    /* Check the validation type is in correct enum range. */
    if (user_options->type <= LXLSX_CONDITIONAL_TYPE_NONE ||
        user_options->type >= LXLSX_CONDITIONAL_TYPE_LAST) {

        LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                         "invalid type value (%d).", user_options->type);

        return LXLSX_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a copy of the parameters from the user data validation. */
    cond_format = calloc(1, sizeof(lxlsx_cond_format_obj));
    GOTO_LABEL_ON_MEM_ERROR(cond_format, error);

    /* Create the data validation range. */
    if (first_row == last_row && first_col == last_col)
        lxlsx_rowcol_to_cell(cond_format->sqref, first_row, first_col);
    else
        lxlsx_rowcol_to_range(cond_format->sqref, first_row, first_col,
                            last_row, last_col);

    /* Store the first cell string for text and date rules. */
    lxlsx_rowcol_to_cell(cond_format->first_cell, first_row, first_col);

    /* Overwrite the sqref range with a user supplied set of ranges. */
    if (user_options->multi_range) {

        if (strlen(user_options->multi_range) >= LXLSX_MAX_ATTRIBUTE_LENGTH) {
            LXLSX_WARN_FORMAT1("lxlsx_worksheet_conditional_format_cell()/_range(): "
                             "multi_range >= limit = %d.",
                             LXLSX_MAX_ATTRIBUTE_LENGTH);
            err = LXLSX_ERROR_PARAMETER_VALIDATION;
            goto error;
        }

        LXLSX_ATTRIBUTE_COPY(cond_format->sqref, user_options->multi_range);
    }

    /* Get the conditional format dxf format index. */
    if (user_options->format)
        cond_format->dxf_index =
            lxlsx_format_get_dxf_index(user_options->format);
    else
        cond_format->dxf_index = LXLSX_PROPERTY_UNSET;

    /* Set some common option for all validation types. */
    cond_format->type = user_options->type;
    cond_format->criteria = user_options->criteria;
    cond_format->stop_if_true = user_options->stop_if_true;
    cond_format->type_string = lxlsx_strdup(type_strings[cond_format->type]);

    /* Check that the criteria matches the conditional type. */
    err = _validate_conditional_criteria(cond_format);
    if (err)
        goto error;

    /* Validate the user input for various types of rules. */
    if (user_options->type == LXLSX_CONDITIONAL_TYPE_CELL
        || cond_format->type == LXLSX_CONDITIONAL_TYPE_DUPLICATE
        || cond_format->type == LXLSX_CONDITIONAL_TYPE_UNIQUE) {

        err = _validate_conditional_cell(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXLSX_CONDITIONAL_TYPE_TEXT) {

        err = _validate_conditional_text(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXLSX_CONDITIONAL_TYPE_TIME_PERIOD) {

        err = _validate_conditional_time_period(user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXLSX_CONDITIONAL_TYPE_AVERAGE) {

        err = _validate_conditional_average(user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXLSX_CONDITIONAL_TYPE_BOTTOM) {

        err = _validate_conditional_top(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXLSX_CONDITIONAL_TYPE_FORMULA) {

        err = _validate_conditional_formula(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_2_COLOR_SCALE
             || cond_format->type == LXLSX_CONDITIONAL_3_COLOR_SCALE) {

        err = _validate_conditional_scale(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_DATA_BAR) {

        err = _validate_conditional_data_bar(self, cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXLSX_CONDITIONAL_TYPE_ICON_SETS) {

        err = _validate_conditional_icons(user_options);
        if (err)
            goto error;

        cond_format->icon_style = user_options->icon_style;
        cond_format->reverse_icons = user_options->reverse_icons;
        cond_format->icons_only = user_options->icons_only;
    }

    /* Set the priority based on the order of adding. */
    cond_format->dxf_priority = ++self->dxf_priority;

    /* Store the conditional format object. */
    err = _store_conditional_format_object(self, cond_format);

    if (err)
        goto error;
    else
        return LXLSX_NO_ERROR;

error:
    _free_cond_format(cond_format);
    return err;
}

/*
 * Add a conditional format to a worksheet, for a cell.
 */
lxlsx_error
lxlsx_worksheet_conditional_format_cell(lxlsx_worksheet *self,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  lxlsx_conditional_format *options)
{
    return lxlsx_worksheet_conditional_format_range(self, row, col,
                                              row, col, options);
}

/*
 * Insert a button object into the worksheet.
 */
lxlsx_error
lxlsx_worksheet_insert_button(lxlsx_worksheet *self, lxlsx_row_t row_num,
                        lxlsx_col_t col_num, lxlsx_button_options *options)
{
    lxlsx_error err;
    lxlsx_vml_obj *button;

    err = _check_dimensions(self, row_num, col_num, LXLSX_TRUE, LXLSX_TRUE);
    if (err)
        return err;

    button = calloc(1, sizeof(lxlsx_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(button, mem_error);

    button->row = row_num;
    button->col = col_num;

    /* Set user and default parameters for the button. */
    err = _get_button_params(button, 1 + self->num_buttons, options);
    if (err)
        goto mem_error;

    /* Calculate the worksheet position of the button. */
    _worksheet_position_vml_object(self, button);

    self->has_vml = LXLSX_TRUE;
    self->has_buttons = LXLSX_TRUE;
    self->num_buttons++;

    STAILQ_INSERT_TAIL(self->button_objs, button, list_pointers);

    return LXLSX_NO_ERROR;

mem_error:
    if (button)
        _free_vml_object(button);

    return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Set the VBA name for the worksheet.
 */
lxlsx_error
lxlsx_worksheet_set_vba_name(lxlsx_worksheet *self, const char *name)
{
    if (!name) {
        LXLSX_WARN("lxlsx_worksheet_set_vba_name(): " "name must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    self->vba_codename = lxlsx_strdup(name);

    return LXLSX_NO_ERROR;
}

/*
 * Set the default author of the cell comments.
 */
void
lxlsx_worksheet_set_comments_author(lxlsx_worksheet *self, const char *author)
{
    self->comment_author = lxlsx_strdup(author);
}

/*
 * Make any comments in the worksheet visible, unless explicitly hidden.
 */
void
lxlsx_worksheet_show_comments(lxlsx_worksheet *self)
{
    self->comment_display_default = LXLSX_COMMENT_DISPLAY_VISIBLE;
}

/*
 * Ignore various Excel errors/warnings in a worksheet for user defined ranges.
 */
lxlsx_error
lxlsx_worksheet_ignore_errors(lxlsx_worksheet *self, uint8_t type, const char *range)
{
    if (!range) {
        LXLSX_WARN("lxlsx_worksheet_ignore_errors(): " "'range' must be specified.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (type <= 0 || type >= LXLSX_IGNORE_LAST_OPTION) {
        LXLSX_WARN("lxlsx_worksheet_ignore_errors(): " "unknown option in 'type'.");
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Set the ranges to be ignored. */
    if (type == LXLSX_IGNORE_NUMBER_STORED_AS_TEXT) {
        free(self->ignore_number_stored_as_text);
        self->ignore_number_stored_as_text = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_EVAL_ERROR) {
        free(self->ignore_eval_error);
        self->ignore_eval_error = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_FORMULA_DIFFERS) {
        free(self->ignore_formula_differs);
        self->ignore_formula_differs = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_FORMULA_RANGE) {
        free(self->ignore_formula_range);
        self->ignore_formula_range = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_FORMULA_UNLOCKED) {
        free(self->ignore_formula_unlocked);
        self->ignore_formula_unlocked = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_EMPTY_CELL_REFERENCE) {
        free(self->ignore_empty_cell_reference);
        self->ignore_empty_cell_reference = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_LIST_DATA_VALIDATION) {
        free(self->ignore_list_data_validation);
        self->ignore_list_data_validation = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_CALCULATED_COLUMN) {
        free(self->ignore_calculated_column);
        self->ignore_calculated_column = lxlsx_strdup(range);
    }
    else if (type == LXLSX_IGNORE_TWO_DIGIT_TEXT_YEAR) {
        free(self->ignore_two_digit_text_year);
        self->ignore_two_digit_text_year = lxlsx_strdup(range);
    }

    self->has_ignore_errors = LXLSX_TRUE;

    return LXLSX_NO_ERROR;
}

/*
 * Write an error cell for versions of Excel that don't support embedded images.
 */
void
lxlsx_worksheet_set_error_cell(lxlsx_worksheet *self,
                         lxlsx_object_properties *object_props, uint32_t ref_id)
{
    lxlsx_row_t row_num = object_props->row;
    lxlsx_col_t col_num = object_props->col;

    lxlsx_cell *cell =
        _new_error_cell(row_num, col_num, ref_id, object_props->format);
    _insert_cell(self, row_num, col_num, cell);

}


/****************************************************************************
 *
 * XLSX worksheet row/cell read support.
 *
 ****************************************************************************/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "xlsx_private.h"
#include "xlsx_util.h"

/* ------------------------------------------------------------------------- */
/* Helpers                                                                   */
/* ------------------------------------------------------------------------- */

/* Excel serial date -> Unix timestamp.
 * 1900 system: serial 1 == 1900-01-01, but Excel treats 1900 as a leap year
 * (Lotus 1-2-3 bug), so we anchor at 1899-12-30 to keep all serials >= 61
 * correct. Serials 0..59 inherit the historical Lotus interpretation.
 * 1904 system: serial 0 == 1904-01-01. */
int64_t lxlsx_reader_excel_serial_to_unix(double serial, int uses_1904)
{
    static const int64_t EPOCH_1900 = -2209161600LL; /* 1899-12-30T00:00:00Z */
    static const int64_t EPOCH_1904 = -2082844800LL; /* 1904-01-01T00:00:00Z */
    int64_t base = uses_1904 ? EPOCH_1904 : EPOCH_1900;
    return base + (int64_t)(serial * 86400.0);
}

/* ------------------------------------------------------------------------- */
/* Cell type inference                                                       */
/* ------------------------------------------------------------------------- */

static void emit_cell(lxlsx_reader_worksheet *ws, lxlsx_cell *out)
{
    const lxlsx_reader_styles *st = ws->wb ? ws->wb->styles : NULL;
    const lxlsx_reader_xf     *xf = NULL;
    const char       *t  = ws->cell_t;

    memset(out, 0, sizeof(*out));
    out->row_num  = (lxlsx_row_t)ws->cell_row;
    out->col_num  = (lxlsx_col_t)ws->cell_col;
    out->style_id = ws->cell_style_id;
    out->style_ref = ws->cell_style_id;
    out->raw.ptr  = ws->cell_value;
    out->raw.len  = ws->cell_value_len;

    if (st) xf = lxlsx_reader_styles_get_xf(st, ws->cell_style_id);

    /* Empty cell (no <v> and no inline string and no formula) */
    if (!ws->cell_has_formula && !ws->cell_has_inline &&
        ws->cell_value_len == 0) {
        out->type = BLANK_CELL;
        return;
    }

    if (ws->cell_has_formula) {
        out->type = FORMULA_CELL;
        out->value.formula.formula.ptr = ws->cell_formula;
        out->value.formula.formula.len = ws->cell_formula_len;
        out->value.formula.cached.ptr  = ws->cell_value;
        out->value.formula.cached.len  = ws->cell_value_len;
        out->value.formula.kind        = ws->cell_formula_kind;
        out->value.formula.ref.ptr     = ws->cell_formula_ref[0]
            ? ws->cell_formula_ref : NULL;
        out->value.formula.ref.len     = ws->cell_formula_ref[0]
            ? strlen(ws->cell_formula_ref) : 0;
        out->value.formula.si          = ws->cell_formula_si;
        out->value.formula.is_dynamic  = ws->cell_formula_is_dynamic;
        return;
    }

    if (t[0] == 0 || strcmp(t, "n") == 0) {
        if (xf && (xf->category == LXLSX_READER_FMT_CATEGORY_DATE ||
                   xf->category == LXLSX_READER_FMT_CATEGORY_TIME ||
                   xf->category == LXLSX_READER_FMT_CATEGORY_DATETIME)) {
            double serial = ws->cell_value ? strtod(ws->cell_value, NULL) : 0.0;
            out->type = DATETIME_CELL;
            out->value.unix_timestamp =
                lxlsx_reader_excel_serial_to_unix(serial,
                                         ws->wb ? ws->wb->uses_1904 : 0);
        } else {
            out->type = NUMBER_CELL;
            out->value.number = ws->cell_value ? strtod(ws->cell_value, NULL) : 0.0;
        }
        return;
    }

    if (strcmp(t, "s") == 0) {
        uint32_t idx = ws->cell_value
            ? (uint32_t)strtoul(ws->cell_value, NULL, 10) : 0;
        const char *s = ws->wb && ws->wb->sst
            ? lxlsx_reader_sst_get(ws->wb->sst, idx) : NULL;
        out->type = STRING_CELL;
        if (s) {
            out->value.string.ptr = s;
            out->value.string.len = strlen(s);
        } else {
            out->value.string.ptr = "";
            out->value.string.len = 0;
        }
        return;
    }

    if (strcmp(t, "inlineStr") == 0) {
        out->type = INLINE_STRING_CELL;
        out->value.string.ptr = ws->cell_inline;
        out->value.string.len = ws->cell_inline_len;
        return;
    }

    if (strcmp(t, "str") == 0) {
        out->type = STRING_CELL;
        out->value.string.ptr = ws->cell_value ? ws->cell_value : "";
        out->value.string.len = ws->cell_value_len;
        return;
    }

    if (strcmp(t, "b") == 0) {
        out->type = BOOLEAN_CELL;
        out->value.boolean = (ws->cell_value && ws->cell_value[0] == '1') ? 1 : 0;
        return;
    }

    if (strcmp(t, "e") == 0) {
        size_t n = ws->cell_value_len;
        if (n >= sizeof(out->value.error_code))
            n = sizeof(out->value.error_code) - 1;
        out->type = ERROR_CELL;
        if (ws->cell_value) memcpy(out->value.error_code, ws->cell_value, n);
        out->value.error_code[n] = 0;
        return;
    }

    /* Unknown 't': fall back to string. */
    out->type = STRING_CELL;
    out->value.string.ptr = ws->cell_value ? ws->cell_value : "";
    out->value.string.len = ws->cell_value_len;
}

/* ------------------------------------------------------------------------- */
/* Cell reset                                                                */
/* ------------------------------------------------------------------------- */

static void reset_cell(lxlsx_reader_worksheet *ws)
{
    ws->cell_t[0] = 0;
    ws->cell_ref[0] = 0;
    ws->cell_style_id = 0;
    ws->cell_value_len = 0;
    if (ws->cell_value) ws->cell_value[0] = 0;
    ws->cell_formula_len = 0;
    if (ws->cell_formula) ws->cell_formula[0] = 0;
    ws->cell_inline_len = 0;
    if (ws->cell_inline) ws->cell_inline[0] = 0;
    ws->cell_has_formula = 0;
    ws->cell_has_inline = 0;
    ws->cell_row = 0;
    ws->cell_col = 0;
    ws->cell_formula_kind       = LXLSX_FORMULA_NORMAL;
    ws->cell_formula_ref[0]     = 0;
    ws->cell_formula_si         = -1;
    ws->cell_formula_is_dynamic = 0;
    /* Free any rich runs accumulated for the previous inline-string cell. */
    if (ws->inline_runs) {
        size_t i;
        for (i = 0; i < ws->inline_runs_count; i++) {
            free(ws->inline_runs[i].text);
            free(ws->inline_runs[i].font_name);
            free(ws->inline_runs[i].color);
        }
        ws->inline_runs_count = 0;
    }
    ws->inline_in_r = ws->inline_in_rpr = ws->inline_in_run_t = 0;
    free(ws->inline_pending_text);     ws->inline_pending_text = NULL;
    free(ws->inline_pending_font_name); ws->inline_pending_font_name = NULL;
    free(ws->inline_pending_color);     ws->inline_pending_color = NULL;
}

/* ------------------------------------------------------------------------- */
/* SAX dispatch                                                              */
/* ------------------------------------------------------------------------- */

static void deliver_cell(lxlsx_reader_worksheet *ws)
{
    if (ws->pull_mode == LXLSX_READER_WS_PULL_CELL) {
        ws->pending_cell = 1;
        lxlsx_reader_xml_pump_suspend(ws->pump);
        return;
    }
    if (ws->user_cell_cb) {
        lxlsx_cell c;
        emit_cell(ws, &c);
        if (ws->user_cell_cb(&c, ws->user_data) != 0) {
            ws->callback_stop = 1;
            lxlsx_reader_xml_pump_suspend(ws->pump);
        }
    }
}

static void deliver_row_end(lxlsx_reader_worksheet *ws)
{
    if (ws->pull_mode == LXLSX_READER_WS_PULL_CELL) {
        /* No more cells in this row */
        ws->pending_row_end = 1;
        lxlsx_reader_xml_pump_suspend(ws->pump);
        return;
    }
    if (ws->user_row_cb) {
        if (ws->user_row_cb(ws->row_nr, ws->max_col_seen, ws->user_data) != 0) {
            ws->callback_stop = 1;
            lxlsx_reader_xml_pump_suspend(ws->pump);
        }
    }
}

static void fail_parse(lxlsx_reader_worksheet *ws, lxlsx_reader_error err)
{
    if (!ws->parse_error)
        ws->parse_error = err;
    if (ws->pump)
        lxlsx_reader_xml_pump_suspend(ws->pump);
}

static void on_start(void *ud, const char *name, const char **attrs)
{
    lxlsx_reader_worksheet *ws = (lxlsx_reader_worksheet *)ud;

    if (ws->parse_error)
        return;

    if (ws->state == LXLSX_READER_WS_SKIP) {
        ws->skip_depth++;
        return;
    }

    switch (ws->state) {
    case LXLSX_READER_WS_INIT:
        if (lxlsx_reader_xml_name_eq(name, "worksheet")) ws->state = LXLSX_READER_WS_IN_WORKSHEET;
        break;

    case LXLSX_READER_WS_IN_WORKSHEET:
        if (lxlsx_reader_xml_name_eq(name, "sheetData")) ws->state = LXLSX_READER_WS_IN_SHEETDATA;
        break;

    case LXLSX_READER_WS_IN_SHEETDATA:
        if (lxlsx_reader_xml_name_eq(name, "row")) {
            const char *r_attr = lxlsx_reader_xml_attr(attrs, "r");
            const char *hidden = lxlsx_reader_xml_attr(attrs, "hidden");
            ws->row_nr = r_attr ? (size_t)strtoul(r_attr, NULL, 10) : ws->row_nr + 1;
            ws->row_hidden = (hidden && (strcmp(hidden, "1") == 0 ||
                                         strcmp(hidden, "true") == 0));
            ws->row_in_progress = 1;
            ws->state = LXLSX_READER_WS_IN_ROW;

            if (ws->row_hidden && (ws->flags & LXLSX_READER_SKIP_HIDDEN_ROWS)) {
                /* skip the entire row content */
                free(ws->skip_tag);
                ws->skip_tag = strdup("row");
                ws->state_before_skip = LXLSX_READER_WS_IN_SHEETDATA;
                ws->state = LXLSX_READER_WS_SKIP;
                ws->skip_depth = 1;
                ws->row_in_progress = 0;
                return;
            }

            if (ws->skip_rows_remaining > 0) {
                ws->skip_rows_remaining--;
                free(ws->skip_tag);
                ws->skip_tag = strdup("row");
                ws->state_before_skip = LXLSX_READER_WS_IN_SHEETDATA;
                ws->state = LXLSX_READER_WS_SKIP;
                ws->skip_depth = 1;
                ws->row_in_progress = 0;
                return;
            }

            if (ws->pull_mode == LXLSX_READER_WS_PULL_ROW) {
                ws->pending_row_start = 1;
                lxlsx_reader_xml_pump_suspend(ws->pump);
            }
        }
        break;

    case LXLSX_READER_WS_IN_ROW:
        if (lxlsx_reader_xml_name_eq(name, "c")) {
            const char *r_attr = lxlsx_reader_xml_attr(attrs, "r");
            const char *t_attr = lxlsx_reader_xml_attr(attrs, "t");
            const char *s_attr = lxlsx_reader_xml_attr(attrs, "s");

            reset_cell(ws);
            if (lxlsx_reader_copy_attr(ws->cell_ref, sizeof(ws->cell_ref), r_attr) != 0) {
                fail_parse(ws, LXLSX_READER_ERROR_INVALID_CELL_REF);
                return;
            }
            if (lxlsx_reader_copy_attr(ws->cell_t, sizeof(ws->cell_t), t_attr) != 0) {
                fail_parse(ws, LXLSX_READER_ERROR_FILE_CORRUPTED);
                return;
            }
            ws->cell_style_id = s_attr ? (uint32_t)strtoul(s_attr, NULL, 10) : 0;
            if (r_attr) lxlsx_reader_parse_a1_ref(r_attr, &ws->cell_row, &ws->cell_col);
            else { ws->cell_row = ws->row_nr; ws->cell_col = 0; }
            if (ws->cell_col > ws->max_col_seen) ws->max_col_seen = ws->cell_col;

            ws->state = LXLSX_READER_WS_IN_CELL;
        }
        break;

    case LXLSX_READER_WS_IN_CELL:
        if (lxlsx_reader_xml_name_eq(name, "v")) {
            ws->state = LXLSX_READER_WS_IN_VALUE;
        } else if (lxlsx_reader_xml_name_eq(name, "f")) {
            const char *t_attr   = lxlsx_reader_xml_attr(attrs, "t");
            const char *ref_attr = lxlsx_reader_xml_attr(attrs, "ref");
            const char *si_attr  = lxlsx_reader_xml_attr(attrs, "si");
            const char *aca_attr = lxlsx_reader_xml_attr(attrs, "aca");
            ws->cell_has_formula = 1;
            if (t_attr) {
                if (strcmp(t_attr, "array") == 0)          ws->cell_formula_kind = LXLSX_FORMULA_ARRAY;
                else if (strcmp(t_attr, "dataTable") == 0) ws->cell_formula_kind = LXLSX_FORMULA_DATATABLE;
                else if (strcmp(t_attr, "shared") == 0)    ws->cell_formula_kind = LXLSX_FORMULA_SHARED;
                else                                        ws->cell_formula_kind = LXLSX_FORMULA_NORMAL;
            }
            if (ref_attr &&
                lxlsx_reader_copy_attr(ws->cell_formula_ref,
                                       sizeof(ws->cell_formula_ref), ref_attr) != 0) {
                fail_parse(ws, LXLSX_READER_ERROR_INVALID_CELL_REF);
                return;
            }
            if (si_attr) ws->cell_formula_si = (int)strtol(si_attr, NULL, 10);
            if (aca_attr && (strcmp(aca_attr, "1") == 0 ||
                             strcmp(aca_attr, "true") == 0))
                ws->cell_formula_is_dynamic = 1;
            ws->state = LXLSX_READER_WS_IN_FORMULA;
        } else if (lxlsx_reader_xml_name_eq(name, "is")) {
            ws->cell_has_inline = 1;
            ws->state = LXLSX_READER_WS_IN_INLINE_STR;
        } else {
            /* Skip unknown sub-elements (extLst, etc.) */
            free(ws->skip_tag);
            ws->skip_tag = strdup(name);
            ws->state_before_skip = LXLSX_READER_WS_IN_CELL;
            ws->state = LXLSX_READER_WS_SKIP;
            ws->skip_depth = 1;
        }
        break;

    case LXLSX_READER_WS_IN_INLINE_STR:
        if (lxlsx_reader_xml_name_eq(name, "t")) {
            ws->state = LXLSX_READER_WS_IN_INLINE_STR_T;
        } else if (lxlsx_reader_xml_name_eq(name, "r")) {
            /* Begin a new rich-text run. */
            ws->inline_in_r = 1;
            free(ws->inline_pending_text);     ws->inline_pending_text = NULL;
            free(ws->inline_pending_font_name); ws->inline_pending_font_name = NULL;
            free(ws->inline_pending_color);     ws->inline_pending_color = NULL;
            ws->inline_pending_font_size = 0;
            ws->inline_pending_bold = ws->inline_pending_italic = 0;
            ws->inline_pending_strike = ws->inline_pending_underline = 0;
            ws->inline_run_text_len = 0;
            if (ws->inline_run_text_buf) ws->inline_run_text_buf[0] = 0;
        } else {
            free(ws->skip_tag);
            ws->skip_tag = strdup(name);
            ws->state_before_skip = LXLSX_READER_WS_IN_INLINE_STR;
            ws->state = LXLSX_READER_WS_SKIP;
            ws->skip_depth = 1;
        }
        break;

    default:
        /* Inline-string rich runs sit on top of the existing FSM. We don't
         * dedicate a new state because the structure is local and small. */
        if (ws->state == LXLSX_READER_WS_IN_INLINE_STR && ws->inline_in_r) {
            /* (handled above) */
        }
        break;
    }

    /* Out-of-band: <r>/<rPr>/<t>/<rPr children> handling for inline rich runs. */
    if (ws->inline_in_r) {
        if (lxlsx_reader_xml_name_eq(name, "rPr")) {
            ws->inline_in_rpr = 1;
        } else if (ws->inline_in_rpr) {
            const char *v;
            if (lxlsx_reader_xml_name_eq(name, "rFont")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "val"))) {
                    free(ws->inline_pending_font_name);
                    ws->inline_pending_font_name = strdup(v);
                    if (!ws->inline_pending_font_name) {
                        fail_parse(ws, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                        return;
                    }
                }
            } else if (lxlsx_reader_xml_name_eq(name, "sz")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "val"))) ws->inline_pending_font_size = strtod(v, NULL);
            } else if (lxlsx_reader_xml_name_eq(name, "b")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                ws->inline_pending_bold = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "i")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                ws->inline_pending_italic = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "strike")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                ws->inline_pending_strike = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "u")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                if (!v || strcmp(v, "single") == 0)             ws->inline_pending_underline = 1;
                else if (strcmp(v, "double") == 0)              ws->inline_pending_underline = 2;
                else if (strcmp(v, "singleAccounting") == 0)    ws->inline_pending_underline = 3;
                else if (strcmp(v, "doubleAccounting") == 0)    ws->inline_pending_underline = 4;
            } else if (lxlsx_reader_xml_name_eq(name, "color")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "rgb"))) {
                    free(ws->inline_pending_color);
                    ws->inline_pending_color = strdup(v);
                    if (!ws->inline_pending_color) {
                        fail_parse(ws, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                        return;
                    }
                }
            }
        } else if (lxlsx_reader_xml_name_eq(name, "t")) {
            ws->inline_in_run_t = 1;
        }
    }
}

static void on_text(void *ud, const char *text, int len)
{
    lxlsx_reader_worksheet *ws = (lxlsx_reader_worksheet *)ud;
    int rc = 0;
    if (ws->parse_error || len <= 0) return;

    switch (ws->state) {
    case LXLSX_READER_WS_IN_VALUE:
        rc = lxlsx_reader_buf_append(&ws->cell_value, &ws->cell_value_len,
                                     &ws->cell_value_cap, text, (size_t)len);
        break;
    case LXLSX_READER_WS_IN_FORMULA:
        rc = lxlsx_reader_buf_append(&ws->cell_formula, &ws->cell_formula_len,
                                     &ws->cell_formula_cap, text, (size_t)len);
        break;
    case LXLSX_READER_WS_IN_INLINE_STR_T:
        rc = lxlsx_reader_buf_append(&ws->cell_inline, &ws->cell_inline_len,
                                     &ws->cell_inline_cap, text, (size_t)len);
        /* Fall-through-style: when this <t> sits inside an <r>, also feed
         * the per-run text accumulator. */
        if (rc == 0 && ws->inline_in_run_t) {
            rc = lxlsx_reader_buf_append(&ws->inline_run_text_buf,
                                         &ws->inline_run_text_len,
                                         &ws->inline_run_text_cap,
                                         text, (size_t)len);
        }
        break;
    default:
        break;
    }
    if (rc != 0)
        fail_parse(ws, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
}

static void on_end(void *ud, const char *name)
{
    lxlsx_reader_worksheet *ws = (lxlsx_reader_worksheet *)ud;

    if (ws->parse_error)
        return;

    if (ws->state == LXLSX_READER_WS_SKIP) {
        ws->skip_depth--;
        if (ws->skip_depth == 0 && ws->skip_tag &&
            lxlsx_reader_xml_name_eq(name, ws->skip_tag)) {
            free(ws->skip_tag);
            ws->skip_tag = NULL;
            ws->state = ws->state_before_skip;
            /* For row-level skips, transition out the same way a real
             * </row> would (back to IN_SHEETDATA). */
        }
        return;
    }

    /* Out-of-band: rich-text run end transitions for inline strings. */
    if (ws->inline_in_run_t && lxlsx_reader_xml_name_eq(name, "t")) {
        ws->inline_in_run_t = 0;
    } else if (ws->inline_in_rpr && lxlsx_reader_xml_name_eq(name, "rPr")) {
        ws->inline_in_rpr = 0;
    } else if (ws->inline_in_r && lxlsx_reader_xml_name_eq(name, "r")) {
        /* Commit run to inline_runs[]. */
        if (ws->inline_runs_count >= ws->inline_runs_cap) {
            size_t nc = ws->inline_runs_cap ? ws->inline_runs_cap * 2 : 4;
            void *nb = realloc(ws->inline_runs, nc * sizeof(*ws->inline_runs));
            if (nb) {
                ws->inline_runs = nb;
                ws->inline_runs_cap = nc;
            } else {
                fail_parse(ws, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                return;
            }
        }
        if (ws->inline_runs_count < ws->inline_runs_cap) {
            ws->inline_runs[ws->inline_runs_count].text =
                ws->inline_run_text_len > 0 ? strdup(ws->inline_run_text_buf) : strdup("");
            if (!ws->inline_runs[ws->inline_runs_count].text) {
                fail_parse(ws, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                return;
            }
            ws->inline_runs[ws->inline_runs_count].font_name = ws->inline_pending_font_name;
            ws->inline_runs[ws->inline_runs_count].color     = ws->inline_pending_color;
            ws->inline_runs[ws->inline_runs_count].font_size = ws->inline_pending_font_size;
            ws->inline_runs[ws->inline_runs_count].bold      = ws->inline_pending_bold;
            ws->inline_runs[ws->inline_runs_count].italic    = ws->inline_pending_italic;
            ws->inline_runs[ws->inline_runs_count].strike    = ws->inline_pending_strike;
            ws->inline_runs[ws->inline_runs_count].underline = ws->inline_pending_underline;
            ws->inline_runs_count++;
            /* The pending struct's owned strings are now owned by inline_runs. */
            ws->inline_pending_font_name = NULL;
            ws->inline_pending_color     = NULL;
        }
        ws->inline_in_r = 0;
    }

    switch (ws->state) {
    case LXLSX_READER_WS_IN_VALUE:
        if (lxlsx_reader_xml_name_eq(name, "v")) ws->state = LXLSX_READER_WS_IN_CELL;
        break;
    case LXLSX_READER_WS_IN_FORMULA:
        if (lxlsx_reader_xml_name_eq(name, "f")) ws->state = LXLSX_READER_WS_IN_CELL;
        break;
    case LXLSX_READER_WS_IN_INLINE_STR_T:
        if (lxlsx_reader_xml_name_eq(name, "t")) ws->state = LXLSX_READER_WS_IN_INLINE_STR;
        break;
    case LXLSX_READER_WS_IN_INLINE_STR:
        if (lxlsx_reader_xml_name_eq(name, "is")) ws->state = LXLSX_READER_WS_IN_CELL;
        break;
    case LXLSX_READER_WS_IN_CELL:
        if (lxlsx_reader_xml_name_eq(name, "c")) {
            ws->state = LXLSX_READER_WS_IN_ROW;
            deliver_cell(ws);
        }
        break;
    case LXLSX_READER_WS_IN_ROW:
        if (lxlsx_reader_xml_name_eq(name, "row")) {
            ws->state = LXLSX_READER_WS_IN_SHEETDATA;
            ws->row_in_progress = 0;
            deliver_row_end(ws);
        }
        break;
    case LXLSX_READER_WS_IN_SHEETDATA:
        if (lxlsx_reader_xml_name_eq(name, "sheetData")) ws->state = LXLSX_READER_WS_IN_WORKSHEET;
        break;
    case LXLSX_READER_WS_IN_WORKSHEET:
        if (lxlsx_reader_xml_name_eq(name, "worksheet")) {
            ws->state = LXLSX_READER_WS_INIT;
            ws->eof   = 1;
        }
        break;
    default:
        break;
    }
}

/* ------------------------------------------------------------------------- */
/* Open / close                                                              */
/* ------------------------------------------------------------------------- */

static lxlsx_reader_error open_internal(lxlsx_reader_workbook *wb, const char *target,
                               uint32_t flags, lxlsx_reader_worksheet **out)
{
    lxlsx_reader_worksheet *ws;
    if (!wb || !target || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;

    /* Validate the entry exists up front (cheap locate, no open) so that a
     * bad path still fails at open time. The data stream and metadata are
     * loaded lazily — see ensure_data_open / lxlsx_reader_worksheet_ensure_meta. */
    if (!lxlsx_reader_zip_entry_exists(wb->zip, target)) {
        return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;
    }

    ws = (lxlsx_reader_worksheet *)calloc(1, sizeof(*ws));
    if (!ws) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    ws->wb          = wb;
    ws->flags       = flags;
    ws->state       = LXLSX_READER_WS_INIT;
    ws->target_path = strdup(target);
    if (!ws->target_path) { free(ws); return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED; }

    *out = ws;
    return LXLSX_READER_NO_ERROR;
}

/* Lazily load worksheet metadata. minizip allows only one open entry at a
 * time, so the metadata pass cannot run while the data stream holds the
 * entry. If data is mid-stream we leave metadata unloaded (accessors return
 * empty); callers that need metadata during reads must trigger a load before
 * the first data read. When the data pump has reached EOF we may safely
 * reclaim the entry. */
lxlsx_reader_error lxlsx_reader_worksheet_ensure_meta(const lxlsx_reader_worksheet *ws_c)
{
    lxlsx_reader_worksheet *ws = (lxlsx_reader_worksheet *)ws_c;
    lxlsx_reader_error      rc;

    if (!ws) return LXLSX_READER_ERROR_NULL_PARAMETER;
    if (ws->meta_loaded) return LXLSX_READER_NO_ERROR;

    /* Data stream active and not exhausted: loading now would require
     * tearing down the pump mid-row, losing unread cells. Decline; metadata
     * accessors will see an empty cache. */
    if (ws->data_opened && ws->pump && !lxlsx_reader_xml_pump_is_eof(ws->pump)) {
        return LXLSX_READER_NO_ERROR;
    }

    /* Reclaim the entry if the data pump is done (or never started). */
    if (ws->data_opened) {
        if (ws->pump) { lxlsx_reader_xml_pump_destroy(ws->pump); ws->pump = NULL; }
        if (ws->zf)   { lxlsx_reader_zip_close_entry(ws->zf);    ws->zf = NULL; }
        ws->data_opened = 0;
    }

    rc = lxlsx_reader_worksheet_meta_load(ws);
    if (rc != LXLSX_READER_NO_ERROR) return rc;
    ws->meta_loaded = 1;
    return LXLSX_READER_NO_ERROR;
}

/* Open the data entry + pump on first use. Callers must route every data
 * read through this so the entry is ready. */
lxlsx_reader_error lxlsx_reader_worksheet_ensure_data_open(lxlsx_reader_worksheet *ws)
{
    if (!ws) return LXLSX_READER_ERROR_NULL_PARAMETER;
    if (ws->data_opened) return LXLSX_READER_NO_ERROR;

    /* Metadata must be loaded before the data entry opens. minizip allows
     * only one active entry per handle, so loading it later while rows are
     * mid-stream would either fail silently or require discarding unread
     * cells. DEFER_METADATA keeps the old pure-streaming behaviour, except
     * merge-follow still needs metadata to produce correct row data. */
    if (!(ws->flags & LXLSX_READER_DEFER_METADATA) ||
        (ws->flags & LXLSX_READER_SKIP_MERGED_FOLLOW)) {
        lxlsx_reader_error rc = lxlsx_reader_worksheet_ensure_meta(ws);
        if (rc != LXLSX_READER_NO_ERROR) return rc;
    }

    ws->zf = lxlsx_reader_zip_open_entry(ws->wb->zip, ws->target_path);
    if (!ws->zf) return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;

    ws->pump = lxlsx_reader_xml_pump_create_zip_file(ws->zf);
    if (!ws->pump) {
        lxlsx_reader_zip_close_entry(ws->zf);
        ws->zf = NULL;
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }
    lxlsx_reader_xml_pump_set_handlers(ws->pump, on_start, on_end, on_text, ws);
    ws->data_opened = 1;
    return LXLSX_READER_NO_ERROR;
}

lxlsx_reader_error lxlsx_reader_worksheet_open_internal(lxlsx_reader_workbook *wb, const char *target,
                                      uint32_t flags, lxlsx_reader_worksheet **out)
{
    return open_internal(wb, target, flags, out);
}

void lxlsx_reader_worksheet_close(lxlsx_reader_worksheet *ws)
{
    if (!ws) return;
    if (ws->pump) lxlsx_reader_xml_pump_destroy(ws->pump);
    if (ws->zf)   lxlsx_reader_zip_close_entry(ws->zf);
    lxlsx_reader_worksheet_meta_free(&ws->meta);
    free(ws->merge_order);
    free(ws->cell_value);
    free(ws->cell_formula);
    free(ws->cell_inline);
    free(ws->skip_tag);
    free(ws->target_path);
    /* Inline rich-text run state. */
    if (ws->inline_runs) {
        size_t i;
        for (i = 0; i < ws->inline_runs_count; i++) {
            free(ws->inline_runs[i].text);
            free(ws->inline_runs[i].font_name);
            free(ws->inline_runs[i].color);
        }
        free(ws->inline_runs);
    }
    free(ws->inline_run_text_buf);
    free(ws->inline_pending_text);
    free(ws->inline_pending_font_name);
    free(ws->inline_pending_color);
    free(ws);
}

/* ------------------------------------------------------------------------- */
/* Pull mode                                                                 */
/* ------------------------------------------------------------------------- */

static lxlsx_reader_error drive(lxlsx_reader_worksheet *ws)
{
    lxlsx_reader_error rc;
    if (ws->parse_error)
        return ws->parse_error;
    if (lxlsx_reader_xml_pump_is_eof(ws->pump) && !lxlsx_reader_xml_pump_is_suspended(ws->pump)) {
        return LXLSX_READER_ERROR_END_OF_DATA;
    }
    if (lxlsx_reader_xml_pump_is_suspended(ws->pump)) {
        rc = lxlsx_reader_xml_pump_resume(ws->pump);
    } else {
        rc = lxlsx_reader_xml_pump_run(ws->pump);
    }
    if (ws->parse_error)
        return ws->parse_error;
    return rc;
}

lxlsx_reader_error lxlsx_reader_worksheet_next_row(lxlsx_reader_worksheet *ws)
{
    lxlsx_reader_error rc;
    if (!ws) return LXLSX_READER_ERROR_NULL_PARAMETER;

    rc = lxlsx_reader_worksheet_ensure_data_open(ws);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    /* Drain anything left from a previous row by calling next_cell until it
     * returns END_OF_DATA. next_cell uses the suspend mechanism, so the pump
     * stops cleanly at the </row> boundary instead of running into the next
     * row (which would steal it from this consumer). */
    if (ws->row_in_progress) {
        lxlsx_cell dummy;
        for (;;) {
            rc = lxlsx_reader_worksheet_next_cell(ws, &dummy);
            if (rc == LXLSX_READER_NO_ERROR)
                continue;
            if (rc != LXLSX_READER_ERROR_END_OF_DATA) {
                ws->pull_mode = LXLSX_READER_WS_PULL_NONE;
                return rc;
            }
            break;
        }
    }

    ws->pending_row_start = 0;
    ws->pending_row_end   = 0;
    ws->pending_cell      = 0;
    ws->pull_mode         = LXLSX_READER_WS_PULL_ROW;

    while (!ws->pending_row_start && !ws->eof) {
        rc = drive(ws);
        if (rc != LXLSX_READER_NO_ERROR) {
            ws->pull_mode = LXLSX_READER_WS_PULL_NONE;
            return rc;
        }
        if (!ws->pending_row_start && lxlsx_reader_xml_pump_is_eof(ws->pump)) break;
    }

    ws->pull_mode = LXLSX_READER_WS_PULL_NONE;
    if (!ws->pending_row_start) return LXLSX_READER_ERROR_END_OF_DATA;

    ws->pending_row_start = 0;
    ws->max_col_seen = 0;
    return LXLSX_READER_NO_ERROR;
}

lxlsx_reader_error lxlsx_reader_worksheet_next_cell(lxlsx_reader_worksheet *ws, lxlsx_cell *out)
{
    if (!ws || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    if (!ws->row_in_progress) return LXLSX_READER_ERROR_END_OF_DATA;

    ws->pending_row_end = 0;
    ws->pending_cell    = 0;
    ws->pull_mode       = LXLSX_READER_WS_PULL_CELL;

    while (!ws->pending_cell && !ws->pending_row_end && !ws->eof) {
        lxlsx_reader_error rc = drive(ws);
        if (rc != LXLSX_READER_NO_ERROR) {
            ws->pull_mode = LXLSX_READER_WS_PULL_NONE;
            return rc;
        }
        if (lxlsx_reader_xml_pump_is_eof(ws->pump) && !ws->pending_cell) break;
    }

    ws->pull_mode = LXLSX_READER_WS_PULL_NONE;

    if (ws->pending_cell) {
        ws->pending_cell = 0;
        emit_cell(ws, out);
        return LXLSX_READER_NO_ERROR;
    }
    /* row ended */
    return LXLSX_READER_ERROR_END_OF_DATA;
}

size_t lxlsx_reader_worksheet_current_row(const lxlsx_reader_worksheet *ws)
{
    return ws ? ws->row_nr : 0;
}

size_t lxlsx_reader_worksheet_max_column_seen(const lxlsx_reader_worksheet *ws)
{
    return ws ? ws->max_col_seen : 0;
}

uint32_t lxlsx_reader_worksheet_flags(const lxlsx_reader_worksheet *ws)
{
    return ws ? ws->flags : 0;
}

/* ------------------------------------------------------------------------- */
/* Push (callback) mode                                                      */
/* ------------------------------------------------------------------------- */

lxlsx_reader_error lxlsx_reader_worksheet_process(lxlsx_reader_worksheet *ws,
                                lxlsx_reader_cell_cb    cell_cb,
                                lxlsx_reader_row_end_cb row_cb,
                                void          *userdata)
{
    lxlsx_reader_error rc;
    if (!ws) return LXLSX_READER_ERROR_NULL_PARAMETER;

    rc = lxlsx_reader_worksheet_ensure_data_open(ws);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    ws->pull_mode      = LXLSX_READER_WS_PULL_NONE;
    ws->user_cell_cb   = cell_cb;
    ws->user_row_cb    = row_cb;
    ws->user_data      = userdata;
    ws->callback_stop  = 0;

    while (!ws->eof && !ws->callback_stop) {
        rc = drive(ws);
        if (rc != LXLSX_READER_NO_ERROR) {
            ws->user_cell_cb = NULL;
            ws->user_row_cb  = NULL;
            return rc;
        }
        if (lxlsx_reader_xml_pump_is_eof(ws->pump)) break;
    }
    ws->user_cell_cb = NULL;
    ws->user_row_cb  = NULL;
    return LXLSX_READER_NO_ERROR;
}

/* ------------------------------------------------------------------------- */
/* Skip                                                                      */
/* ------------------------------------------------------------------------- */

lxlsx_reader_error lxlsx_reader_worksheet_skip_rows(lxlsx_reader_worksheet *ws, size_t n)
{
    if (!ws) return LXLSX_READER_ERROR_NULL_PARAMETER;
    ws->skip_rows_remaining += n;
    return LXLSX_READER_NO_ERROR;
}


/****************************************************************************
 *
 * XLSX worksheet metadata read support.
 *
 ****************************************************************************/

/*
 * Phase-1 sheet metadata parser.
 *
 * The libxlsx reader streaming pump in worksheet.c is single-purpose: it walks
 * <sheetData> and emits cells. But several worksheet-level elements
 * (<mergeCells>, <hyperlinks>, <sheetProtection>, <cols>, <sheetFormatPr>,
 * plus row attrs on <row> elements themselves) live as siblings of
 * <sheetData> and may appear *after* it in the XML stream. They're also too
 * small to justify on-demand re-parsing.
 *
 * This file implements an eager scan at sheet open: open a fresh entry,
 * skip <c> children inside <sheetData> (still tokenising the bytes — minizip
 * has to decompress them anyway), and capture every metadata element into
 * the worksheet's lxlsx_reader_worksheet_meta cache. Accessors below read from this
 * cache.
 *
 * Hyperlinks may carry r:id pointing at xl/worksheets/_rels/sheetN.xml.rels;
 * the rels file is parsed separately to resolve those into absolute URLs.
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "xlsx_private.h"
#include "xlsx_util.h"

/* ------------------------------------------------------------------------- */
/* Helpers                                                                   */
/* ------------------------------------------------------------------------- */

static int attr_truthy(const char *v)
{
    if (!v) return 0;
    return (strcmp(v, "1") == 0 || strcmp(v, "true") == 0) ? 1 : 0;
}

/* ------------------------------------------------------------------------- */
/* Storage helpers                                                           */
/* ------------------------------------------------------------------------- */

static int meta_push_merge(lxlsx_reader_worksheet_meta *m, lxlsx_reader_range r)
{
    if (m->merges_count >= m->merges_cap) {
        size_t nc = m->merges_cap ? m->merges_cap * 2 : 8;
        lxlsx_reader_range *nb = (lxlsx_reader_range *)realloc(m->merges, nc * sizeof(*nb));
        if (!nb) return -1;
        m->merges = nb;
        m->merges_cap = nc;
    }
    m->merges[m->merges_count++] = r;
    return 0;
}

static int meta_push_hyperlink(lxlsx_reader_worksheet_meta *m,
                               lxlsx_reader_range r,
                               const char *url, const char *location,
                               const char *display, const char *tooltip)
{
    struct lxlsx_reader_hyperlink_owned *h;
    if (m->hyperlinks_count >= m->hyperlinks_cap) {
        size_t nc = m->hyperlinks_cap ? m->hyperlinks_cap * 2 : 8;
        struct lxlsx_reader_hyperlink_owned *nb = (struct lxlsx_reader_hyperlink_owned *)realloc(
            m->hyperlinks, nc * sizeof(*nb));
        if (!nb) return -1;
        m->hyperlinks = nb;
        m->hyperlinks_cap = nc;
    }
    h = &m->hyperlinks[m->hyperlinks_count++];
    h->range    = r;
    h->url      = url      ? strdup(url)      : NULL;
    h->location = location ? strdup(location) : NULL;
    h->display  = display  ? strdup(display)  : NULL;
    h->tooltip  = tooltip  ? strdup(tooltip)  : NULL;
    return 0;
}

static struct lxlsx_reader_row_meta *meta_get_or_make_row(lxlsx_reader_worksheet_meta *m, size_t row)
{
    /* Rows arrive in order; a quick check on the tail is enough. */
    if (m->rows_count > 0 && m->rows[m->rows_count - 1].row == row) {
        return &m->rows[m->rows_count - 1];
    }
    if (m->rows_count >= m->rows_cap) {
        size_t nc = m->rows_cap ? m->rows_cap * 2 : 16;
        struct lxlsx_reader_row_meta *nb = (struct lxlsx_reader_row_meta *)realloc(
            m->rows, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->rows = nb;
        m->rows_cap = nc;
    }
    {
        struct lxlsx_reader_row_meta *r = &m->rows[m->rows_count++];
        memset(r, 0, sizeof(*r));
        r->row = row;
        return r;
    }
}

static struct lxlsx_reader_col_meta *meta_push_col(lxlsx_reader_worksheet_meta *m)
{
    if (m->cols_count >= m->cols_cap) {
        size_t nc = m->cols_cap ? m->cols_cap * 2 : 8;
        struct lxlsx_reader_col_meta *nb = (struct lxlsx_reader_col_meta *)realloc(
            m->cols, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->cols = nb;
        m->cols_cap = nc;
    }
    {
        struct lxlsx_reader_col_meta *c = &m->cols[m->cols_count++];
        memset(c, 0, sizeof(*c));
        return c;
    }
}

/* ------------------------------------------------------------------------- */
/* Metadata SAX                                                              */
/* ------------------------------------------------------------------------- */

typedef enum {
    M_INIT = 0,
    M_IN_WORKSHEET,
    M_IN_SHEETDATA,
    M_IN_ROW,
    M_IN_COLS,
    M_IN_MERGECELLS,
    M_IN_HYPERLINKS,
    M_IN_DATAVALIDATIONS,
    M_IN_DATAVALIDATION,    /* parsing a single <dataValidation>; capturing
                              <formula1>/<formula2> children */
    M_IN_DV_FORMULA1,
    M_IN_DV_FORMULA2,
    M_IN_AUTOFILTER,
    M_IN_FILTERCOLUMN,
    M_IN_FILTERS,           /* <filters> inside a <filterColumn> */
    /* §8.2.5 page setup — text-collecting states for header/footer. */
    M_IN_HEADERFOOTER,
    M_IN_ODD_HEADER,
    M_IN_ODD_FOOTER,
    M_IN_EVEN_HEADER,
    M_IN_EVEN_FOOTER,
    M_IN_FIRST_HEADER,
    M_IN_FIRST_FOOTER,
    M_IN_CONDFMT,
    M_IN_CFRULE,
    M_IN_CF_FORMULA,
    M_SKIP
} m_state;

typedef struct {
    lxlsx_reader_worksheet_meta *m;
    const lxlsx_reader_rel_map *rels;
    m_state             state;
    m_state             state_before_skip;
    int                 skip_depth;
    char               *skip_tag;
    /* DV text accumulator — ECMA-376 puts formulas inside <formula1>/<formula2>
     * as text children, so we collect them like definedName text. */
    char               *txt;
    size_t              txt_len;
    size_t              txt_cap;
} m_ctx;

/* DV/filter helpers — append entries. */

static struct lxlsx_reader_dv_owned *meta_push_dv(lxlsx_reader_worksheet_meta *m)
{
    if (m->dvs_count >= m->dvs_cap) {
        size_t nc = m->dvs_cap ? m->dvs_cap * 2 : 4;
        struct lxlsx_reader_dv_owned *nb = (struct lxlsx_reader_dv_owned *)realloc(
            m->dvs, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->dvs = nb;
        m->dvs_cap = nc;
    }
    {
        struct lxlsx_reader_dv_owned *d = &m->dvs[m->dvs_count++];
        memset(d, 0, sizeof(*d));
        return d;
    }
}

static struct lxlsx_reader_filter_column_owned *meta_push_filter_column(lxlsx_reader_worksheet_meta *m)
{
    if (m->filter_columns_count >= m->filter_columns_cap) {
        size_t nc = m->filter_columns_cap ? m->filter_columns_cap * 2 : 4;
        struct lxlsx_reader_filter_column_owned *nb = (struct lxlsx_reader_filter_column_owned *)realloc(
            m->filter_columns, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->filter_columns = nb;
        m->filter_columns_cap = nc;
    }
    {
        struct lxlsx_reader_filter_column_owned *fc = &m->filter_columns[m->filter_columns_count++];
        memset(fc, 0, sizeof(*fc));
        return fc;
    }
}

/* For LXLSX_READER_FILTER_LIST: append to a NULL-terminated owned values array. */
static void filter_column_push_value(struct lxlsx_reader_filter_column_owned *fc, const char *val)
{
    size_t n = 0;
    char **nv;
    if (!val) return;
    if (fc->values) while (fc->values[n]) n++;
    nv = (char **)realloc(fc->values, (n + 2) * sizeof(*nv));
    if (!nv) return;
    nv[n]     = strdup(val);
    nv[n + 1] = NULL;
    fc->values = nv;
}

static void enter_skip(m_ctx *c, const char *tag)
{
    free(c->skip_tag);
    c->skip_tag = strdup(tag);
    c->state_before_skip = c->state;
    c->state = M_SKIP;
    c->skip_depth = 1;
}

static void m_on_start(void *ud, const char *name, const char **attrs)
{
    m_ctx *c = (m_ctx *)ud;

    if (c->state == M_SKIP) { c->skip_depth++; return; }

    switch (c->state) {
    case M_INIT:
        if (lxlsx_reader_xml_name_eq(name, "worksheet")) c->state = M_IN_WORKSHEET;
        break;

    case M_IN_WORKSHEET:
        if (lxlsx_reader_xml_name_eq(name, "sheetData")) {
            c->state = M_IN_SHEETDATA;
        } else if (lxlsx_reader_xml_name_eq(name, "sheetFormatPr")) {
            const char *drh = lxlsx_reader_xml_attr(attrs, "defaultRowHeight");
            const char *dcw = lxlsx_reader_xml_attr(attrs, "defaultColWidth");
            if (drh) {
                c->m->has_default_row_height = 1;
                c->m->default_row_height = strtod(drh, NULL);
            }
            if (dcw) {
                c->m->has_default_col_width = 1;
                c->m->default_col_width = strtod(dcw, NULL);
            }
            /* element is empty / self-closing in practice — stay in WORKSHEET */
        } else if (lxlsx_reader_xml_name_eq(name, "cols")) {
            c->state = M_IN_COLS;
        } else if (lxlsx_reader_xml_name_eq(name, "mergeCells")) {
            c->state = M_IN_MERGECELLS;
        } else if (lxlsx_reader_xml_name_eq(name, "hyperlinks")) {
            c->state = M_IN_HYPERLINKS;
        } else if (lxlsx_reader_xml_name_eq(name, "dataValidations")) {
            c->state = M_IN_DATAVALIDATIONS;
        } else if (lxlsx_reader_xml_name_eq(name, "autoFilter")) {
            const char *ref = lxlsx_reader_xml_attr(attrs, "ref");
            c->m->autofilter_present = 1;
            if (ref) {
                free(c->m->autofilter_range);
                c->m->autofilter_range = strdup(ref);
            }
            c->state = M_IN_AUTOFILTER;
        } else if (lxlsx_reader_xml_name_eq(name, "pageMargins")) {
            const char *l = lxlsx_reader_xml_attr(attrs, "left");
            const char *r = lxlsx_reader_xml_attr(attrs, "right");
            const char *t = lxlsx_reader_xml_attr(attrs, "top");
            const char *b = lxlsx_reader_xml_attr(attrs, "bottom");
            const char *h = lxlsx_reader_xml_attr(attrs, "header");
            const char *f = lxlsx_reader_xml_attr(attrs, "footer");
            c->m->page_has_margins = 1;
            if (l) c->m->page_margin_left   = strtod(l, NULL);
            if (r) c->m->page_margin_right  = strtod(r, NULL);
            if (t) c->m->page_margin_top    = strtod(t, NULL);
            if (b) c->m->page_margin_bottom = strtod(b, NULL);
            if (h) c->m->page_margin_header = strtod(h, NULL);
            if (f) c->m->page_margin_footer = strtod(f, NULL);
        } else if (lxlsx_reader_xml_name_eq(name, "pageSetup")) {
            const char *v;
            c->m->page_has_setup = 1;
            if ((v = lxlsx_reader_xml_attr(attrs, "paperSize")))    c->m->page_paper_size    = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "fitToWidth")))   c->m->page_fit_to_width  = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "fitToHeight")))  c->m->page_fit_to_height = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "scale")))        c->m->page_scale         = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "orientation"))
                && strcmp(v, "landscape") == 0)            c->m->page_orientation_landscape = 1;
            if ((v = lxlsx_reader_xml_attr(attrs, "horizontalDpi")))   c->m->page_horizontal_dpi = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "verticalDpi")))     c->m->page_vertical_dpi   = (int)strtol(v, NULL, 10);
            if ((v = lxlsx_reader_xml_attr(attrs, "firstPageNumber"))) c->m->page_first_page_number = (int)strtol(v, NULL, 10);
            c->m->page_use_first_page_number = attr_truthy(lxlsx_reader_xml_attr(attrs, "useFirstPageNumber"));
        } else if (lxlsx_reader_xml_name_eq(name, "printOptions")) {
            c->m->page_print_h_centered  = attr_truthy(lxlsx_reader_xml_attr(attrs, "horizontalCentered"));
            c->m->page_print_v_centered  = attr_truthy(lxlsx_reader_xml_attr(attrs, "verticalCentered"));
            c->m->page_print_grid_lines  = attr_truthy(lxlsx_reader_xml_attr(attrs, "gridLines"));
            c->m->page_print_headings    = attr_truthy(lxlsx_reader_xml_attr(attrs, "headings"));
        } else if (lxlsx_reader_xml_name_eq(name, "conditionalFormatting")) {
            const char *sqref = lxlsx_reader_xml_attr(attrs, "sqref");
            if (c->m->cf_blocks_count >= c->m->cf_blocks_cap) {
                size_t nc = c->m->cf_blocks_cap ? c->m->cf_blocks_cap * 2 : 4;
                struct lxlsx_reader_cf_block_owned *nb = realloc(c->m->cf_blocks,
                                                        nc * sizeof(*nb));
                if (nb) { c->m->cf_blocks = nb; c->m->cf_blocks_cap = nc; }
            }
            if (c->m->cf_blocks_count < c->m->cf_blocks_cap) {
                struct lxlsx_reader_cf_block_owned *bk = &c->m->cf_blocks[c->m->cf_blocks_count++];
                memset(bk, 0, sizeof(*bk));
                if (sqref) bk->sqref = strdup(sqref);
            }
            c->state = M_IN_CONDFMT;
        } else if (lxlsx_reader_xml_name_eq(name, "headerFooter")) {
            c->m->page_different_odd_even = attr_truthy(lxlsx_reader_xml_attr(attrs, "differentOddEven"));
            c->m->page_different_first    = attr_truthy(lxlsx_reader_xml_attr(attrs, "differentFirst"));
            c->m->page_scale_with_doc     = attr_truthy(lxlsx_reader_xml_attr(attrs, "scaleWithDoc"));
            c->m->page_align_with_margins = attr_truthy(lxlsx_reader_xml_attr(attrs, "alignWithMargins"));
            c->state = M_IN_HEADERFOOTER;
        } else if (lxlsx_reader_xml_name_eq(name, "sheetProtection")) {
            const char *v;
            c->m->prot_present = 1;
            v = lxlsx_reader_xml_attr(attrs, "password");
            if (v) {
                size_t n = strlen(v);
                if (n >= sizeof(c->m->prot_hash)) n = sizeof(c->m->prot_hash) - 1;
                memcpy(c->m->prot_hash, v, n);
                c->m->prot_hash[n] = 0;
            }
            c->m->prot_sheet               = attr_truthy(lxlsx_reader_xml_attr(attrs, "sheet"));
            c->m->prot_content             = attr_truthy(lxlsx_reader_xml_attr(attrs, "content"));
            c->m->prot_objects             = attr_truthy(lxlsx_reader_xml_attr(attrs, "objects"));
            c->m->prot_scenarios           = attr_truthy(lxlsx_reader_xml_attr(attrs, "scenarios"));
            c->m->prot_format_cells        = attr_truthy(lxlsx_reader_xml_attr(attrs, "formatCells"));
            c->m->prot_format_columns      = attr_truthy(lxlsx_reader_xml_attr(attrs, "formatColumns"));
            c->m->prot_format_rows         = attr_truthy(lxlsx_reader_xml_attr(attrs, "formatRows"));
            c->m->prot_insert_columns      = attr_truthy(lxlsx_reader_xml_attr(attrs, "insertColumns"));
            c->m->prot_insert_rows         = attr_truthy(lxlsx_reader_xml_attr(attrs, "insertRows"));
            c->m->prot_insert_hyperlinks   = attr_truthy(lxlsx_reader_xml_attr(attrs, "insertHyperlinks"));
            c->m->prot_delete_columns      = attr_truthy(lxlsx_reader_xml_attr(attrs, "deleteColumns"));
            c->m->prot_delete_rows         = attr_truthy(lxlsx_reader_xml_attr(attrs, "deleteRows"));
            c->m->prot_select_locked_cells = attr_truthy(lxlsx_reader_xml_attr(attrs, "selectLockedCells"));
            c->m->prot_sort                = attr_truthy(lxlsx_reader_xml_attr(attrs, "sort"));
            c->m->prot_auto_filter         = attr_truthy(lxlsx_reader_xml_attr(attrs, "autoFilter"));
            c->m->prot_pivot_tables        = attr_truthy(lxlsx_reader_xml_attr(attrs, "pivotTables"));
            c->m->prot_select_unlocked_cells = attr_truthy(lxlsx_reader_xml_attr(attrs, "selectUnlockedCells"));
        }
        break;

    case M_IN_COLS:
        if (lxlsx_reader_xml_name_eq(name, "col")) {
            const char *vmin = lxlsx_reader_xml_attr(attrs, "min");
            const char *vmax = lxlsx_reader_xml_attr(attrs, "max");
            const char *vw   = lxlsx_reader_xml_attr(attrs, "width");
            const char *cw   = lxlsx_reader_xml_attr(attrs, "customWidth");
            const char *hd   = lxlsx_reader_xml_attr(attrs, "hidden");
            const char *ol   = lxlsx_reader_xml_attr(attrs, "outlineLevel");
            const char *cl   = lxlsx_reader_xml_attr(attrs, "collapsed");
            struct lxlsx_reader_col_meta *col = meta_push_col(c->m);
            if (col) {
                col->min = vmin ? (size_t)strtoul(vmin, NULL, 10) : 0;
                col->max = vmax ? (size_t)strtoul(vmax, NULL, 10) : col->min;
                if (vw) {
                    col->width = strtod(vw, NULL);
                    /* OOXML lists customWidth=1 when an explicit width was set;
                     * absent customWidth on a width attr is the default-derived
                     * value the writer emits. Treat any present width as
                     * "user-visible width". */
                    col->has_width = 1;
                    (void)cw;
                }
                col->hidden        = attr_truthy(hd);
                col->outline_level = ol ? (int)strtol(ol, NULL, 10) : 0;
                col->collapsed     = attr_truthy(cl);
            }
        }
        break;

    case M_IN_MERGECELLS:
        if (lxlsx_reader_xml_name_eq(name, "mergeCell")) {
            const char *ref = lxlsx_reader_xml_attr(attrs, "ref");
            lxlsx_reader_range r;
            lxlsx_reader_parse_a1_range(ref, &r);
            if (r.first_row && r.first_col) meta_push_merge(c->m, r);
        }
        break;

    case M_IN_DATAVALIDATIONS:
        if (lxlsx_reader_xml_name_eq(name, "dataValidation")) {
            struct lxlsx_reader_dv_owned *d = meta_push_dv(c->m);
            const char *v;
            if (!d) break;
            if ((v = lxlsx_reader_xml_attr(attrs, "type")))               d->type        = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "operator")))           d->operator_   = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "errorStyle")))         d->error_style = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "prompt")))             d->prompt      = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "promptTitle")))        d->prompt_title= strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "error")))              d->error       = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "errorTitle")))         d->error_title = strdup(v);
            if ((v = lxlsx_reader_xml_attr(attrs, "sqref")))              d->sqref       = strdup(v);
            d->allow_blank          = attr_truthy(lxlsx_reader_xml_attr(attrs, "allowBlank"));
            d->show_drop_down       = attr_truthy(lxlsx_reader_xml_attr(attrs, "showDropDown"));
            d->show_input_message   = attr_truthy(lxlsx_reader_xml_attr(attrs, "showInputMessage"));
            d->show_error_message   = attr_truthy(lxlsx_reader_xml_attr(attrs, "showErrorMessage"));
            c->state = M_IN_DATAVALIDATION;
        }
        break;

    case M_IN_DATAVALIDATION:
        if (lxlsx_reader_xml_name_eq(name, "formula1")) {
            lxlsx_reader_buf_reset(c->txt, &c->txt_len);
            c->state = M_IN_DV_FORMULA1;
        } else if (lxlsx_reader_xml_name_eq(name, "formula2")) {
            lxlsx_reader_buf_reset(c->txt, &c->txt_len);
            c->state = M_IN_DV_FORMULA2;
        } else {
            enter_skip(c, name);
        }
        break;

    case M_IN_AUTOFILTER:
        if (lxlsx_reader_xml_name_eq(name, "filterColumn")) {
            const char *cid = lxlsx_reader_xml_attr(attrs, "colId");
            struct lxlsx_reader_filter_column_owned *fc = meta_push_filter_column(c->m);
            if (fc) {
                fc->col_id = cid ? (int)strtol(cid, NULL, 10) : 0;
                fc->kind   = LXLSX_READER_FILTER_NONE;
            }
            c->state = M_IN_FILTERCOLUMN;
        } else {
            enter_skip(c, name);
        }
        break;

    case M_IN_FILTERCOLUMN:
        if (c->m->filter_columns_count == 0) { enter_skip(c, name); break; }
        {
            struct lxlsx_reader_filter_column_owned *fc =
                &c->m->filter_columns[c->m->filter_columns_count - 1];
            if (lxlsx_reader_xml_name_eq(name, "filters")) {
                fc->kind = LXLSX_READER_FILTER_LIST;
                c->state = M_IN_FILTERS;
            } else if (lxlsx_reader_xml_name_eq(name, "customFilters")) {
                const char *and_attr = lxlsx_reader_xml_attr(attrs, "and");
                fc->kind = LXLSX_READER_FILTER_CUSTOM;
                fc->custom_and = attr_truthy(and_attr);
                /* children customFilter come in <customFilter operator val>;
                 * parse them inline by transitioning into a list-style scan. */
                c->state = M_IN_FILTERS;  /* reuse */
            } else if (lxlsx_reader_xml_name_eq(name, "top10")) {
                const char *val_attr  = lxlsx_reader_xml_attr(attrs, "val");
                const char *top_attr  = lxlsx_reader_xml_attr(attrs, "top");
                const char *pct_attr  = lxlsx_reader_xml_attr(attrs, "percent");
                fc->kind      = LXLSX_READER_FILTER_TOP10;
                fc->top       = top_attr ? attr_truthy(top_attr) : 1;
                fc->percent   = attr_truthy(pct_attr);
                fc->top_value = val_attr ? strtod(val_attr, NULL) : 0;
                enter_skip(c, name);
            } else if (lxlsx_reader_xml_name_eq(name, "dynamicFilter")) {
                fc->kind = LXLSX_READER_FILTER_DYNAMIC;
                enter_skip(c, name);
            } else {
                enter_skip(c, name);
            }
        }
        break;

    case M_IN_CONDFMT:
        if (lxlsx_reader_xml_name_eq(name, "cfRule") && c->m->cf_blocks_count > 0) {
            struct lxlsx_reader_cf_block_owned *bk = &c->m->cf_blocks[c->m->cf_blocks_count - 1];
            const char *v;
            if (bk->rules_count >= bk->rules_cap) {
                size_t nc = bk->rules_cap ? bk->rules_cap * 2 : 2;
                struct lxlsx_reader_cf_rule_owned *nb = realloc(bk->rules, nc * sizeof(*nb));
                if (nb) { bk->rules = nb; bk->rules_cap = nc; }
            }
            if (bk->rules_count < bk->rules_cap) {
                struct lxlsx_reader_cf_rule_owned *r = &bk->rules[bk->rules_count++];
                memset(r, 0, sizeof(*r));
                r->dxf_id = -1;
                if ((v = lxlsx_reader_xml_attr(attrs, "type")))        r->type        = strdup(v);
                if ((v = lxlsx_reader_xml_attr(attrs, "operator")))    r->operator_   = strdup(v);
                if ((v = lxlsx_reader_xml_attr(attrs, "priority")))    r->priority    = (int)strtol(v, NULL, 10);
                if ((v = lxlsx_reader_xml_attr(attrs, "dxfId")))       r->dxf_id      = (int)strtol(v, NULL, 10);
                if ((v = lxlsx_reader_xml_attr(attrs, "rank")))        r->rank        = strtod(v, NULL);
                if ((v = lxlsx_reader_xml_attr(attrs, "text")))        r->text        = strdup(v);
                if ((v = lxlsx_reader_xml_attr(attrs, "timePeriod")))  r->time_period = strdup(v);
                r->stop_if_true = attr_truthy(lxlsx_reader_xml_attr(attrs, "stopIfTrue"));
                r->percent      = attr_truthy(lxlsx_reader_xml_attr(attrs, "percent"));
                r->bottom       = attr_truthy(lxlsx_reader_xml_attr(attrs, "bottom"));
            }
            c->state = M_IN_CFRULE;
        }
        break;

    case M_IN_CFRULE:
        if (lxlsx_reader_xml_name_eq(name, "formula")) {
            lxlsx_reader_buf_reset(c->txt, &c->txt_len);
            c->state = M_IN_CF_FORMULA;
        } else {
            /* colorScale, dataBar, iconSet — skip nested element trees. */
            enter_skip(c, name);
        }
        break;

    case M_IN_HEADERFOOTER:
        if      (lxlsx_reader_xml_name_eq(name, "oddHeader"))   { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_ODD_HEADER;   }
        else if (lxlsx_reader_xml_name_eq(name, "oddFooter"))   { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_ODD_FOOTER;   }
        else if (lxlsx_reader_xml_name_eq(name, "evenHeader"))  { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_EVEN_HEADER;  }
        else if (lxlsx_reader_xml_name_eq(name, "evenFooter"))  { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_EVEN_FOOTER;  }
        else if (lxlsx_reader_xml_name_eq(name, "firstHeader")) { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_FIRST_HEADER; }
        else if (lxlsx_reader_xml_name_eq(name, "firstFooter")) { lxlsx_reader_buf_reset(c->txt, &c->txt_len); c->state = M_IN_FIRST_FOOTER; }
        else                                            enter_skip(c, name);
        break;

    case M_IN_FILTERS:
        if (c->m->filter_columns_count == 0) { enter_skip(c, name); break; }
        {
            struct lxlsx_reader_filter_column_owned *fc =
                &c->m->filter_columns[c->m->filter_columns_count - 1];
            if (lxlsx_reader_xml_name_eq(name, "filter")) {
                const char *v = lxlsx_reader_xml_attr(attrs, "val");
                if (v) filter_column_push_value(fc, v);
                enter_skip(c, name);
            } else if (lxlsx_reader_xml_name_eq(name, "customFilter")) {
                const char *op  = lxlsx_reader_xml_attr(attrs, "operator");
                const char *val = lxlsx_reader_xml_attr(attrs, "val");
                if (!fc->custom_op_1) {
                    fc->custom_op_1  = op  ? strdup(op)  : strdup("equal");
                    fc->custom_val_1 = val ? strdup(val) : NULL;
                } else if (!fc->custom_op_2) {
                    fc->custom_op_2  = op  ? strdup(op)  : strdup("equal");
                    fc->custom_val_2 = val ? strdup(val) : NULL;
                }
                enter_skip(c, name);
            } else {
                enter_skip(c, name);
            }
        }
        break;

    case M_IN_HYPERLINKS:
        if (lxlsx_reader_xml_name_eq(name, "hyperlink")) {
            const char *ref     = lxlsx_reader_xml_attr(attrs, "ref");
            const char *loc     = lxlsx_reader_xml_attr(attrs, "location");
            const char *display = lxlsx_reader_xml_attr(attrs, "display");
            const char *tooltip = lxlsx_reader_xml_attr(attrs, "tooltip");
            const char *rid     = lxlsx_reader_xml_attr(attrs, "id");
            const char *url     = NULL;
            lxlsx_reader_range r;

            if (!rid) {
                /* "r:id" — namespace lookup by walking attrs. */
                const char **a = attrs;
                while (a && *a) {
                    if (lxlsx_reader_xml_name_eq(*a, "id")) { rid = *(a + 1); break; }
                    a += 2;
                }
            }
            if (rid) url = lxlsx_reader_rel_map_target_for_id(c->rels, rid);

            lxlsx_reader_parse_a1_range(ref, &r);
            if (r.first_row && r.first_col)
                meta_push_hyperlink(c->m, r, url, loc, display, tooltip);
        }
        break;

    case M_IN_SHEETDATA:
        if (lxlsx_reader_xml_name_eq(name, "row")) {
            const char *r_attr  = lxlsx_reader_xml_attr(attrs, "r");
            const char *ht      = lxlsx_reader_xml_attr(attrs, "ht");
            const char *hidden  = lxlsx_reader_xml_attr(attrs, "hidden");
            const char *ol      = lxlsx_reader_xml_attr(attrs, "outlineLevel");
            const char *cl      = lxlsx_reader_xml_attr(attrs, "collapsed");
            const char *ch      = lxlsx_reader_xml_attr(attrs, "customHeight");
            size_t row_nr = r_attr ? (size_t)strtoul(r_attr, NULL, 10) : 0;
            /* Only cache rows that carry at least one metadata attr — else
             * getRowOptions(N) would return an empty struct for rows whose
             * only purpose is holding cells. */
            int has_meta = (ht || ol ||
                            attr_truthy(hidden) ||
                            attr_truthy(cl) ||
                            attr_truthy(ch));
            if (row_nr > 0 && has_meta) {
                struct lxlsx_reader_row_meta *rm = meta_get_or_make_row(c->m, row_nr);
                if (rm) {
                    if (ht) {
                        rm->has_height = 1;
                        rm->height = strtod(ht, NULL);
                    }
                    rm->hidden        = attr_truthy(hidden);
                    rm->outline_level = ol ? (int)strtol(ol, NULL, 10) : 0;
                    rm->collapsed     = attr_truthy(cl);
                    rm->custom_height = attr_truthy(ch);
                }
            }
            c->state = M_IN_ROW;
        }
        break;

    case M_IN_ROW:
        /* Skip every <c> and child — metadata only cares about row attrs. */
        enter_skip(c, name);
        break;

    default:
        break;
    }
}

static void m_on_end(void *ud, const char *name)
{
    m_ctx *c = (m_ctx *)ud;

    if (c->state == M_SKIP) {
        c->skip_depth--;
        if (c->skip_depth == 0 && c->skip_tag &&
            lxlsx_reader_xml_name_eq(name, c->skip_tag)) {
            free(c->skip_tag);
            c->skip_tag = NULL;
            c->state = c->state_before_skip;
        }
        return;
    }

    switch (c->state) {
    case M_IN_ROW:
        if (lxlsx_reader_xml_name_eq(name, "row")) c->state = M_IN_SHEETDATA;
        break;
    case M_IN_SHEETDATA:
        if (lxlsx_reader_xml_name_eq(name, "sheetData")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_COLS:
        if (lxlsx_reader_xml_name_eq(name, "cols")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_MERGECELLS:
        if (lxlsx_reader_xml_name_eq(name, "mergeCells")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_HYPERLINKS:
        if (lxlsx_reader_xml_name_eq(name, "hyperlinks")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_DATAVALIDATION:
        if (lxlsx_reader_xml_name_eq(name, "dataValidation")) c->state = M_IN_DATAVALIDATIONS;
        break;
    case M_IN_DATAVALIDATIONS:
        if (lxlsx_reader_xml_name_eq(name, "dataValidations")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_DV_FORMULA1:
        if (lxlsx_reader_xml_name_eq(name, "formula1")) {
            if (c->m->dvs_count > 0 && c->txt_len > 0) {
                struct lxlsx_reader_dv_owned *d = &c->m->dvs[c->m->dvs_count - 1];
                free(d->formula1);
                d->formula1 = strdup(c->txt);
            }
            c->state = M_IN_DATAVALIDATION;
        }
        break;
    case M_IN_DV_FORMULA2:
        if (lxlsx_reader_xml_name_eq(name, "formula2")) {
            if (c->m->dvs_count > 0 && c->txt_len > 0) {
                struct lxlsx_reader_dv_owned *d = &c->m->dvs[c->m->dvs_count - 1];
                free(d->formula2);
                d->formula2 = strdup(c->txt);
            }
            c->state = M_IN_DATAVALIDATION;
        }
        break;
    case M_IN_FILTERCOLUMN:
        if (lxlsx_reader_xml_name_eq(name, "filterColumn")) c->state = M_IN_AUTOFILTER;
        break;
    case M_IN_FILTERS:
        if (lxlsx_reader_xml_name_eq(name, "filters") ||
            lxlsx_reader_xml_name_eq(name, "customFilters")) {
            c->state = M_IN_FILTERCOLUMN;
        }
        break;
    case M_IN_AUTOFILTER:
        if (lxlsx_reader_xml_name_eq(name, "autoFilter")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_HEADERFOOTER:
        if (lxlsx_reader_xml_name_eq(name, "headerFooter")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_CFRULE:
        if (lxlsx_reader_xml_name_eq(name, "cfRule")) c->state = M_IN_CONDFMT;
        break;
    case M_IN_CONDFMT:
        if (lxlsx_reader_xml_name_eq(name, "conditionalFormatting"))
            c->state = M_IN_WORKSHEET;
        break;
    case M_IN_CF_FORMULA:
        if (lxlsx_reader_xml_name_eq(name, "formula")) {
            if (c->m->cf_blocks_count > 0) {
                struct lxlsx_reader_cf_block_owned *bk = &c->m->cf_blocks[c->m->cf_blocks_count - 1];
                if (bk->rules_count > 0 && c->txt_len > 0) {
                    struct lxlsx_reader_cf_rule_owned *r = &bk->rules[bk->rules_count - 1];
                    if (!r->formula1)      r->formula1 = strdup(c->txt);
                    else if (!r->formula2) r->formula2 = strdup(c->txt);
                }
            }
            c->state = M_IN_CFRULE;
        }
        break;
#define HF_END(state_, tagname_, slot_)                              \
    case state_:                                                     \
        if (lxlsx_reader_xml_name_eq(name, tagname_)) {                       \
            free(c->m->slot_);                                       \
            c->m->slot_ = c->txt_len > 0 ? strdup(c->txt) : NULL;    \
            c->state = M_IN_HEADERFOOTER;                            \
        } break;
    HF_END(M_IN_ODD_HEADER,   "oddHeader",   page_odd_header)
    HF_END(M_IN_ODD_FOOTER,   "oddFooter",   page_odd_footer)
    HF_END(M_IN_EVEN_HEADER,  "evenHeader",  page_even_header)
    HF_END(M_IN_EVEN_FOOTER,  "evenFooter",  page_even_footer)
    HF_END(M_IN_FIRST_HEADER, "firstHeader", page_first_header)
    HF_END(M_IN_FIRST_FOOTER, "firstFooter", page_first_footer)
#undef HF_END
    case M_IN_WORKSHEET:
        if (lxlsx_reader_xml_name_eq(name, "worksheet")) c->state = M_INIT;
        break;
    default:
        break;
    }
}

/* Text capture for <formula1>/<formula2> and header/footer text fields. */
static void m_on_text(void *ud, const char *text, int len)
{
    m_ctx *c = (m_ctx *)ud;
    if (len <= 0) return;
    switch (c->state) {
    case M_IN_DV_FORMULA1:
    case M_IN_DV_FORMULA2:
    case M_IN_ODD_HEADER:
    case M_IN_ODD_FOOTER:
    case M_IN_EVEN_HEADER:
    case M_IN_EVEN_FOOTER:
    case M_IN_FIRST_HEADER:
    case M_IN_FIRST_FOOTER:
    case M_IN_CF_FORMULA:
        lxlsx_reader_buf_append(&c->txt, &c->txt_len, &c->txt_cap,
                                text, (size_t)len);
        break;
    default: break;
    }
}

/* ------------------------------------------------------------------------- */
/* Public load / free                                                         */
/* ------------------------------------------------------------------------- */

lxlsx_reader_error lxlsx_reader_worksheet_meta_load(lxlsx_reader_worksheet *ws)
{
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    lxlsx_reader_rel_map rels;
    m_ctx         ctx;
    lxlsx_reader_error     rc;

    if (!ws || !ws->wb || !ws->target_path) return LXLSX_READER_ERROR_NULL_PARAMETER;

    /* Sheet rels first (required to resolve external hyperlinks). */
    rc = lxlsx_reader_load_rels(ws->wb->zip, ws->target_path, &rels, 1);
    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_rel_map_free(&rels);
        return rc;
    }

    zf = lxlsx_reader_zip_open_entry(ws->wb->zip, ws->target_path);
    if (!zf) {
        lxlsx_reader_rel_map_free(&rels);
        return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;
    }

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxlsx_reader_zip_close_entry(zf);
        lxlsx_reader_rel_map_free(&rels);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }

    memset(&ctx, 0, sizeof(ctx));
    ctx.m     = &ws->meta;
    ctx.rels  = &rels;
    ctx.state = M_INIT;

    lxlsx_reader_xml_pump_set_handlers(pump, m_on_start, m_on_end, m_on_text, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);

    free(ctx.skip_tag);
    free(ctx.txt);
    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    lxlsx_reader_rel_map_free(&rels);
    return rc;
}

void lxlsx_reader_worksheet_meta_free(lxlsx_reader_worksheet_meta *m)
{
    size_t i;
    if (!m) return;
    free(m->merges);
    if (m->hyperlinks) {
        for (i = 0; i < m->hyperlinks_count; i++) {
            free(m->hyperlinks[i].url);
            free(m->hyperlinks[i].location);
            free(m->hyperlinks[i].display);
            free(m->hyperlinks[i].tooltip);
        }
        free(m->hyperlinks);
    }
    free(m->rows);
    free(m->cols);
    if (m->dvs) {
        for (i = 0; i < m->dvs_count; i++) {
            free(m->dvs[i].type);
            free(m->dvs[i].operator_);
            free(m->dvs[i].error_style);
            free(m->dvs[i].formula1);
            free(m->dvs[i].formula2);
            free(m->dvs[i].prompt);
            free(m->dvs[i].prompt_title);
            free(m->dvs[i].error);
            free(m->dvs[i].error_title);
            free(m->dvs[i].sqref);
        }
        free(m->dvs);
    }
    free(m->autofilter_range);
    if (m->filter_columns) {
        for (i = 0; i < m->filter_columns_count; i++) {
            if (m->filter_columns[i].values) {
                size_t j = 0;
                while (m->filter_columns[i].values[j]) {
                    free(m->filter_columns[i].values[j]);
                    j++;
                }
                free(m->filter_columns[i].values);
            }
            free(m->filter_columns[i].custom_op_1);
            free(m->filter_columns[i].custom_val_1);
            free(m->filter_columns[i].custom_op_2);
            free(m->filter_columns[i].custom_val_2);
        }
        free(m->filter_columns);
    }
    free(m->filter_columns_pub);
    free(m->page_odd_header);
    free(m->page_odd_footer);
    free(m->page_even_header);
    free(m->page_even_footer);
    free(m->page_first_header);
    free(m->page_first_footer);
    if (m->cf_blocks) {
        for (i = 0; i < m->cf_blocks_count; i++) {
            size_t j;
            free(m->cf_blocks[i].sqref);
            for (j = 0; j < m->cf_blocks[i].rules_count; j++) {
                free(m->cf_blocks[i].rules[j].type);
                free(m->cf_blocks[i].rules[j].operator_);
                free(m->cf_blocks[i].rules[j].text);
                free(m->cf_blocks[i].rules[j].time_period);
                free(m->cf_blocks[i].rules[j].formula1);
                free(m->cf_blocks[i].rules[j].formula2);
            }
            free(m->cf_blocks[i].rules);
        }
        free(m->cf_blocks);
    }
    free(m->cf_rules_pub);
    memset(m, 0, sizeof(*m));
}

/* ------------------------------------------------------------------------- */
/* Public accessors                                                          */
/* ------------------------------------------------------------------------- */

size_t lxlsx_reader_worksheet_merged_count(const lxlsx_reader_worksheet *ws)
{
    if (!ws) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    return ws->meta.merges_count;
}

int lxlsx_reader_worksheet_merged_get(const lxlsx_reader_worksheet *ws, size_t idx, lxlsx_reader_range *out)
{
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    if (idx >= ws->meta.merges_count) return 0;
    *out = ws->meta.merges[idx];
    return 1;
}

/* Comparator for the merge_key array used to build merge_order. */
struct merge_key { size_t idx; size_t last_row; };
static int merge_key_cmp(const void *a, const void *b)
{
    size_t la = ((const struct merge_key *)a)->last_row;
    size_t lb = ((const struct merge_key *)b)->last_row;
    if (la < lb) return -1;
    if (la > lb) return 1;
    return 0;
}

/* Build merge_order: indices into meta.merges sorted by ascending last_row.
 * Uses a merge_key array so the qsort comparator is self-contained
 * (qsort_r is not portable across glibc/BSD/MSVC). On OOM the index is left
 * NULL and in_merge_follow falls back to a plain linear scan. */
static void build_merge_index(lxlsx_reader_worksheet *ws)
{
    size_t n = ws->meta.merges_count;

    ws->merge_index_built = 1;
    ws->merge_order       = NULL;
    ws->merge_cursor      = 0;
    ws->merge_last_row    = 0;
    if (n == 0) return;

    struct merge_key *keys = (struct merge_key *)malloc(n * sizeof(*keys));
    if (!keys) return;
    for (size_t i = 0; i < n; i++) {
        keys[i].idx      = i;
        keys[i].last_row = ws->meta.merges[i].last_row;
    }
    qsort(keys, n, sizeof(*keys), merge_key_cmp);

    ws->merge_order = (size_t *)malloc(n * sizeof(size_t));
    if (!ws->merge_order) { free(keys); return; }
    for (size_t i = 0; i < n; i++) ws->merge_order[i] = keys[i].idx;
    free(keys);
}

int lxlsx_reader_worksheet_in_merge_follow(lxlsx_reader_worksheet *ws, size_t row, size_t col)
{
    size_t i, n;
    if (!ws) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    n = ws->meta.merges_count;
    if (n == 0) return 0;

    if (!ws->merge_index_built) build_merge_index(ws);

    /* Index build failed (OOM): fall back to a plain linear scan. */
    if (!ws->merge_order) {
        for (i = 0; i < n; i++) {
            const lxlsx_reader_range *r = &ws->meta.merges[i];
            if (row >= r->first_row && row <= r->last_row &&
                col >= r->first_col && col <= r->last_col) {
                return (row == r->first_row && col == r->first_col) ? 0 : 1;
            }
        }
        return 0;
    }

    /* Reset the cursor when access is not monotonically increasing in row,
     * so out-of-order callers still get correct results. */
    if (row < ws->merge_last_row) ws->merge_cursor = 0;
    ws->merge_last_row = row;

    /* Drop merges that end above this row — they can never match again. */
    while (ws->merge_cursor < n &&
           ws->meta.merges[ws->merge_order[ws->merge_cursor]].last_row < row) {
        ws->merge_cursor++;
    }

    /* Check the remaining active merges (last_row >= row). */
    for (i = ws->merge_cursor; i < n; i++) {
        const lxlsx_reader_range *r = &ws->meta.merges[ws->merge_order[i]];
        if (row >= r->first_row && row <= r->last_row &&
            col >= r->first_col && col <= r->last_col) {
            /* Inside this range. Master is (first_row, first_col). */
            return (row == r->first_row && col == r->first_col) ? 0 : 1;
        }
    }
    return 0;
}

size_t lxlsx_reader_worksheet_hyperlink_count(const lxlsx_reader_worksheet *ws)
{
    if (!ws) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    return ws->meta.hyperlinks_count;
}

int lxlsx_reader_worksheet_hyperlink_get(const lxlsx_reader_worksheet *ws, size_t idx,
                                lxlsx_reader_hyperlink *out)
{
    const struct lxlsx_reader_hyperlink_owned *h;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    if (idx >= ws->meta.hyperlinks_count) return 0;
    h = &ws->meta.hyperlinks[idx];
    out->range    = h->range;
    out->url      = h->url;
    out->location = h->location;
    out->display  = h->display;
    out->tooltip  = h->tooltip;
    return 1;
}

const char *lxlsx_reader_worksheet_hyperlink_url(const lxlsx_reader_worksheet *ws,
                                        size_t row, size_t col)
{
    size_t i;
    if (!ws) return NULL;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return NULL;
    for (i = 0; i < ws->meta.hyperlinks_count; i++) {
        const struct lxlsx_reader_hyperlink_owned *h = &ws->meta.hyperlinks[i];
        if (row >= h->range.first_row && row <= h->range.last_row &&
            col >= h->range.first_col && col <= h->range.last_col) {
            return h->url ? h->url : h->location;
        }
    }
    return NULL;
}

int lxlsx_reader_worksheet_protection(const lxlsx_reader_worksheet *ws, lxlsx_reader_protection *out)
{
    const lxlsx_reader_worksheet_meta *m;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    m = &ws->meta;
    out->is_present                = m->prot_present;
    {
        size_t n = strlen(m->prot_hash);
        if (n >= sizeof(out->password_hash)) n = sizeof(out->password_hash) - 1;
        memcpy(out->password_hash, m->prot_hash, n);
        out->password_hash[n] = 0;
    }
    out->sheet                = m->prot_sheet;
    out->content              = m->prot_content;
    out->objects              = m->prot_objects;
    out->scenarios            = m->prot_scenarios;
    out->format_cells         = m->prot_format_cells;
    out->format_columns       = m->prot_format_columns;
    out->format_rows          = m->prot_format_rows;
    out->insert_columns       = m->prot_insert_columns;
    out->insert_rows          = m->prot_insert_rows;
    out->insert_hyperlinks    = m->prot_insert_hyperlinks;
    out->delete_columns       = m->prot_delete_columns;
    out->delete_rows          = m->prot_delete_rows;
    out->select_locked_cells  = m->prot_select_locked_cells;
    out->sort                 = m->prot_sort;
    out->auto_filter          = m->prot_auto_filter;
    out->pivot_tables         = m->prot_pivot_tables;
    out->select_unlocked_cells = m->prot_select_unlocked_cells;
    return m->prot_present;
}

int lxlsx_reader_worksheet_row_options(const lxlsx_reader_worksheet *ws, size_t row,
                              lxlsx_reader_row_options *out)
{
    size_t i;
    if (!ws || !out || row == 0) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    for (i = 0; i < ws->meta.rows_count; i++) {
        const struct lxlsx_reader_row_meta *r = &ws->meta.rows[i];
        if (r->row == row) {
            out->has_height    = r->has_height;
            out->height        = r->height;
            out->hidden        = r->hidden;
            out->outline_level = r->outline_level;
            out->collapsed     = r->collapsed;
            out->custom_height = r->custom_height;
            return 1;
        }
    }
    return 0;
}

int lxlsx_reader_worksheet_col_options(const lxlsx_reader_worksheet *ws, size_t col,
                              lxlsx_reader_col_options *out)
{
    size_t i;
    if (!ws || !out || col == 0) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    for (i = 0; i < ws->meta.cols_count; i++) {
        const struct lxlsx_reader_col_meta *c = &ws->meta.cols[i];
        if (col >= c->min && col <= c->max) {
            out->has_width     = c->has_width;
            out->width         = c->width;
            out->hidden        = c->hidden;
            out->outline_level = c->outline_level;
            out->collapsed     = c->collapsed;
            return 1;
        }
    }
    return 0;
}

int lxlsx_reader_worksheet_default_row_height(const lxlsx_reader_worksheet *ws, double *out)
{
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    if (!ws->meta.has_default_row_height) return 0;
    *out = ws->meta.default_row_height;
    return 1;
}

int lxlsx_reader_worksheet_default_col_width(const lxlsx_reader_worksheet *ws, double *out)
{
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    if (!ws->meta.has_default_col_width) return 0;
    *out = ws->meta.default_col_width;
    return 1;
}

/* ---- Data validation accessors ----------------------------------------- */

size_t lxlsx_reader_worksheet_data_validation_count(const lxlsx_reader_worksheet *ws)
{
    if (!ws) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    return ws->meta.dvs_count;
}

int lxlsx_reader_worksheet_data_validation_get(const lxlsx_reader_worksheet *ws, size_t idx,
                                      lxlsx_reader_data_validation *out)
{
    const struct lxlsx_reader_dv_owned *d;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    if (idx >= ws->meta.dvs_count) return 0;
    d = &ws->meta.dvs[idx];
    out->type               = d->type;
    out->operator_          = d->operator_;
    out->error_style        = d->error_style;
    out->formula1           = d->formula1;
    out->formula2           = d->formula2;
    out->prompt             = d->prompt;
    out->prompt_title       = d->prompt_title;
    out->error              = d->error;
    out->error_title        = d->error_title;
    out->sqref              = d->sqref;
    out->allow_blank        = d->allow_blank;
    out->show_drop_down     = d->show_drop_down;
    out->show_input_message = d->show_input_message;
    out->show_error_message = d->show_error_message;
    return 1;
}

/* ---- Conditional format accessors (§8.2.3) ----------------------------- */

size_t lxlsx_reader_worksheet_cf_block_count(const lxlsx_reader_worksheet *ws)
{
    if (!ws) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    return ws->meta.cf_blocks_count;
}

int lxlsx_reader_worksheet_cf_block_get(const lxlsx_reader_worksheet *ws, size_t idx,
                               lxlsx_reader_cf_block *out)
{
    lxlsx_reader_worksheet_meta *m;
    const struct lxlsx_reader_cf_block_owned *bk;
    size_t i;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    m = (lxlsx_reader_worksheet_meta *)&ws->meta;
    if (idx >= m->cf_blocks_count) return 0;
    bk = &m->cf_blocks[idx];

    /* Materialise a public-shape rules array per block, lazily and once.
     * We over-allocate cf_rules_pub to hold ALL rules across all blocks
     * (positions are sequential within the array per block). */
    if (!m->cf_rules_pub) {
        size_t total = 0, j, off = 0;
        for (j = 0; j < m->cf_blocks_count; j++) total += m->cf_blocks[j].rules_count;
        m->cf_rules_pub = (lxlsx_reader_cf_rule *)calloc(total ? total : 1, sizeof(lxlsx_reader_cf_rule));
        if (!m->cf_rules_pub) return 0;
        m->cf_rules_pub_n = total;
        for (j = 0; j < m->cf_blocks_count; j++) {
            for (i = 0; i < m->cf_blocks[j].rules_count; i++, off++) {
                const struct lxlsx_reader_cf_rule_owned *src = &m->cf_blocks[j].rules[i];
                lxlsx_reader_cf_rule                    *dst = &m->cf_rules_pub[off];
                dst->type        = src->type;
                dst->operator_   = src->operator_;
                dst->priority    = src->priority;
                dst->stop_if_true = src->stop_if_true;
                dst->dxf_id      = src->dxf_id;
                dst->percent     = src->percent;
                dst->bottom      = src->bottom;
                dst->rank        = src->rank;
                dst->text        = src->text;
                dst->time_period = src->time_period;
                dst->formula1    = src->formula1;
                dst->formula2    = src->formula2;
            }
        }
    }

    /* Find the offset of this block's rules in the contiguous array. */
    {
        size_t j, off = 0;
        for (j = 0; j < idx; j++) off += m->cf_blocks[j].rules_count;
        out->sqref       = bk->sqref;
        out->rules       = bk->rules_count > 0 ? &m->cf_rules_pub[off] : NULL;
        out->rules_count = bk->rules_count;
    }
    return 1;
}

/* ---- Rich-text runs accessor (§8.2.2) ---------------------------------- */

size_t lxlsx_reader_cell_string_runs(const lxlsx_reader_worksheet *ws, const lxlsx_cell *c,
                            lxlsx_reader_string_run *out, size_t cap)
{
    if (!ws || !c) return 0;

    if (c->type == STRING_CELL) {
        /* SST cell: cell.raw holds the SST index as decimal text. */
        uint32_t idx;
        size_t   count = 0;
        const lxlsx_reader_sst_run *runs;
        size_t i;
        if (!ws->wb || !ws->wb->sst || !c->raw.ptr) return 0;
        idx = (uint32_t)strtoul(c->raw.ptr, NULL, 10);
        runs = lxlsx_reader_sst_get_runs(ws->wb->sst, idx, &count);
        if (count == 0) return 0;
        if (out) {
            for (i = 0; i < count && i < cap; i++) {
                out[i].text       = runs[i].text;
                out[i].text_len   = runs[i].text ? strlen(runs[i].text) : 0;
                out[i].font_name  = runs[i].font_name;
                out[i].font_size  = runs[i].font_size;
                out[i].bold       = runs[i].bold;
                out[i].italic     = runs[i].italic;
                out[i].strike     = runs[i].strike;
                out[i].underline  = runs[i].underline;
                out[i].color      = runs[i].color;
            }
        }
        return count;
    }

    if (c->type == INLINE_STRING_CELL) {
        size_t i, n = ws->inline_runs_count;
        if (n == 0) return 0;
        if (out) {
            for (i = 0; i < n && i < cap; i++) {
                out[i].text       = ws->inline_runs[i].text;
                out[i].text_len   = ws->inline_runs[i].text ? strlen(ws->inline_runs[i].text) : 0;
                out[i].font_name  = ws->inline_runs[i].font_name;
                out[i].font_size  = ws->inline_runs[i].font_size;
                out[i].bold       = ws->inline_runs[i].bold;
                out[i].italic     = ws->inline_runs[i].italic;
                out[i].strike     = ws->inline_runs[i].strike;
                out[i].underline  = ws->inline_runs[i].underline;
                out[i].color      = ws->inline_runs[i].color;
            }
        }
        return n;
    }

    return 0;
}

/* ---- Page setup accessor ------------------------------------------------ */

int lxlsx_reader_worksheet_page_setup(const lxlsx_reader_worksheet *ws, lxlsx_reader_page_setup *out)
{
    const lxlsx_reader_worksheet_meta *m;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    m = &ws->meta;
    /* If nothing about the page was parsed, surface 0 so callers can return
     * an explicit "no page setup" sentinel. */
    if (!m->page_has_margins && !m->page_has_setup &&
        !m->page_print_h_centered && !m->page_print_v_centered &&
        !m->page_print_grid_lines && !m->page_print_headings &&
        !m->page_odd_header && !m->page_odd_footer &&
        !m->page_even_header && !m->page_even_footer &&
        !m->page_first_header && !m->page_first_footer) return 0;

    memset(out, 0, sizeof(*out));
    out->has_margins   = m->page_has_margins;
    out->margin_left   = m->page_margin_left;
    out->margin_right  = m->page_margin_right;
    out->margin_top    = m->page_margin_top;
    out->margin_bottom = m->page_margin_bottom;
    out->margin_header = m->page_margin_header;
    out->margin_footer = m->page_margin_footer;

    out->has_setup           = m->page_has_setup;
    out->paper_size          = m->page_paper_size;
    out->fit_to_width        = m->page_fit_to_width;
    out->fit_to_height       = m->page_fit_to_height;
    out->scale               = m->page_scale;
    out->orientation_landscape = m->page_orientation_landscape;
    out->horizontal_dpi      = m->page_horizontal_dpi;
    out->vertical_dpi        = m->page_vertical_dpi;
    out->first_page_number   = m->page_first_page_number;
    out->use_first_page_number = m->page_use_first_page_number;

    out->print_horizontal_centered = m->page_print_h_centered;
    out->print_vertical_centered   = m->page_print_v_centered;
    out->print_grid_lines          = m->page_print_grid_lines;
    out->print_headings            = m->page_print_headings;

    out->odd_header   = m->page_odd_header;
    out->odd_footer   = m->page_odd_footer;
    out->even_header  = m->page_even_header;
    out->even_footer  = m->page_even_footer;
    out->first_header = m->page_first_header;
    out->first_footer = m->page_first_footer;
    out->different_odd_even = m->page_different_odd_even;
    out->different_first    = m->page_different_first;
    out->scale_with_doc     = m->page_scale_with_doc;
    out->align_with_margins = m->page_align_with_margins;
    return 1;
}

/* ---- AutoFilter accessor ------------------------------------------------ */

int lxlsx_reader_worksheet_autofilter(const lxlsx_reader_worksheet *ws, lxlsx_reader_autofilter *out)
{
    lxlsx_reader_worksheet_meta *m;
    size_t i;
    if (!ws || !out) return 0;
    if (lxlsx_reader_worksheet_ensure_meta(ws) != LXLSX_READER_NO_ERROR) return 0;
    m = (lxlsx_reader_worksheet_meta *)&ws->meta;  /* mutable for cache materialisation */
    if (!m->autofilter_present) return 0;

    /* Lazily materialise a stable public-shape array of filter columns so
     * out->columns has a predictable lifetime (same as the worksheet). */
    if (!m->filter_columns_pub && m->filter_columns_count > 0) {
        m->filter_columns_pub = (lxlsx_reader_filter_column *)
            calloc(m->filter_columns_count, sizeof(lxlsx_reader_filter_column));
        if (m->filter_columns_pub) {
            for (i = 0; i < m->filter_columns_count; i++) {
                const struct lxlsx_reader_filter_column_owned *src = &m->filter_columns[i];
                lxlsx_reader_filter_column *dst = &m->filter_columns_pub[i];
                dst->col_id        = src->col_id;
                dst->kind          = (lxlsx_reader_filter_kind)src->kind;
                dst->values        = (const char **)src->values;
                dst->custom_and    = src->custom_and;
                dst->custom_op_1   = src->custom_op_1;
                dst->custom_val_1  = src->custom_val_1;
                dst->custom_op_2   = src->custom_op_2;
                dst->custom_val_2  = src->custom_val_2;
                dst->top           = src->top;
                dst->percent       = src->percent;
                dst->top_value     = src->top_value;
            }
        }
    }

    out->range         = m->autofilter_range;
    out->columns       = m->filter_columns_pub;
    out->columns_count = m->filter_columns_count;
    return 1;
}
