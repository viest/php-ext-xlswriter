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

#include "lxlsx/xmlwriter.h"
#include "lxlsx/worksheet.h"
#include "lxlsx/format.h"
#include "lxlsx/utility.h"

#ifdef USE_OPENSSL_MD5
#include <openssl/md5.h>
#else
#ifndef USE_NO_MD5
#include "lxlsx/third_party/md5.h"
#endif
#endif

#define LXW_STR_MAX                      32767
#define LXW_BUFFER_SIZE                  4096
#define LXW_PRINT_ACROSS                 1
#define LXW_VALIDATION_MAX_TITLE_LENGTH  32
#define LXW_VALIDATION_MAX_STRING_LENGTH 255
#define LXW_THIS_ROW "[#This Row],"
/*
 * Forward declarations.
 */
STATIC void _worksheet_write_rows(lxw_worksheet *self);
STATIC int _row_cmp(lxw_row *row1, lxw_row *row2);
STATIC int _cell_cmp(lxw_cell *cell1, lxw_cell *cell2);
STATIC int _drawing_rel_id_cmp(lxw_drawing_rel_id *tuple1,
                               lxw_drawing_rel_id *tuple2);
STATIC int _cond_format_hash_cmp(lxw_cond_format_hash_element *elem_1,
                                 lxw_cond_format_hash_element *elem_2);

#ifndef __clang_analyzer__
LXW_RB_GENERATE_ROW(lxw_table_rows, lxw_row, tree_pointers, _row_cmp);
LXW_RB_GENERATE_CELL(lxw_table_cells, lxw_cell, tree_pointers, _cell_cmp);
LXW_RB_GENERATE_DRAWING_REL_IDS(lxw_drawing_rel_ids, lxw_drawing_rel_id,
                                tree_pointers, _drawing_rel_id_cmp);
LXW_RB_GENERATE_VML_DRAWING_REL_IDS(lxw_vml_drawing_rel_ids,
                                    lxw_drawing_rel_id, tree_pointers,
                                    _drawing_rel_id_cmp);
LXW_RB_GENERATE_COND_FORMAT_HASH(lxw_cond_format_hash,
                                 lxw_cond_format_hash_element, tree_pointers,
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
lxw_row *
lxw_worksheet_find_row(lxw_worksheet *self, lxw_row_t row_num)
{
    lxw_row tmp_row;

    tmp_row.row_num = row_num;

    return RB_FIND(lxw_table_rows, self->table, &tmp_row);
}

/*
 * Find but don't create a cell object for a given row object and col number.
 */
lxw_cell *
lxw_worksheet_find_cell_in_row(lxw_row *row, lxw_col_t col_num)
{
    lxw_cell tmp_cell;

    if (!row)
        return NULL;

    tmp_cell.col_num = col_num;

    return RB_FIND(lxw_table_cells, row->cells, &tmp_cell);
}

/*
 * Create a new worksheet object.
 */
lxw_worksheet *
lxw_worksheet_new(lxw_worksheet_init_data *init_data)
{
    lxw_worksheet *worksheet = calloc(1, sizeof(lxw_worksheet));
    GOTO_LABEL_ON_MEM_ERROR(worksheet, mem_error);

    worksheet->table = calloc(1, sizeof(struct lxw_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->table, mem_error);
    RB_INIT(worksheet->table);

    worksheet->hyperlinks = calloc(1, sizeof(struct lxw_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->hyperlinks, mem_error);
    RB_INIT(worksheet->hyperlinks);

    worksheet->comments = calloc(1, sizeof(struct lxw_table_rows));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->comments, mem_error);
    RB_INIT(worksheet->comments);

    /* Initialize the cached rows. */
    worksheet->table->cached_row_num = LXW_ROW_MAX + 1;
    worksheet->hyperlinks->cached_row_num = LXW_ROW_MAX + 1;
    worksheet->comments->cached_row_num = LXW_ROW_MAX + 1;

    if (init_data && init_data->optimize) {
        worksheet->array = calloc(LXW_COL_MAX, sizeof(struct lxw_cell *));
        GOTO_LABEL_ON_MEM_ERROR(worksheet->array, mem_error);
    }

    worksheet->col_options =
        calloc(LXW_COL_META_MAX, sizeof(lxw_col_options *));
    worksheet->col_options_max = LXW_COL_META_MAX;
    GOTO_LABEL_ON_MEM_ERROR(worksheet->col_options, mem_error);

    worksheet->col_formats = calloc(LXW_COL_META_MAX, sizeof(lxw_format *));
    worksheet->col_formats_max = LXW_COL_META_MAX;
    GOTO_LABEL_ON_MEM_ERROR(worksheet->col_formats, mem_error);

    worksheet->optimize_row = calloc(1, sizeof(struct lxw_row));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_row, mem_error);
    worksheet->optimize_row->height = LXW_DEF_ROW_HEIGHT;

    worksheet->merged_ranges = calloc(1, sizeof(struct lxw_merged_ranges));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->merged_ranges, mem_error);
    STAILQ_INIT(worksheet->merged_ranges);

    worksheet->image_props = calloc(1, sizeof(struct lxw_image_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->image_props, mem_error);
    STAILQ_INIT(worksheet->image_props);

    worksheet->embedded_image_props =
        calloc(1, sizeof(struct lxw_embedded_image_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->embedded_image_props, mem_error);
    STAILQ_INIT(worksheet->embedded_image_props);

    worksheet->chart_data = calloc(1, sizeof(struct lxw_chart_props));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->chart_data, mem_error);
    STAILQ_INIT(worksheet->chart_data);

    worksheet->comment_objs = calloc(1, sizeof(struct lxw_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->comment_objs, mem_error);
    STAILQ_INIT(worksheet->comment_objs);

    worksheet->header_image_objs = calloc(1, sizeof(struct lxw_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->header_image_objs, mem_error);
    STAILQ_INIT(worksheet->header_image_objs);

    worksheet->button_objs = calloc(1, sizeof(struct lxw_comment_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->button_objs, mem_error);
    STAILQ_INIT(worksheet->button_objs);

    worksheet->selections = calloc(1, sizeof(struct lxw_selections));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->selections, mem_error);
    STAILQ_INIT(worksheet->selections);

    worksheet->data_validations =
        calloc(1, sizeof(struct lxw_data_validations));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->data_validations, mem_error);
    STAILQ_INIT(worksheet->data_validations);

    worksheet->table_objs = calloc(1, sizeof(struct lxw_table_objs));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->table_objs, mem_error);
    STAILQ_INIT(worksheet->table_objs);

    worksheet->external_hyperlinks = calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_hyperlinks, mem_error);
    STAILQ_INIT(worksheet->external_hyperlinks);

    worksheet->external_drawing_links =
        calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_drawing_links, mem_error);
    STAILQ_INIT(worksheet->external_drawing_links);

    worksheet->drawing_links = calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->drawing_links, mem_error);
    STAILQ_INIT(worksheet->drawing_links);

    worksheet->vml_drawing_links = calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->vml_drawing_links, mem_error);
    STAILQ_INIT(worksheet->vml_drawing_links);

    worksheet->external_table_links =
        calloc(1, sizeof(struct lxw_rel_tuples));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->external_table_links, mem_error);
    STAILQ_INIT(worksheet->external_table_links);

    if (init_data && init_data->optimize) {
        FILE *tmpfile;

        worksheet->optimize_buffer = NULL;
        worksheet->optimize_buffer_size = 0;
        tmpfile = lxw_get_filehandle(&worksheet->optimize_buffer,
                                     &worksheet->optimize_buffer_size,
                                     init_data->tmpdir);
        if (!tmpfile) {
            LXW_ERROR("Error creating tmpfile() for worksheet in "
                      "'constant_memory' mode.");
            goto mem_error;
        }

        worksheet->optimize_tmpfile = tmpfile;
        GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_tmpfile, mem_error);
        worksheet->file = worksheet->optimize_tmpfile;
    }

    worksheet->drawing_rel_ids =
        calloc(1, sizeof(struct lxw_drawing_rel_ids));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->drawing_rel_ids, mem_error);
    RB_INIT(worksheet->drawing_rel_ids);

    worksheet->vml_drawing_rel_ids =
        calloc(1, sizeof(struct lxw_vml_drawing_rel_ids));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->vml_drawing_rel_ids, mem_error);
    RB_INIT(worksheet->vml_drawing_rel_ids);

    worksheet->conditional_formats =
        calloc(1, sizeof(struct lxw_cond_format_hash));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->conditional_formats, mem_error);
    RB_INIT(worksheet->conditional_formats);

    /* Initialize the worksheet dimensions. */
    worksheet->dim_rowmax = 0;
    worksheet->dim_colmax = 0;
    worksheet->dim_rowmin = LXW_ROW_MAX;
    worksheet->dim_colmin = LXW_COL_MAX;

    worksheet->default_row_height = LXW_DEF_ROW_HEIGHT;
    worksheet->default_row_pixels = 20;
    worksheet->default_col_pixels = 64;

    /* Initialize the page setup properties. */
    worksheet->fit_height = 0;
    worksheet->fit_width = 0;
    worksheet->page_start = 0;
    worksheet->print_scale = 100;
    worksheet->fit_page = 0;
    worksheet->orientation = LXW_TRUE;
    worksheet->page_order = 0;
    worksheet->page_setup_changed = LXW_FALSE;
    worksheet->page_view = LXW_FALSE;
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
    worksheet->zoom_scale_normal = LXW_TRUE;
    worksheet->show_zeros = LXW_TRUE;
    worksheet->outline_on = LXW_TRUE;
    worksheet->outline_style = LXW_TRUE;
    worksheet->outline_below = LXW_TRUE;
    worksheet->outline_right = LXW_FALSE;
    worksheet->tab_color = LXW_COLOR_UNSET;
    worksheet->max_url_length = 2079;
    worksheet->comment_display_default = LXW_COMMENT_DISPLAY_HIDDEN;

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
    lxw_worksheet_free(worksheet);
    return NULL;
}

/*
 * Free vml object.
 */
STATIC void
_free_vml_object(lxw_vml_obj *vml_obj)
{
    if (!vml_obj)
        return;

    free(vml_obj->author);
    free(vml_obj->font_name);
    free(vml_obj->text);
    free(vml_obj->image_position);
    free(vml_obj->name);
    free(vml_obj->macro);

    free(vml_obj);
}

/*
 * Free autofilter rule object.
 */
STATIC void
_free_filter_rule(lxw_filter_rule_obj *rule_obj)
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
_free_filter_rules(lxw_worksheet *worksheet)
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
_free_cell(lxw_cell *cell)
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
_free_row(lxw_row *row)
{
    lxw_cell *cell;
    lxw_cell *next_cell;

    if (!row)
        return;

    for (cell = RB_MIN(lxw_table_cells, row->cells); cell; cell = next_cell) {
        next_cell = RB_NEXT(lxw_table_cells, row->cells, cell);
        RB_REMOVE(lxw_table_cells, row->cells, cell);
        _free_cell(cell);
    }

    free(row->cells);
    free(row);
}

/*
 * Free a worksheet image_options.
 */
STATIC void
_free_object_properties(lxw_object_properties *object_property)
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
_free_data_validation(lxw_data_val_obj *data_validation)
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
_free_cond_format(lxw_cond_format_obj *cond_format)
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
_free_relationship(lxw_rel_tuple *relationship)
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
_free_worksheet_table_column(lxw_table_column *column)
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
_free_worksheet_table(lxw_table_obj *table)
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
lxw_worksheet_free(lxw_worksheet *worksheet)
{
    lxw_row *row;
    lxw_row *next_row;
    lxw_col_t col;
    lxw_merged_range *merged_range;
    lxw_object_properties *object_props;
    lxw_vml_obj *vml_obj;
    lxw_selection *selection;
    lxw_data_val_obj *data_validation;
    lxw_rel_tuple *relationship;
    lxw_cond_format_obj *cond_format;
    lxw_table_obj *table_obj;
    struct lxw_drawing_rel_id *drawing_rel_id;
    struct lxw_drawing_rel_id *next_drawing_rel_id;
    struct lxw_cond_format_hash_element *cond_format_elem;
    struct lxw_cond_format_hash_element *next_cond_format_elem;

    if (!worksheet)
        return;

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
        for (row = RB_MIN(lxw_table_rows, worksheet->table); row;
             row = next_row) {

            next_row = RB_NEXT(lxw_table_rows, worksheet->table, row);
            RB_REMOVE(lxw_table_rows, worksheet->table, row);
            _free_row(row);
        }

        free(worksheet->table);
    }

    if (worksheet->hyperlinks) {
        for (row = RB_MIN(lxw_table_rows, worksheet->hyperlinks); row;
             row = next_row) {

            next_row = RB_NEXT(lxw_table_rows, worksheet->hyperlinks, row);
            RB_REMOVE(lxw_table_rows, worksheet->hyperlinks, row);
            _free_row(row);
        }

        free(worksheet->hyperlinks);
    }

    if (worksheet->comments) {
        for (row = RB_MIN(lxw_table_rows, worksheet->comments); row;
             row = next_row) {

            next_row = RB_NEXT(lxw_table_rows, worksheet->comments, row);
            RB_REMOVE(lxw_table_rows, worksheet->comments, row);
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

    if (worksheet->chart_data) {
        while (!STAILQ_EMPTY(worksheet->chart_data)) {
            object_props = STAILQ_FIRST(worksheet->chart_data);
            STAILQ_REMOVE_HEAD(worksheet->chart_data, list_pointers);
            _free_object_properties(object_props);
        }

        free(worksheet->chart_data);
    }

    /* Just free the list. The list objects are freed from the RB tree. */
    free(worksheet->comment_objs);

    if (worksheet->header_image_objs) {
        while (!STAILQ_EMPTY(worksheet->header_image_objs)) {
            vml_obj = STAILQ_FIRST(worksheet->header_image_objs);
            STAILQ_REMOVE_HEAD(worksheet->header_image_objs, list_pointers);
            _free_vml_object(vml_obj);
        }

        free(worksheet->header_image_objs);
    }

    if (worksheet->button_objs) {
        while (!STAILQ_EMPTY(worksheet->button_objs)) {
            vml_obj = STAILQ_FIRST(worksheet->button_objs);
            STAILQ_REMOVE_HEAD(worksheet->button_objs, list_pointers);
            _free_vml_object(vml_obj);
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

    if (worksheet->table_objs) {
        while (!STAILQ_EMPTY(worksheet->table_objs)) {
            table_obj = STAILQ_FIRST(worksheet->table_objs);
            STAILQ_REMOVE_HEAD(worksheet->table_objs, list_pointers);
            _free_worksheet_table(table_obj);
        }

        free(worksheet->table_objs);
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

    while (!STAILQ_EMPTY(worksheet->drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->drawing_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->drawing_links);

    while (!STAILQ_EMPTY(worksheet->vml_drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->vml_drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->vml_drawing_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->vml_drawing_links);

    while (!STAILQ_EMPTY(worksheet->external_table_links)) {
        relationship = STAILQ_FIRST(worksheet->external_table_links);
        STAILQ_REMOVE_HEAD(worksheet->external_table_links, list_pointers);
        _free_relationship(relationship);
    }
    free(worksheet->external_table_links);

    if (worksheet->drawing_rel_ids) {
        for (drawing_rel_id =
             RB_MIN(lxw_drawing_rel_ids, worksheet->drawing_rel_ids);
             drawing_rel_id; drawing_rel_id = next_drawing_rel_id) {

            next_drawing_rel_id =
                RB_NEXT(lxw_drawing_rel_ids, worksheet->drawing_rel_id,
                        drawing_rel_id);
            RB_REMOVE(lxw_drawing_rel_ids, worksheet->drawing_rel_ids,
                      drawing_rel_id);
            free(drawing_rel_id->target);
            free(drawing_rel_id);
        }

        free(worksheet->drawing_rel_ids);
    }

    if (worksheet->vml_drawing_rel_ids) {
        for (drawing_rel_id =
             RB_MIN(lxw_vml_drawing_rel_ids, worksheet->vml_drawing_rel_ids);
             drawing_rel_id; drawing_rel_id = next_drawing_rel_id) {

            next_drawing_rel_id =
                RB_NEXT(lxw_vml_drawing_rel_ids, worksheet->drawing_rel_id,
                        drawing_rel_id);
            RB_REMOVE(lxw_vml_drawing_rel_ids, worksheet->vml_drawing_rel_ids,
                      drawing_rel_id);
            free(drawing_rel_id->target);
            free(drawing_rel_id);
        }

        free(worksheet->vml_drawing_rel_ids);
    }

    if (worksheet->conditional_formats) {
        for (cond_format_elem =
             RB_MIN(lxw_cond_format_hash, worksheet->conditional_formats);
             cond_format_elem; cond_format_elem = next_cond_format_elem) {

            next_cond_format_elem = RB_NEXT(lxw_cond_format_hash,
                                            worksheet->conditional_formats,
                                            cond_format_elem);
            RB_REMOVE(lxw_cond_format_hash,
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
        for (col = 0; col < LXW_COL_MAX; col++) {
            _free_cell(worksheet->array[col]);
        }
        free(worksheet->array);
    }

    if (worksheet->optimize_row)
        free(worksheet->optimize_row);

    if (worksheet->drawing)
        lxw_drawing_free(worksheet->drawing);

    free(worksheet->hbreaks);
    free(worksheet->vbreaks);
    free((void *) worksheet->name);
    free((void *) worksheet->quoted_name);
    free(worksheet->vba_codename);
    free(worksheet->vml_data_id_str);
    free(worksheet->vml_header_id_str);
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
STATIC lxw_row *
_new_row(lxw_row_t row_num)
{
    lxw_row *row = calloc(1, sizeof(lxw_row));

    if (row) {
        row->row_num = row_num;
        row->cells = calloc(1, sizeof(struct lxw_table_cells));
        row->height = LXW_DEF_ROW_HEIGHT;

        if (row->cells)
            RB_INIT(row->cells);
        else
            LXW_MEM_ERROR();
    }
    else {
        LXW_MEM_ERROR();
    }

    return row;
}

/*
 * Create a new worksheet number cell object.
 */
STATIC lxw_cell *
_new_number_cell(lxw_row_t row_num,
                 lxw_col_t col_num, double value, lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_string_cell(lxw_row_t row_num,
                 lxw_col_t col_num, int32_t string_id, char *sst_string,
                 lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = STRING_CELL;
    cell->format = format;
    cell->u.string_id = string_id;
    cell->sst_string = sst_string;

    return cell;
}

/*
 * Create a new worksheet inline_string cell object.
 */
STATIC lxw_cell *
_new_inline_string_cell(lxw_row_t row_num,
                        lxw_col_t col_num, char *string, lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_inline_rich_string_cell(lxw_row_t row_num,
                             lxw_col_t col_num, const char *string,
                             lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_formula_cell(lxw_row_t row_num,
                  lxw_col_t col_num, char *formula, lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_array_formula_cell(lxw_row_t row_num, lxw_col_t col_num, char *formula,
                        char *range, lxw_format *format, uint8_t is_dynamic)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_blank_cell(lxw_row_t row_num, lxw_col_t col_num, lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_boolean_cell(lxw_row_t row_num, lxw_col_t col_num, int value,
                  lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_error_cell(lxw_row_t row_num, lxw_col_t col_num, uint32_t value,
                lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_comment_cell(lxw_row_t row_num, lxw_col_t col_num,
                  lxw_vml_obj *comment_obj)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_cell *
_new_hyperlink_cell(lxw_row_t row_num, lxw_col_t col_num,
                    enum cell_types link_type, char *url, char *string,
                    char *tooltip)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
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
STATIC lxw_row *
_get_row_list(struct lxw_table_rows *table, lxw_row_t row_num)
{
    lxw_row *row;
    lxw_row *existing_row;

    if (table->cached_row_num == row_num)
        return table->cached_row;

    /* Create a new row and try and insert it. */
    row = _new_row(row_num);
    existing_row = RB_INSERT(lxw_table_rows, table, row);

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
STATIC lxw_row *
_get_row(lxw_worksheet *self, lxw_row_t row_num)
{
    lxw_row *row;

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
            lxw_worksheet_write_single_row(self);
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
_insert_cell_list(struct lxw_table_cells *cell_list,
                  lxw_cell *cell, lxw_col_t col_num)
{
    lxw_cell *existing_cell;

    cell->col_num = col_num;

    existing_cell = RB_INSERT(lxw_table_cells, cell_list, cell);

    /* If existing_cell is not NULL, then that cell already existed. */
    /* Remove existing_cell and add new one in again. */
    if (existing_cell) {
        RB_REMOVE(lxw_table_cells, cell_list, existing_cell);

        /* Add it in again. */
        RB_INSERT(lxw_table_cells, cell_list, cell);
        _free_cell(existing_cell);
    }

    return;
}

/*
 * Insert a cell object into the cell list or array.
 */
STATIC void
_insert_cell(lxw_worksheet *self, lxw_row_t row_num, lxw_col_t col_num,
             lxw_cell *cell)
{
    lxw_row *row = _get_row(self, row_num);

    if (!self->optimize) {
        row->data_changed = LXW_TRUE;
        _insert_cell_list(row->cells, cell, col_num);
    }
    else {
        if (row) {
            row->data_changed = LXW_TRUE;

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
_insert_cell_placeholder(lxw_worksheet *self, lxw_row_t row_num,
                         lxw_col_t col_num)
{
    lxw_row *row;
    lxw_cell *cell;

    /* The spans calculation isn't required in constant_memory mode. */
    if (self->optimize)
        return;

    cell = _new_blank_cell(row_num, col_num, NULL);
    if (!cell)
        return;

    /* Only add a cell if one doesn't already exist. */
    row = _get_row(self, row_num);
    if (!RB_FIND(lxw_table_cells, row->cells, cell)) {
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
_insert_hyperlink(lxw_worksheet *self, lxw_row_t row_num, lxw_col_t col_num,
                  lxw_cell *link)
{
    lxw_row *row = _get_row_list(self->hyperlinks, row_num);

    _insert_cell_list(row->cells, link, col_num);
}

/*
 * Insert a comment into the comment RB tree.
 */
STATIC void
_insert_comment(lxw_worksheet *self, lxw_row_t row_num, lxw_col_t col_num,
                lxw_cell *link)
{
    lxw_row *row = _get_row_list(self->comments, row_num);

    _insert_cell_list(row->cells, link, col_num);
}

/*
 * Next power of two for column reallocs. Taken from bithacks in the public
 * domain.
 */
STATIC lxw_col_t
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
STATIC lxw_error
_check_dimensions(lxw_worksheet *self,
                  lxw_row_t row_num,
                  lxw_col_t col_num, int8_t ignore_row, int8_t ignore_col)
{
    if (row_num >= LXW_ROW_MAX)
        return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    if (col_num >= LXW_COL_MAX)
        return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    /* In optimization mode we don't change dimensions for rows that are */
    /* already written. */
    if (!ignore_row && !ignore_col && self->optimize) {
        if (row_num < self->optimize_row->row_num)
            return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
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

    return LXW_NO_ERROR;
}

/*
 * Comparator for the row structure red/black tree.
 */
STATIC int
_row_cmp(lxw_row *row1, lxw_row *row2)
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
_cell_cmp(lxw_cell *cell1, lxw_cell *cell2)
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
_drawing_rel_id_cmp(lxw_drawing_rel_id *rel_id1, lxw_drawing_rel_id *rel_id2)
{
    return strcmp(rel_id1->target, rel_id2->target);
}

/*
 * Comparator for the conditional format RB hash elements.
 */
STATIC int
_cond_format_hash_cmp(lxw_cond_format_hash_element *elem_1,
                      lxw_cond_format_hash_element *elem_2)
{
    return strcmp(elem_1->sqref, elem_2->sqref);
}

/*
 * Get the index used to address a drawing rel link.
 */
STATIC uint32_t
_get_drawing_rel_index(lxw_worksheet *self, char *target)
{
    lxw_drawing_rel_id tmp_drawing_rel_id;
    lxw_drawing_rel_id *found_duplicate_target = NULL;
    lxw_drawing_rel_id *new_drawing_rel_id = NULL;

    if (target) {
        tmp_drawing_rel_id.target = target;
        found_duplicate_target = RB_FIND(lxw_drawing_rel_ids,
                                         self->drawing_rel_ids,
                                         &tmp_drawing_rel_id);
    }

    if (found_duplicate_target) {
        return found_duplicate_target->id;
    }
    else {
        self->drawing_rel_id++;

        if (target) {
            new_drawing_rel_id = calloc(1, sizeof(lxw_drawing_rel_id));

            if (new_drawing_rel_id) {
                new_drawing_rel_id->id = self->drawing_rel_id;
                new_drawing_rel_id->target = lxw_strdup(target);

                RB_INSERT(lxw_drawing_rel_ids, self->drawing_rel_ids,
                          new_drawing_rel_id);
            }
        }

        return self->drawing_rel_id;
    }
}

/*
 * find the index used to address a drawing rel link.
 */
STATIC uint32_t
_find_drawing_rel_index(lxw_worksheet *self, char *target)
{
    lxw_drawing_rel_id tmp_drawing_rel_id;
    lxw_drawing_rel_id *found_duplicate_target = NULL;

    if (!target)
        return 0;

    tmp_drawing_rel_id.target = target;
    found_duplicate_target = RB_FIND(lxw_drawing_rel_ids,
                                     self->drawing_rel_ids,
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
_get_vml_drawing_rel_index(lxw_worksheet *self, char *target)
{
    lxw_drawing_rel_id tmp_drawing_rel_id;
    lxw_drawing_rel_id *found_duplicate_target = NULL;
    lxw_drawing_rel_id *new_drawing_rel_id = NULL;

    if (target) {
        tmp_drawing_rel_id.target = target;
        found_duplicate_target = RB_FIND(lxw_vml_drawing_rel_ids,
                                         self->vml_drawing_rel_ids,
                                         &tmp_drawing_rel_id);
    }

    if (found_duplicate_target) {
        return found_duplicate_target->id;
    }
    else {
        self->vml_drawing_rel_id++;

        if (target) {
            new_drawing_rel_id = calloc(1, sizeof(lxw_drawing_rel_id));

            if (new_drawing_rel_id) {
                new_drawing_rel_id->id = self->vml_drawing_rel_id;
                new_drawing_rel_id->target = lxw_strdup(target);

                RB_INSERT(lxw_vml_drawing_rel_ids, self->vml_drawing_rel_ids,
                          new_drawing_rel_id);
            }
        }

        return self->vml_drawing_rel_id;
    }
}

/*
 * find the index used to address a VML drawing rel link.
 */
STATIC uint32_t
_find_vml_drawing_rel_index(lxw_worksheet *self, char *target)
{
    lxw_drawing_rel_id tmp_drawing_rel_id;
    lxw_drawing_rel_id *found_duplicate_target = NULL;

    if (!target)
        return 0;

    tmp_drawing_rel_id.target = target;
    found_duplicate_target = RB_FIND(lxw_vml_drawing_rel_ids,
                                     self->vml_drawing_rel_ids,
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
lxw_basename(const char *path)
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

    while (list[i] && length < LXW_VALIDATION_MAX_STRING_LENGTH) {
        /* Include commas in the length. */
        length += 1 + lxw_utf8_strlen(list[i]);
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
    str = calloc(1, LXW_VALIDATION_MAX_STRING_LENGTH * 4 + 3);
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

    if (pixels == LXW_DEF_COL_WIDTH_PIXELS)
        width = LXW_DEF_COL_WIDTH;
    else if (pixels <= 12.0)
        width = pixels / (max_digit_width + padding);
    else
        width = (pixels - padding) / max_digit_width;

    return width;
}

STATIC double
_pixels_to_height(double pixels)
{
    if (pixels == LXW_DEF_ROW_HEIGHT_PIXELS)
        return LXW_DEF_ROW_HEIGHT;
    else
        return pixels * 0.75;
}

/* Check and set if an autofilter is a standard or custom filter. */
void
_set_custom_filter(lxw_filter_rule_obj *rule_obj)
{
    rule_obj->is_custom = LXW_TRUE;

    if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_EQUAL_TO)
        rule_obj->is_custom = LXW_FALSE;

    if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
        rule_obj->is_custom = LXW_FALSE;

    if (rule_obj->criteria2 != LXW_FILTER_CRITERIA_NONE) {
        if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_EQUAL_TO)
            rule_obj->is_custom = LXW_FALSE;

        if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
            rule_obj->is_custom = LXW_FALSE;

        if (rule_obj->type == LXW_FILTER_TYPE_AND)
            rule_obj->is_custom = LXW_TRUE;
    }

    if (rule_obj->value1_string && strpbrk(rule_obj->value1_string, "*?"))
        rule_obj->is_custom = LXW_TRUE;

    if (rule_obj->value2_string && strpbrk(rule_obj->value2_string, "*?"))
        rule_obj->is_custom = LXW_TRUE;
}

/* Check and copy user input for table styles in worksheet_add_table(). */
void
_check_and_copy_table_style(lxw_table_obj *table_obj,
                            lxw_table_options *user_options)
{
    if (!user_options)
        return;

    /* Set the defaults. */
    table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
    table_obj->style_type_number = 9;

    if (user_options->style_type > LXW_TABLE_STYLE_TYPE_DARK) {
        LXW_WARN_FORMAT1
            ("worksheet_add_table(): invalid style_type = %d. "
             "Using default TableStyleMedium9", user_options->style_type);

        table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
        table_obj->style_type_number = 9;
    }
    else {
        table_obj->style_type = user_options->style_type;
    }

    /* Each type (light, medium and dark) has a different number of styles. */
    if (user_options->style_type == LXW_TABLE_STYLE_TYPE_LIGHT) {
        if (user_options->style_type_number > 21) {
            LXW_WARN_FORMAT1("worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXW_TABLE_STYLE_TYPE_LIGHT. "
                             "Using default TableStyleMedium9",
                             user_options->style_type);

            table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
            table_obj->style_type_number = 9;
        }
        else {
            table_obj->style_type_number = user_options->style_type_number;
        }
    }

    if (user_options->style_type == LXW_TABLE_STYLE_TYPE_MEDIUM) {
        if (user_options->style_type_number < 1
            || user_options->style_type_number > 28) {
            LXW_WARN_FORMAT1("worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXW_TABLE_STYLE_TYPE_MEDIUM. "
                             "Using default TableStyleMedium9",
                             user_options->style_type_number);

            table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
            table_obj->style_type_number = 9;
        }
        else {
            table_obj->style_type_number = user_options->style_type_number;
        }
    }

    if (user_options->style_type == LXW_TABLE_STYLE_TYPE_DARK) {
        if (user_options->style_type_number < 1
            || user_options->style_type_number > 11) {
            LXW_WARN_FORMAT1("worksheet_add_table(): "
                             "invalid style_type_number = %d for style type "
                             "LXW_TABLE_STYLE_TYPE_DARK. "
                             "Using default TableStyleMedium9",
                             user_options->style_type_number);

            table_obj->style_type = LXW_TABLE_STYLE_TYPE_MEDIUM;
            table_obj->style_type_number = 9;
        }
        else {
            table_obj->style_type_number = user_options->style_type_number;
        }
    }
}

/* Set the defaults for table columns in worksheet_add_table(). */
lxw_error
_set_default_table_columns(lxw_table_obj *table_obj)
{

    char col_name[LXW_ATTR_32];
    char *header;
    uint16_t i;
    lxw_table_column *column;
    uint16_t num_cols = table_obj->num_cols;
    lxw_table_column **columns = table_obj->columns;

    for (i = 0; i < num_cols; i++) {
        lxw_snprintf(col_name, LXW_ATTR_32, "Column%d", i + 1);

        column = calloc(num_cols, sizeof(lxw_table_column));
        RETURN_ON_MEM_ERROR(column, LXW_ERROR_MEMORY_MALLOC_FAILED);

        header = lxw_strdup(col_name);
        if (!header) {
            free(column);
            RETURN_ON_MEM_ERROR(header, LXW_ERROR_MEMORY_MALLOC_FAILED);
        }
        columns[i] = column;
        columns[i]->header = header;
    }

    return LXW_NO_ERROR;
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
        expanded = lxw_strdup_formula(formula);
    }
    else {
        /* Convert "@" in the formula string to "[#This Row],".  */
        expanded_len = strlen(formula) + (sizeof(LXW_THIS_ROW) * ref_count);
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
                strcat(&expanded[i], LXW_THIS_ROW);
                i += sizeof(LXW_THIS_ROW) - 1;
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

/* Set user values for table columns in worksheet_add_table(). */
lxw_error
_set_custom_table_columns(lxw_table_obj *table_obj,
                          lxw_table_options *user_options)
{
    char *str;
    uint16_t i;
    lxw_table_column *table_column;
    lxw_table_column *user_column;
    uint16_t num_cols = table_obj->num_cols;
    lxw_table_column **user_columns = user_options->columns;

    for (i = 0; i < num_cols; i++) {

        user_column = user_columns[i];
        table_column = table_obj->columns[i];

        /* NULL indicates end of user input array. */
        if (user_column == NULL)
            return LXW_NO_ERROR;

        if (user_column->header) {
            if (lxw_utf8_strlen(user_column->header) > 255) {
                LXW_WARN_FORMAT("worksheet_add_table(): column parameter "
                                "'header' exceeds Excel length limit of 255.");
                return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
            }

            str = lxw_strdup(user_column->header);
            RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

            /* Free the default column header. */
            free((void *) table_column->header);
            table_column->header = str;
        }

        if (user_column->total_string) {
            str = lxw_strdup(user_column->total_string);
            RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

            table_column->total_string = str;
        }

        if (user_column->formula) {
            str = _expand_table_formula(user_column->formula);
            RETURN_ON_MEM_ERROR(str, LXW_ERROR_MEMORY_MALLOC_FAILED);

            table_column->formula = str;
        }

        table_column->format = user_column->format;
        table_column->total_value = user_column->total_value;
        table_column->header_format = user_column->header_format;
        table_column->total_function = user_column->total_function;
    }

    return LXW_NO_ERROR;
}

/* Write a worksheet table column formula like SUBTOTAL(109,[Column1]). */
void
_write_column_function(lxw_worksheet *self, lxw_row_t row, lxw_col_t col,
                       lxw_table_column *column)
{
    size_t offset;
    char formula[LXW_MAX_ATTRIBUTE_LENGTH];
    lxw_format *format = column->format;
    uint8_t total_function = column->total_function;
    double value = column->total_value;
    const char *header = column->header;

    /* Write the subtotal formula number. */
    lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH, "SUBTOTAL(%d,[",
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

    worksheet_write_formula_num(self, row, col, formula, format, value);
}

/* Write a worksheet table column formula. */
void
_write_column_formula(lxw_worksheet *self, lxw_row_t first_row,
                      lxw_row_t last_row, lxw_col_t col,
                      lxw_table_column *column)
{
    lxw_row_t row;
    const char *formula = column->formula;
    lxw_format *format = column->format;

    for (row = first_row; row <= last_row; row++)
        worksheet_write_formula(self, row, col, formula, format);
}

/* Set the defaults for table columns in worksheet_add_table(). */
void
_write_table_column_data(lxw_worksheet *self, lxw_table_obj *table_obj)
{
    uint16_t i;
    lxw_table_column *column;
    lxw_table_column **columns = table_obj->columns;

    lxw_col_t col;
    lxw_row_t first_row = table_obj->first_row;
    lxw_col_t first_col = table_obj->first_col;
    lxw_row_t last_row = table_obj->last_row;
    lxw_row_t first_data_row = first_row;
    lxw_row_t last_data_row = last_row;

    if (!table_obj->no_header_row)
        first_data_row++;

    if (table_obj->total_row)
        last_data_row--;

    for (i = 0; i < table_obj->num_cols; i++) {
        col = first_col + i;
        column = columns[i];

        if (table_obj->no_header_row == LXW_FALSE)
            worksheet_write_string(self, first_row, col, column->header,
                                   column->header_format);

        if (column->total_string)
            worksheet_write_string(self, last_row, col, column->total_string,
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
lxw_error
_check_table_rows(lxw_row_t first_row, lxw_row_t last_row,
                  lxw_table_options *user_options)
{
    lxw_row_t num_non_header_rows = last_row - first_row;

    if (user_options && user_options->no_header_row == LXW_TRUE)
        num_non_header_rows++;

    if (num_non_header_rows == 0) {
        LXW_WARN_FORMAT("worksheet_add_table(): "
                        "table must have at least 1 non-header row.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    return LXW_NO_ERROR;
}

/*
 * Check that the the table name is valid.
 */
lxw_error
_check_table_name(lxw_table_options *user_options)
{
    const char *name;
    char *ptr;
    char first[2] = { 0, 0 };

    if (!user_options)
        return LXW_NO_ERROR;

    if (!user_options->name)
        return LXW_NO_ERROR;

    name = user_options->name;

    /* Check table name length. */
    if (lxw_utf8_strlen(name) > 255) {
        LXW_WARN_FORMAT("worksheet_add_table(): "
                        "Table name exceeds Excel's limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Check some short invalid names. */
    if (strlen(name) == 1
        && (name[0] == 'C' || name[0] == 'c' || name[0] == 'R'
            || name[0] == 'r')) {
        LXW_WARN_FORMAT1("worksheet_add_table(): "
                         "invalid table name \"%s\".", name);
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Check for invalid characters in Table name, while trying to allow
     * for utf8 strings. */
    ptr = strpbrk(name, " !\"#$%&'()*+,-/:;<=>?@[\\]^`{|}~");
    if (ptr) {
        LXW_WARN_FORMAT2("worksheet_add_table(): "
                         "invalid character '%c' in table name \"%s\".",
                         *ptr, name);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check for invalid initial character in Table name, while trying to allow
     * for utf8 strings. */
    first[0] = name[0];
    ptr = strpbrk(first, " !\"#$%&'()*+,-./0123456789:;<=>?@[\\]^`{|}~");
    if (ptr) {
        LXW_WARN_FORMAT2("worksheet_add_table(): "
                         "invalid first character '%c' in table name \"%s\".",
                         *ptr, name);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    return LXW_NO_ERROR;
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
_worksheet_xml_declaration(lxw_worksheet *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <worksheet> element.
 */
STATIC void
_worksheet_write_worksheet(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] = "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main";
    char xmlns_r[] = "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships";
    char xmlns_mc[] = "http://schemas.openxmlformats.org/"
        "markup-compatibility/2006";
    char xmlns_x14ac[] = "http://schemas.microsoft.com/"
        "office/spreadsheetml/2009/9/ac";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    if (self->excel_version == 2010) {
        LXW_PUSH_ATTRIBUTES_STR("xmlns:mc", xmlns_mc);
        LXW_PUSH_ATTRIBUTES_STR("xmlns:x14ac", xmlns_x14ac);
        LXW_PUSH_ATTRIBUTES_STR("mc:Ignorable", "x14ac");
    }

    lxw_xml_start_tag(self->file, "worksheet", &attributes);
    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <dimension> element.
 */
STATIC void
_worksheet_write_dimension(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char ref[LXW_MAX_CELL_RANGE_LENGTH];
    lxw_row_t dim_rowmin = self->dim_rowmin;
    lxw_row_t dim_rowmax = self->dim_rowmax;
    lxw_col_t dim_colmin = self->dim_colmin;
    lxw_col_t dim_colmax = self->dim_colmax;

    if (dim_rowmin == LXW_ROW_MAX && dim_colmin == LXW_COL_MAX) {
        /* If the rows and cols are still the defaults then no dimensions have
         * been set and we use the default range "A1". */
        lxw_rowcol_to_range(ref, 0, 0, 0, 0);
    }
    else if (dim_rowmin == LXW_ROW_MAX && dim_colmin != LXW_COL_MAX) {
        /* If the rows aren't set but the columns are then the dimensions have
         * been changed via set_column(). */
        lxw_rowcol_to_range(ref, 0, dim_colmin, 0, dim_colmax);
    }
    else {
        lxw_rowcol_to_range(ref, dim_rowmin, dim_colmin, dim_rowmax,
                            dim_colmax);
    }

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", ref);

    lxw_xml_empty_tag(self->file, "dimension", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <pane> element for freeze panes.
 */
STATIC void
_worksheet_write_freeze_panes(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    lxw_selection *selection;
    lxw_selection *user_selection;
    lxw_row_t row = self->panes.first_row;
    lxw_col_t col = self->panes.first_col;
    lxw_row_t top_row = self->panes.top_row;
    lxw_col_t left_col = self->panes.left_col;

    char row_cell[LXW_MAX_CELL_NAME_LENGTH];
    char col_cell[LXW_MAX_CELL_NAME_LENGTH];
    char top_left_cell[LXW_MAX_CELL_NAME_LENGTH];
    char active_pane[LXW_PANE_NAME_LENGTH];

    /* If there is a user selection we remove it from the list and use it. */
    if (!STAILQ_EMPTY(self->selections)) {
        user_selection = STAILQ_FIRST(self->selections);
        STAILQ_REMOVE_HEAD(self->selections, list_pointers);
    }
    else {
        /* or else create a new blank selection. */
        user_selection = calloc(1, sizeof(lxw_selection));
        RETURN_VOID_ON_MEM_ERROR(user_selection);
    }

    LXW_INIT_ATTRIBUTES();

    lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* Set the active pane. */
    if (row && col) {
        lxw_strcpy(active_pane, "bottomRight");

        lxw_rowcol_to_cell(row_cell, row, 0);
        lxw_rowcol_to_cell(col_cell, 0, col);

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "topRight");
            lxw_strcpy(selection->active_cell, col_cell);
            lxw_strcpy(selection->sqref, col_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomLeft");
            lxw_strcpy(selection->active_cell, row_cell);
            lxw_strcpy(selection->sqref, row_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomRight");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else if (col) {
        lxw_strcpy(active_pane, "topRight");

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "topRight");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else {
        lxw_strcpy(active_pane, "bottomLeft");

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomLeft");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }

    if (col)
        LXW_PUSH_ATTRIBUTES_INT("xSplit", col);

    if (row)
        LXW_PUSH_ATTRIBUTES_INT("ySplit", row);

    LXW_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);
    LXW_PUSH_ATTRIBUTES_STR("activePane", active_pane);

    if (self->panes.type == FREEZE_PANES)
        LXW_PUSH_ATTRIBUTES_STR("state", "frozen");
    else if (self->panes.type == FREEZE_SPLIT_PANES)
        LXW_PUSH_ATTRIBUTES_STR("state", "frozenSplit");

    lxw_xml_empty_tag(self->file, "pane", &attributes);

    free(user_selection);

    LXW_FREE_ATTRIBUTES();
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
_worksheet_write_split_panes(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    lxw_selection *selection;
    lxw_selection *user_selection;
    lxw_row_t row = self->panes.first_row;
    lxw_col_t col = self->panes.first_col;
    lxw_row_t top_row = self->panes.top_row;
    lxw_col_t left_col = self->panes.left_col;
    double x_split = self->panes.x_split;
    double y_split = self->panes.y_split;
    uint8_t has_selection = LXW_FALSE;

    char row_cell[LXW_MAX_CELL_NAME_LENGTH];
    char col_cell[LXW_MAX_CELL_NAME_LENGTH];
    char top_left_cell[LXW_MAX_CELL_NAME_LENGTH];
    char active_pane[LXW_PANE_NAME_LENGTH];

    /* If there is a user selection we remove it from the list and use it. */
    if (!STAILQ_EMPTY(self->selections)) {
        user_selection = STAILQ_FIRST(self->selections);
        STAILQ_REMOVE_HEAD(self->selections, list_pointers);
        has_selection = LXW_TRUE;
    }
    else {
        /* or else create a new blank selection. */
        user_selection = calloc(1, sizeof(lxw_selection));
        RETURN_VOID_ON_MEM_ERROR(user_selection);
    }

    LXW_INIT_ATTRIBUTES();

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
        top_row = (lxw_row_t) (0.5 + (y_split - 300.0) / 20.0 / 15.0);
        left_col = (lxw_col_t) (0.5 + (x_split - 390.0) / 20.0 / 3.0 / 16.0);
    }

    lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* If there is no selection set the active cell to the top left cell. */
    if (!has_selection) {
        lxw_strcpy(user_selection->active_cell, top_left_cell);
        lxw_strcpy(user_selection->sqref, top_left_cell);
    }

    /* Set the active pane. */
    if (y_split > 0.0 && x_split > 0.0) {
        lxw_strcpy(active_pane, "bottomRight");

        lxw_rowcol_to_cell(row_cell, top_row, 0);
        lxw_rowcol_to_cell(col_cell, 0, left_col);

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "topRight");
            lxw_strcpy(selection->active_cell, col_cell);
            lxw_strcpy(selection->sqref, col_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomLeft");
            lxw_strcpy(selection->active_cell, row_cell);
            lxw_strcpy(selection->sqref, row_cell);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomRight");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else if (x_split > 0.0) {
        lxw_strcpy(active_pane, "topRight");

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "topRight");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }
    else {
        lxw_strcpy(active_pane, "bottomLeft");

        selection = calloc(1, sizeof(lxw_selection));
        if (selection) {
            lxw_strcpy(selection->pane, "bottomLeft");
            lxw_strcpy(selection->active_cell, user_selection->active_cell);
            lxw_strcpy(selection->sqref, user_selection->sqref);

            STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);
        }
    }

    if (x_split > 0.0)
        LXW_PUSH_ATTRIBUTES_DBL("xSplit", x_split);

    if (y_split > 0.0)
        LXW_PUSH_ATTRIBUTES_DBL("ySplit", y_split);

    LXW_PUSH_ATTRIBUTES_STR("topLeftCell", top_left_cell);

    if (has_selection)
        LXW_PUSH_ATTRIBUTES_STR("activePane", active_pane);

    lxw_xml_empty_tag(self->file, "pane", &attributes);

    free(user_selection);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <selection> element.
 */
STATIC void
_worksheet_write_selection(lxw_worksheet *self, lxw_selection *selection)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (*selection->pane)
        LXW_PUSH_ATTRIBUTES_STR("pane", selection->pane);

    if (*selection->active_cell)
        LXW_PUSH_ATTRIBUTES_STR("activeCell", selection->active_cell);

    if (*selection->sqref)
        LXW_PUSH_ATTRIBUTES_STR("sqref", selection->sqref);

    lxw_xml_empty_tag(self->file, "selection", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <selection> elements.
 */
STATIC void
_worksheet_write_selections(lxw_worksheet *self)
{
    lxw_selection *selection;

    STAILQ_FOREACH(selection, self->selections, list_pointers) {
        _worksheet_write_selection(self, selection);
    }
}

/*
 * Write the frozen or split <pane> elements.
 */
STATIC void
_worksheet_write_panes(lxw_worksheet *self)
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
_worksheet_write_sheet_view(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    /* Hide screen gridlines if required */
    if (!self->screen_gridlines)
        LXW_PUSH_ATTRIBUTES_STR("showGridLines", "0");

    /* Hide zeroes in cells. */
    if (!self->show_zeros)
        LXW_PUSH_ATTRIBUTES_STR("showZeros", "0");

    /* Display worksheet right to left for Hebrew, Arabic and others. */
    if (self->right_to_left)
        LXW_PUSH_ATTRIBUTES_STR("rightToLeft", "1");

    /* Show that the sheet tab is selected. */
    if (self->selected)
        LXW_PUSH_ATTRIBUTES_STR("tabSelected", "1");

    /* Turn outlines off. Also required in the outlinePr element. */
    if (!self->outline_on)
        LXW_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

    /* Set the page view/layout mode if required. */
    if (self->page_view)
        LXW_PUSH_ATTRIBUTES_STR("view", "pageLayout");

    /* Set the top left cell if required. */
    if (self->top_left_cell[0])
        LXW_PUSH_ATTRIBUTES_STR("topLeftCell", self->top_left_cell);

    /* Set the zoom level. */
    if (self->zoom != 100 && !self->page_view) {
        LXW_PUSH_ATTRIBUTES_INT("zoomScale", self->zoom);

        if (self->zoom_scale_normal)
            LXW_PUSH_ATTRIBUTES_INT("zoomScaleNormal", self->zoom);
    }

    LXW_PUSH_ATTRIBUTES_STR("workbookViewId", "0");

    if (self->panes.type != NO_PANES || !STAILQ_EMPTY(self->selections)) {
        lxw_xml_start_tag(self->file, "sheetView", &attributes);
        _worksheet_write_panes(self);
        _worksheet_write_selections(self);
        lxw_xml_end_tag(self->file, "sheetView");
    }
    else {
        lxw_xml_empty_tag(self->file, "sheetView", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetViews> element.
 */
STATIC void
_worksheet_write_sheet_views(lxw_worksheet *self)
{
    lxw_xml_start_tag(self->file, "sheetViews", NULL);

    /* Write the sheetView element. */
    _worksheet_write_sheet_view(self);

    lxw_xml_end_tag(self->file, "sheetViews");
}

/*
 * Write the <sheetFormatPr> element.
 */
STATIC void
_worksheet_write_sheet_format_pr(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("defaultRowHeight", self->default_row_height);

    if (self->default_row_height != LXW_DEF_ROW_HEIGHT)
        LXW_PUSH_ATTRIBUTES_STR("customHeight", "1");

    if (self->default_row_zeroed)
        LXW_PUSH_ATTRIBUTES_STR("zeroHeight", "1");

    if (self->outline_row_level)
        LXW_PUSH_ATTRIBUTES_INT("outlineLevelRow", self->outline_row_level);

    if (self->outline_col_level)
        LXW_PUSH_ATTRIBUTES_INT("outlineLevelCol", self->outline_col_level);

    if (self->excel_version == 2010)
        LXW_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

    lxw_xml_empty_tag(self->file, "sheetFormatPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetData> element.
 */
STATIC void
_worksheet_write_sheet_data(lxw_worksheet *self)
{
    if (RB_EMPTY(self->table)) {
        lxw_xml_empty_tag(self->file, "sheetData", NULL);
    }
    else {
        lxw_xml_start_tag(self->file, "sheetData", NULL);
        _worksheet_write_rows(self);
        lxw_xml_end_tag(self->file, "sheetData");
    }
}

/*
 * Write the <sheetData> element when the memory optimization is on. In which
 * case we read the data stored in the temp file and rewrite it to the XML
 * sheet file.
 */
STATIC void
_worksheet_write_optimized_sheet_data(lxw_worksheet *self)
{
    size_t read_size = 1;
    char buffer[LXW_BUFFER_SIZE];

    if (self->dim_rowmin == LXW_ROW_MAX) {
        /* If the dimensions aren't defined then there is no data to write. */
        lxw_xml_empty_tag(self->file, "sheetData", NULL);
    }
    else {

        lxw_xml_start_tag(self->file, "sheetData", NULL);

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
                    fread(buffer, 1, LXW_BUFFER_SIZE, self->optimize_tmpfile);
                /* Ignore return value. There is no easy way to raise error. */
                (void) fwrite(buffer, 1, read_size, self->file);
            }
        }

        fclose(self->optimize_tmpfile);
        free(self->optimize_buffer);

        lxw_xml_end_tag(self->file, "sheetData");
    }
}

/*
 * Write the <pageMargins> element.
 */
STATIC void
_worksheet_write_page_margins(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    double left = self->margin_left;
    double right = self->margin_right;
    double top = self->margin_top;
    double bottom = self->margin_bottom;
    double header = self->margin_header;
    double footer = self->margin_footer;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_DBL("left", left);
    LXW_PUSH_ATTRIBUTES_DBL("right", right);
    LXW_PUSH_ATTRIBUTES_DBL("top", top);
    LXW_PUSH_ATTRIBUTES_DBL("bottom", bottom);
    LXW_PUSH_ATTRIBUTES_DBL("header", header);
    LXW_PUSH_ATTRIBUTES_DBL("footer", footer);

    lxw_xml_empty_tag(self->file, "pageMargins", &attributes);

    LXW_FREE_ATTRIBUTES();
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
_worksheet_write_page_setup(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (!self->page_setup_changed)
        return;

    /* Set paper size. */
    if (self->paper_size)
        LXW_PUSH_ATTRIBUTES_INT("paperSize", self->paper_size);

    /* Set the print_scale. */
    if (self->print_scale != 100)
        LXW_PUSH_ATTRIBUTES_INT("scale", self->print_scale);

    /* Set the "Fit to page" properties. */
    if (self->fit_page && self->fit_width != 1)
        LXW_PUSH_ATTRIBUTES_INT("fitToWidth", self->fit_width);

    if (self->fit_page && self->fit_height != 1)
        LXW_PUSH_ATTRIBUTES_INT("fitToHeight", self->fit_height);

    /* Set the page print direction. */
    if (self->page_order)
        LXW_PUSH_ATTRIBUTES_STR("pageOrder", "overThenDown");

    /* Set start page. */
    if (self->page_start > 1)
        LXW_PUSH_ATTRIBUTES_INT("firstPageNumber", self->page_start);

    /* Set page orientation. */
    if (self->orientation)
        LXW_PUSH_ATTRIBUTES_STR("orientation", "portrait");
    else
        LXW_PUSH_ATTRIBUTES_STR("orientation", "landscape");

    if (self->black_white)
        LXW_PUSH_ATTRIBUTES_STR("blackAndWhite", "1");

    /* Set start page active flag. */
    if (self->page_start)
        LXW_PUSH_ATTRIBUTES_INT("useFirstPageNumber", 1);

    /* Set the DPI. Mainly only for testing. */
    if (self->horizontal_dpi)
        LXW_PUSH_ATTRIBUTES_INT("horizontalDpi", self->horizontal_dpi);

    if (self->vertical_dpi)
        LXW_PUSH_ATTRIBUTES_INT("verticalDpi", self->vertical_dpi);

    lxw_xml_empty_tag(self->file, "pageSetup", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <printOptions> element.
 */
STATIC void
_worksheet_write_print_options(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    if (!self->print_options_changed)
        return;

    LXW_INIT_ATTRIBUTES();

    /* Set horizontal centering. */
    if (self->hcenter) {
        LXW_PUSH_ATTRIBUTES_STR("horizontalCentered", "1");
    }

    /* Set vertical centering. */
    if (self->vcenter) {
        LXW_PUSH_ATTRIBUTES_STR("verticalCentered", "1");
    }

    /* Enable row and column headers. */
    if (self->print_headers) {
        LXW_PUSH_ATTRIBUTES_STR("headings", "1");
    }

    /* Set printed gridlines. */
    if (self->print_gridlines) {
        LXW_PUSH_ATTRIBUTES_STR("gridLines", "1");
    }

    lxw_xml_empty_tag(self->file, "printOptions", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <row> element.
 */
STATIC void
_write_row(lxw_worksheet *self, lxw_row *row, char *spans)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    int32_t xf_index = 0;
    double height;

    if (row->format) {
        xf_index = lxw_format_get_xf_index(row->format);
    }

    if (row->height_changed)
        height = row->height;
    else
        height = self->default_row_height;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("r", row->row_num + 1);

    if (spans)
        LXW_PUSH_ATTRIBUTES_STR("spans", spans);

    if (xf_index)
        LXW_PUSH_ATTRIBUTES_INT("s", xf_index);

    if (row->format)
        LXW_PUSH_ATTRIBUTES_STR("customFormat", "1");

    if (height != LXW_DEF_ROW_HEIGHT)
        LXW_PUSH_ATTRIBUTES_DBL("ht", height);

    if (row->hidden)
        LXW_PUSH_ATTRIBUTES_STR("hidden", "1");

    if (height != LXW_DEF_ROW_HEIGHT)
        LXW_PUSH_ATTRIBUTES_STR("customHeight", "1");

    if (row->level)
        LXW_PUSH_ATTRIBUTES_INT("outlineLevel", row->level);

    if (row->collapsed)
        LXW_PUSH_ATTRIBUTES_STR("collapsed", "1");

    if (self->excel_version == 2010)
        LXW_PUSH_ATTRIBUTES_STR("x14ac:dyDescent", "0.25");

    if (!row->data_changed)
        lxw_xml_empty_tag(self->file, "row", &attributes);
    else
        lxw_xml_start_tag(self->file, "row", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Convert the width of a cell from user's units to pixels. Excel rounds the
 * column width to the nearest pixel. If the width hasn't been set by the user
 * we use the default value. If the column is hidden it has a value of zero.
 */
STATIC int32_t
_worksheet_size_col(lxw_worksheet *self, lxw_col_t col_num, uint8_t anchor)
{
    lxw_col_options *col_opt = NULL;
    uint32_t pixels;
    double width;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;
    lxw_col_t col_index;

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
        if (col_opt->hidden && anchor != LXW_OBJECT_MOVE_AND_SIZE_AFTER) {
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
_worksheet_size_row(lxw_worksheet *self, lxw_row_t row_num, uint8_t anchor)
{
    lxw_row *row;
    uint32_t pixels;

    row = lxw_worksheet_find_row(self, row_num);

    /* Note, the 0.75 below is due to the difference between 72/96 DPI. */
    if (row) {
        if (row->hidden && anchor != LXW_OBJECT_MOVE_AND_SIZE_AFTER)
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
_worksheet_position_object_pixels(lxw_worksheet *self,
                                  lxw_object_properties *object_props,
                                  lxw_drawing_object *drawing_object)
{
    lxw_col_t col_start;        /* Column containing upper left corner.  */
    int32_t x1;                 /* Distance to left side of object.      */

    lxw_row_t row_start;        /* Row containing top left corner.       */
    int32_t y1;                 /* Distance to top of object.            */

    lxw_col_t col_end;          /* Column containing lower right corner. */
    double x2;                  /* Distance to right side of object.     */

    lxw_row_t row_end;          /* Row containing bottom right corner.   */
    double y2;                  /* Distance to bottom of object.         */

    double width;               /* Width of object frame.                */
    double height;              /* Height of object frame.               */

    uint32_t x_abs = 0;         /* Abs. distance to left side of object. */
    uint32_t y_abs = 0;         /* Abs. distance to top  side of object. */

    uint32_t i;
    uint8_t anchor = drawing_object->anchor;
    uint8_t ignore_anchor = LXW_OBJECT_POSITION_DEFAULT;

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
           && col_end < LXW_COL_MAX) {
        width -= _worksheet_size_col(self, col_end, anchor);
        col_end++;
    }

    /* Subtract the underlying cell heights to find the end cell. */
    while (height >= _worksheet_size_row(self, row_end, anchor)
           && row_end < LXW_ROW_MAX) {
        height -= _worksheet_size_row(self, row_end, anchor);
        row_end++;
    }

    /* The end vertices are whatever is left from the width and height. */
    x2 = width;
    y2 = height;

    /* Add the dimensions to the drawing object. */
    drawing_object->from.col = col_start;
    drawing_object->from.row = row_start;
    drawing_object->from.col_offset = x1;
    drawing_object->from.row_offset = y1;
    drawing_object->to.col = col_end;
    drawing_object->to.row = row_end;
    drawing_object->to.col_offset = x2;
    drawing_object->to.row_offset = y2;
    drawing_object->col_absolute = x_abs;
    drawing_object->row_absolute = y_abs;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in EMUs. The vertices are expressed as English
 * Metric Units (EMUs). There are 12,700 EMUs per point.
 * Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
 */
STATIC void
_worksheet_position_object_emus(lxw_worksheet *self,
                                lxw_object_properties *image,
                                lxw_drawing_object *drawing_object)
{

    _worksheet_position_object_pixels(self, image, drawing_object);

    /* Convert the pixel values to EMUs. See above. */
    drawing_object->from.col_offset *= 9525;
    drawing_object->from.row_offset *= 9525;
    drawing_object->to.col_offset *= 9525;
    drawing_object->to.row_offset *= 9525;
    drawing_object->to.col_offset += 0.5;
    drawing_object->to.row_offset += 0.5;
    drawing_object->col_absolute *= 9525;
    drawing_object->row_absolute *= 9525;
}

/*
 * This function handles the additional optional parameters to
 * worksheet_write_comment_opt() as well as calculating the comment object
 * position and vertices.
 */
void
_get_comment_params(lxw_vml_obj *comment, lxw_comment_options *options)
{

    lxw_row_t start_row;
    lxw_col_t start_col;
    int32_t x_offset;
    int32_t y_offset;
    uint32_t height = 74;
    uint32_t width = 128;
    double x_scale = 1.0;
    double y_scale = 1.0;
    lxw_row_t row = comment->row;
    lxw_col_t col = comment->col;;

    /* Set the default start cell and offsets for the comment. These are
     * generally fixed in relation to the parent cell. However there are some
     * edge cases for cells at the, well yes, edges. */
    if (row == 0)
        y_offset = 2;
    else if (row == LXW_ROW_MAX - 3)
        y_offset = 16;
    else if (row == LXW_ROW_MAX - 2)
        y_offset = 16;
    else if (row == LXW_ROW_MAX - 1)
        y_offset = 14;
    else
        y_offset = 10;

    if (col == LXW_COL_MAX - 3)
        x_offset = 49;
    else if (col == LXW_COL_MAX - 2)
        x_offset = 49;
    else if (col == LXW_COL_MAX - 1)
        x_offset = 49;
    else
        x_offset = 15;

    if (row == 0)
        start_row = 0;
    else if (row == LXW_ROW_MAX - 3)
        start_row = LXW_ROW_MAX - 7;
    else if (row == LXW_ROW_MAX - 2)
        start_row = LXW_ROW_MAX - 6;
    else if (row == LXW_ROW_MAX - 1)
        start_row = LXW_ROW_MAX - 5;
    else
        start_row = row - 1;

    if (col == LXW_COL_MAX - 3)
        start_col = LXW_COL_MAX - 6;
    else if (col == LXW_COL_MAX - 2)
        start_col = LXW_COL_MAX - 5;
    else if (col == LXW_COL_MAX - 1)
        start_col = LXW_COL_MAX - 4;
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
        comment->author = lxw_strdup(options->author);
        comment->font_name = lxw_strdup(options->font_name);
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
 * worksheet_insert_button() as well as calculating the button object
 * position and vertices.
 */
lxw_error
_get_button_params(lxw_vml_obj *button, uint16_t button_number,
                   lxw_button_options *options)
{

    int32_t x_offset = 0;
    int32_t y_offset = 0;
    uint32_t height = LXW_DEF_ROW_HEIGHT_PIXELS;
    uint32_t width = LXW_DEF_COL_WIDTH_PIXELS;
    double x_scale = 1.0;
    double y_scale = 1.0;
    lxw_row_t row = button->row;
    lxw_col_t col = button->col;
    char buffer[LXW_ATTR_32];
    uint8_t has_caption = LXW_FALSE;
    uint8_t has_macro = LXW_FALSE;
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
            button->name = lxw_strdup(options->caption);
            RETURN_ON_MEM_ERROR(button->name, LXW_ERROR_MEMORY_MALLOC_FAILED);
            has_caption = LXW_TRUE;
        }

        if (options->macro) {
            len = sizeof("[0]!") + strlen(options->macro);
            button->macro = calloc(1, len);
            RETURN_ON_MEM_ERROR(button->macro,
                                LXW_ERROR_MEMORY_MALLOC_FAILED);

            if (button->macro)
                lxw_snprintf(button->macro, len, "[0]!%s", options->macro);

            has_macro = LXW_TRUE;
        }

        if (options->description) {
            button->text = lxw_strdup(options->description);
            RETURN_ON_MEM_ERROR(button->text, LXW_ERROR_MEMORY_MALLOC_FAILED);
        }
    }

    if (!has_caption) {
        lxw_snprintf(buffer, LXW_ATTR_32, "Button %d", button_number);
        button->name = lxw_strdup(buffer);
        RETURN_ON_MEM_ERROR(button->name, LXW_ERROR_MEMORY_MALLOC_FAILED);
    }

    if (!has_macro) {
        lxw_snprintf(buffer, LXW_ATTR_32, "[0]!Button%d_Click",
                     button_number);
        button->macro = lxw_strdup(buffer);
        RETURN_ON_MEM_ERROR(button->macro, LXW_ERROR_MEMORY_MALLOC_FAILED);
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

    return LXW_NO_ERROR;
}

/*
 * Calculate the vml_obj object position and vertices.
 */
void
_worksheet_position_vml_object(lxw_worksheet *self, lxw_vml_obj *vml_obj)
{
    lxw_object_properties object_props;
    lxw_drawing_object drawing_object;

    object_props.col = vml_obj->start_col;
    object_props.row = vml_obj->start_row;
    object_props.x_offset = vml_obj->x_offset;
    object_props.y_offset = vml_obj->y_offset;
    object_props.width = vml_obj->width;
    object_props.height = vml_obj->height;

    drawing_object.anchor = LXW_OBJECT_DONT_MOVE_DONT_SIZE;

    _worksheet_position_object_pixels(self, &object_props, &drawing_object);

    vml_obj->from.col = drawing_object.from.col;
    vml_obj->from.row = drawing_object.from.row;
    vml_obj->from.col_offset = drawing_object.from.col_offset;
    vml_obj->from.row_offset = drawing_object.from.row_offset;
    vml_obj->to.col = drawing_object.to.col;
    vml_obj->to.row = drawing_object.to.row;
    vml_obj->to.col_offset = drawing_object.to.col_offset;
    vml_obj->to.row_offset = drawing_object.to.row_offset;
    vml_obj->col_absolute = drawing_object.col_absolute;
    vml_obj->row_absolute = drawing_object.row_absolute;
}

/*
 * Set up image/drawings.
 */
void
lxw_worksheet_prepare_image(lxw_worksheet *self,
                            uint32_t image_ref_id, uint32_t drawing_id,
                            lxw_object_properties *object_props)
{
    lxw_drawing_object *drawing_object;
    lxw_rel_tuple *relationship;
    double width;
    double height;
    char *url;
    char *found_string;
    size_t i;
    char filename[LXW_FILENAME_LENGTH];
    enum cell_types link_type = HYPERLINK_URL;

    if (!self->drawing) {
        self->drawing = lxw_drawing_new();
        self->drawing->embedded = LXW_TRUE;
        RETURN_VOID_ON_MEM_ERROR(self->drawing);

        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/drawing");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", drawing_id);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_drawing_links, relationship,
                           list_pointers);
    }

    drawing_object = calloc(1, sizeof(lxw_drawing_object));
    RETURN_VOID_ON_MEM_ERROR(drawing_object);

    drawing_object->anchor = LXW_OBJECT_MOVE_DONT_SIZE;
    if (object_props->object_position)
        drawing_object->anchor = object_props->object_position;

    drawing_object->type = LXW_DRAWING_IMAGE;
    drawing_object->description = lxw_strdup(object_props->description);
    drawing_object->tip = lxw_strdup(object_props->tip);
    drawing_object->rel_index = 0;
    drawing_object->url_rel_index = 0;
    drawing_object->decorative = object_props->decorative;

    /* Scale to user scale. */
    width = object_props->width * object_props->x_scale;
    height = object_props->height * object_props->y_scale;

    /* Scale by non 96dpi resolutions. */
    width *= 96.0 / object_props->x_dpi;
    height *= 96.0 / object_props->y_dpi;

    object_props->width = width;
    object_props->height = height;

    _worksheet_position_object_emus(self, object_props, drawing_object);

    /* Convert from pixels to emus. */
    drawing_object->width = (uint32_t) (0.5 + width * 9525);
    drawing_object->height = (uint32_t) (0.5 + height * 9525);

    lxw_add_drawing_object(self->drawing, drawing_object);

    if (object_props->url) {
        url = object_props->url;

        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/hyperlink");
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
            relationship->target = lxw_strdup(url + sizeof("internal") - 1);
            GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

            /* We need to prefix the internal link/range with #. */
            relationship->target[0] = '#';
        }
        else if (link_type == HYPERLINK_EXTERNAL) {
            relationship->target_mode = lxw_strdup("External");
            GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);

            /* Look for Windows style "C:/" link or Windows share "\\" link. */
            found_string = strchr(url + sizeof("external:") - 1, ':');
            if (!found_string)
                found_string = strstr(url, "\\\\");

            if (found_string) {
                /* Copy the url with some space at the start to overwrite
                 * "external:" with "file:///". */
                relationship->target = lxw_escape_url_characters(url + 1,
                                                                 LXW_TRUE);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                /* Add the file:/// URI to the url if absolute path. */
                memcpy(relationship->target, "file:///",
                       sizeof("file:///") - 1);
            }
            else {
                /* Copy the relative url without "external:". */
                relationship->target =
                    lxw_escape_url_characters(url + sizeof("external:") - 1,
                                              LXW_TRUE);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                /* Switch backslash to forward slash. */
                for (i = 0; i <= strlen(relationship->target); i++)
                    if (relationship->target[i] == '\\')
                        relationship->target[i] = '/';
            }

        }
        else {
            relationship->target_mode = lxw_strdup("External");
            GOTO_LABEL_ON_MEM_ERROR(relationship->target_mode, mem_error);

            relationship->target =
                lxw_escape_url_characters(object_props->url, LXW_FALSE);
            GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);
        }

        /* Check if URL exceeds Excel's length limit. */
        if (lxw_utf8_strlen(url) > self->max_url_length) {
            LXW_WARN_FORMAT2("worksheet_insert_image()/_opt(): URL exceeds "
                             "Excel's allowable length of %d characters: %s",
                             self->max_url_length, url);
            goto mem_error;
        }

        if (!_find_drawing_rel_index(self, url)) {
            STAILQ_INSERT_TAIL(self->drawing_links, relationship,
                               list_pointers);
        }
        else {
            free(relationship->type);
            free(relationship->target);
            free(relationship->target_mode);
            free(relationship);
        }

        drawing_object->url_rel_index = _get_drawing_rel_index(self, url);

    }

    if (!_find_drawing_rel_index(self, object_props->md5)) {
        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/image");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                     object_props->extension);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->drawing_links, relationship, list_pointers);
    }

    drawing_object->rel_index =
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
lxw_worksheet_prepare_header_image(lxw_worksheet *self,
                                   uint32_t image_ref_id,
                                   lxw_object_properties *object_props)
{
    lxw_rel_tuple *relationship = NULL;
    char filename[LXW_FILENAME_LENGTH];
    lxw_vml_obj *header_image_vml;
    char *extension;

    STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);

    if (!_find_vml_drawing_rel_index(self, object_props->md5)) {
        relationship = calloc(1, sizeof(lxw_rel_tuple));
        RETURN_VOID_ON_MEM_ERROR(relationship);

        relationship->type = lxw_strdup("/image");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                     object_props->extension);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->vml_drawing_links, relationship,
                           list_pointers);
    }

    header_image_vml = calloc(1, sizeof(lxw_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(header_image_vml, mem_error);

    header_image_vml->width = (uint32_t) object_props->width;
    header_image_vml->height = (uint32_t) object_props->height;
    header_image_vml->x_dpi = object_props->x_dpi;
    header_image_vml->y_dpi = object_props->y_dpi;
    header_image_vml->rel_index = 1;

    header_image_vml->image_position =
        lxw_strdup(object_props->image_position);
    header_image_vml->name = lxw_strdup(object_props->description);

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
lxw_worksheet_prepare_background(lxw_worksheet *self,
                                 uint32_t image_ref_id,
                                 lxw_object_properties *object_props)
{
    lxw_rel_tuple *relationship = NULL;
    char filename[LXW_FILENAME_LENGTH];

    STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);

    relationship = calloc(1, sizeof(lxw_rel_tuple));
    RETURN_VOID_ON_MEM_ERROR(relationship);

    relationship->type = lxw_strdup("/image");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                 object_props->extension);

    relationship->target = lxw_strdup(filename);
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
lxw_worksheet_prepare_chart(lxw_worksheet *self,
                            uint32_t chart_ref_id,
                            uint32_t drawing_id,
                            lxw_object_properties *object_props,
                            uint8_t is_chartsheet)
{
    lxw_drawing_object *drawing_object;
    lxw_rel_tuple *relationship;
    double width;
    double height;
    char filename[LXW_FILENAME_LENGTH];

    if (!self->drawing) {
        self->drawing = lxw_drawing_new();
        RETURN_VOID_ON_MEM_ERROR(self->drawing);

        if (is_chartsheet) {
            self->drawing->embedded = LXW_FALSE;
            self->drawing->orientation = self->orientation;
        }
        else {
            self->drawing->embedded = LXW_TRUE;
        }

        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/drawing");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", drawing_id);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_drawing_links, relationship,
                           list_pointers);
    }

    drawing_object = calloc(1, sizeof(lxw_drawing_object));
    RETURN_VOID_ON_MEM_ERROR(drawing_object);

    drawing_object->anchor = LXW_OBJECT_MOVE_AND_SIZE;
    if (object_props->object_position)
        drawing_object->anchor = object_props->object_position;

    drawing_object->type = LXW_DRAWING_CHART;
    drawing_object->description = lxw_strdup(object_props->description);
    drawing_object->tip = NULL;
    drawing_object->rel_index = _get_drawing_rel_index(self, NULL);
    drawing_object->url_rel_index = 0;
    drawing_object->decorative = object_props->decorative;

    /* Scale to user scale. */
    width = object_props->width * object_props->x_scale;
    height = object_props->height * object_props->y_scale;

    /* Convert to the nearest pixel. */
    object_props->width = width;
    object_props->height = height;

    _worksheet_position_object_emus(self, object_props, drawing_object);

    /* Convert from pixels to emus. */
    drawing_object->width = (uint32_t) (0.5 + width * 9525);
    drawing_object->height = (uint32_t) (0.5 + height * 9525);

    lxw_add_drawing_object(self->drawing, drawing_object);

    relationship = calloc(1, sizeof(lxw_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxw_strdup("/chart");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxw_snprintf(filename, 32, "../charts/chart%d.xml", chart_ref_id);

    relationship->target = lxw_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    STAILQ_INSERT_TAIL(self->drawing_links, relationship, list_pointers);

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
lxw_worksheet_prepare_vml_objects(lxw_worksheet *self,
                                  uint32_t vml_data_id,
                                  uint32_t vml_shape_id,
                                  uint32_t vml_drawing_id,
                                  uint32_t comment_id)
{
    lxw_row *row;
    lxw_cell *cell;
    lxw_rel_tuple *relationship;
    char filename[LXW_FILENAME_LENGTH];
    uint32_t comment_count = 0;
    uint32_t i;
    uint32_t tmp_data_id;
    size_t data_str_len = 0;
    size_t used = 0;
    char *vml_data_id_str;

    RB_FOREACH(row, lxw_table_rows, self->comments) {

        RB_FOREACH(cell, lxw_table_cells, row->cells) {
            /* Calculate the worksheet position of the comment. */
            _worksheet_position_vml_object(self, cell->comment);

            /* Store comment in a simple list for use by packager. */
            STAILQ_INSERT_TAIL(self->comment_objs, cell->comment,
                               list_pointers);
            comment_count++;
        }
    }

    /* Set up the VML relationship for comments/buttons/header images. */
    relationship = calloc(1, sizeof(lxw_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxw_strdup("/vmlDrawing");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxw_snprintf(filename, 32, "../drawings/vmlDrawing%d.vml",
                 vml_drawing_id);

    relationship->target = lxw_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    self->external_vml_comment_link = relationship;

    if (self->has_comments) {
        /* Only need this relationship object for comment VMLs. */

        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/comments");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, 32, "../comments%d.xml", comment_id);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        self->external_comment_link = relationship;
    }

    /* The vml.c <o:idmap> element data id contains a comma separated range
     * when there is more than one 1024 block of comments, like this:
     * data="1,2,3". Since this could potentially (but unlikely) exceed
     * LXW_MAX_ATTRIBUTE_LENGTH we need to allocate space dynamically. */

    /* Calculate the total space required for the ID for each 1024 block. */
    for (i = 0; i <= comment_count / 1024; i++) {
        tmp_data_id = vml_data_id + i;

        /* Calculate the space required for the digits in the id. */
        while (tmp_data_id) {
            data_str_len++;
            tmp_data_id /= 10;
        }

        /* Add an extra char for comma separator or '\O'. */
        data_str_len++;
    };

    /* If this allocation fails it will be dealt with in packager.c. */
    vml_data_id_str = calloc(1, data_str_len + 2);
    GOTO_LABEL_ON_MEM_ERROR(vml_data_id_str, mem_error);

    /* Create the CSV list in the allocated space. */
    for (i = 0; i <= comment_count / 1024; i++) {
        tmp_data_id = vml_data_id + i;
        lxw_snprintf(vml_data_id_str + used, data_str_len - used, "%d,",
                     tmp_data_id);

        used = strlen(vml_data_id_str);
    };

    self->vml_shape_id = vml_shape_id;
    self->vml_data_id_str = vml_data_id_str;

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
lxw_worksheet_prepare_header_vml_objects(lxw_worksheet *self,
                                         uint32_t vml_header_id,
                                         uint32_t vml_drawing_id)
{

    lxw_rel_tuple *relationship;
    char filename[LXW_FILENAME_LENGTH];
    char *vml_data_id_str;

    self->vml_header_id = vml_header_id;

    /* Set up the VML relationship for header images. */
    relationship = calloc(1, sizeof(lxw_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxw_strdup("/vmlDrawing");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxw_snprintf(filename, 32, "../drawings/vmlDrawing%d.vml",
                 vml_drawing_id);

    relationship->target = lxw_strdup(filename);
    GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

    self->external_vml_header_link = relationship;

    /* If this allocation fails it will be dealt with in packager.c. */
    vml_data_id_str = calloc(1, sizeof("4294967295"));
    GOTO_LABEL_ON_MEM_ERROR(vml_data_id_str, mem_error);

    lxw_snprintf(vml_data_id_str, sizeof("4294967295"), "%d", vml_header_id);

    self->vml_header_id_str = vml_data_id_str;

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
lxw_worksheet_prepare_tables(lxw_worksheet *self, uint32_t table_id)
{
    lxw_table_obj *table_obj;
    lxw_rel_tuple *relationship;
    char name[LXW_ATTR_32];
    char filename[LXW_FILENAME_LENGTH];

    STAILQ_FOREACH(table_obj, self->table_objs, list_pointers) {

        relationship = calloc(1, sizeof(lxw_rel_tuple));
        GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

        relationship->type = lxw_strdup("/table");
        GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "../tables/table%d.xml", table_id);

        relationship->target = lxw_strdup(filename);
        GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

        STAILQ_INSERT_TAIL(self->external_table_links, relationship,
                           list_pointers);

        if (!table_obj->name) {
            lxw_snprintf(name, LXW_ATTR_32, "Table%d", table_id);
            table_obj->name = lxw_strdup(name);
            GOTO_LABEL_ON_MEM_ERROR(table_obj->name, mem_error);
        }
        table_obj->id = table_id;
        table_id++;
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
STATIC lxw_error
_process_png(lxw_object_properties *object_props)
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
        length = LXW_UINT32_NETWORK(length);

        /* The offset for next fseek() is the field length + type length. */
        offset = length + 4;

        if (memcmp(type, "IHDR", 4) == 0) {
            if (fread(&width, sizeof(width), 1, stream) < 1)
                break;

            if (fread(&height, sizeof(height), 1, stream) < 1)
                break;

            width = LXW_UINT32_NETWORK(width);
            height = LXW_UINT32_NETWORK(height);

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
                x_ppu = LXW_UINT32_NETWORK(x_ppu);
                y_ppu = LXW_UINT32_NETWORK(y_ppu);

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
    object_props->image_type = LXW_IMAGE_PNG;
    object_props->width = width;
    object_props->height = height;
    object_props->x_dpi = x_dpi ? x_dpi : 96;
    object_props->y_dpi = y_dpi ? y_dpi : 96;
    object_props->extension = lxw_strdup("png");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", object_props->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a JPEG file.
 */
STATIC lxw_error
_process_jpeg(lxw_object_properties *image_props)
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
        marker = LXW_UINT16_NETWORK(marker);
        length = LXW_UINT16_NETWORK(length);

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

            height = LXW_UINT16_NETWORK(height);
            width = LXW_UINT16_NETWORK(width);

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

            x_density = LXW_UINT16_NETWORK(x_density);
            y_density = LXW_UINT16_NETWORK(y_density);

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
    image_props->image_type = LXW_IMAGE_JPEG;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi ? x_dpi : 96;
    image_props->y_dpi = y_dpi ? y_dpi : 96;
    image_props->extension = lxw_strdup("jpeg");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a BMP file.
 */
STATIC lxw_error
_process_bmp(lxw_object_properties *image_props)
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

    height = LXW_UINT32_HOST(height);
    width = LXW_UINT32_HOST(width);

    /* Set the image metadata. */
    image_props->image_type = LXW_IMAGE_BMP;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi;
    image_props->y_dpi = y_dpi;
    image_props->extension = lxw_strdup("bmp");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a GIF file.
 */
STATIC lxw_error
_process_gif(lxw_object_properties *image_props)
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

    height = LXW_UINT16_HOST(height);
    width = LXW_UINT16_HOST(width);

    /* Set the image metadata. */
    image_props->image_type = LXW_IMAGE_GIF;
    image_props->width = width;
    image_props->height = height;
    image_props->x_dpi = x_dpi;
    image_props->y_dpi = y_dpi;
    image_props->extension = lxw_strdup("gif");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet image insertion: "
                     "no size data found in: %s.", image_props->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract information from the image file such as dimension, type, filename,
 * and extension.
 */
STATIC lxw_error
_get_image_properties(lxw_object_properties *image_props)
{
    unsigned char signature[4];
#ifndef USE_NO_MD5
    uint8_t i;
    MD5_CTX md5_context;
    size_t size_read;
    char buffer[LXW_IMAGE_BUFFER_SIZE];
    unsigned char md5_checksum[LXW_MD5_SIZE];
#endif

    /* Read 4 bytes to look for the file header/signature. */
    if (fread(signature, 1, 4, image_props->stream) < 4) {
        LXW_WARN_FORMAT1("worksheet image insertion: "
                         "couldn't read image type for: %s.",
                         image_props->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    if (memcmp(&signature[1], "PNG", 3) == 0) {
        if (_process_png(image_props) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (signature[0] == 0xFF && signature[1] == 0xD8) {
        if (_process_jpeg(image_props) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "BM", 2) == 0) {
        if (_process_bmp(image_props) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "GIF8", 4) == 0) {
        if (_process_gif(image_props) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else {
        LXW_WARN_FORMAT1("worksheet image insertion: "
                         "unsupported image format for: %s.",
                         image_props->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

#ifndef USE_NO_MD5
    /* Calculate an MD5 checksum for the image so that we can remove duplicate
     * images to reduce the xlsx file size.*/
    rewind(image_props->stream);

    MD5_Init(&md5_context);

    size_read = fread(buffer, 1, LXW_IMAGE_BUFFER_SIZE, image_props->stream);
    while (size_read) {
        MD5_Update(&md5_context, buffer, (unsigned long) size_read);
        size_read =
            fread(buffer, 1, LXW_IMAGE_BUFFER_SIZE, image_props->stream);
    }

    MD5_Final(md5_checksum, &md5_context);

    /* Create a 32 char hex string buffer for the MD5 checksum. */
    image_props->md5 = calloc(1, LXW_MD5_SIZE * 2 + 1);

    /* If this calloc fails we just return and don't remove duplicates. */
    RETURN_ON_MEM_ERROR(image_props->md5, LXW_NO_ERROR);

    /* Convert the 16 byte MD5 buffer to a 32 char hex string. */
    for (i = 0; i < LXW_MD5_SIZE; i++) {
        lxw_snprintf(&image_props->md5[2 * i], 3, "%02x", md5_checksum[i]);
    }
#endif

    return LXW_NO_ERROR;
}

/* Conditional formats that refer to the same cell sqref range, like A or
 * B1:B9, need to be written as part of one xml structure. Therefore we need
 * to store them in a RB hash/tree keyed by sqref. Within the RB hash element
 * we then store conditional formats that refer to sqref in a STAILQ list. */
lxw_error
_store_conditional_format_object(lxw_worksheet *self,
                                 lxw_cond_format_obj *cond_format)
{
    lxw_cond_format_hash_element tmp_hash_element;
    lxw_cond_format_hash_element *found_hash_element = NULL;
    lxw_cond_format_hash_element *new_hash_element = NULL;

    /* Create a temp hash element to do the lookup. */
    LXW_ATTRIBUTE_COPY(tmp_hash_element.sqref, cond_format->sqref);
    found_hash_element = RB_FIND(lxw_cond_format_hash,
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
        new_hash_element = calloc(1, sizeof(lxw_cond_format_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(new_hash_element, mem_error);

        /* Use the sqref as the key. */
        LXW_ATTRIBUTE_COPY(new_hash_element->sqref, cond_format->sqref);

        /* Also create the list where we store the cond format objects. */
        new_hash_element->cond_formats =
            calloc(1, sizeof(struct lxw_cond_format_list));
        GOTO_LABEL_ON_MEM_ERROR(new_hash_element->cond_formats, mem_error);

        /* Initialize the list and add the conditional format object. */
        STAILQ_INIT(new_hash_element->cond_formats);
        STAILQ_INSERT_TAIL(new_hash_element->cond_formats, cond_format,
                           list_pointers);

        /* Now insert the RB hash element into the tree. */
        RB_INSERT(lxw_cond_format_hash, self->conditional_formats,
                  new_hash_element);

    }

    return LXW_NO_ERROR;

mem_error:
    free(new_hash_element);
    return LXW_ERROR_MEMORY_MALLOC_FAILED;
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
_write_number_cell(lxw_worksheet *self, char *range,
                   int32_t style_index, lxw_cell *cell)
{
#ifdef USE_DTOA_LIBRARY
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, cell->u.number);

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
_write_string_cell(lxw_worksheet *self, char *range,
                   int32_t style_index, lxw_cell *cell)
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
_write_inline_string_cell(lxw_worksheet *self, char *range,
                          int32_t style_index, lxw_cell *cell)
{
    char *string = lxw_escape_data(cell->u.string);

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
_write_inline_rich_string_cell(lxw_worksheet *self, char *range,
                               int32_t style_index, lxw_cell *cell)
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
_write_formula_num_cell(lxw_worksheet *self, lxw_cell *cell)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, cell->formula_result);
    lxw_xml_data_element(self->file, "f", cell->u.string, NULL);
    lxw_xml_data_element(self->file, "v", data, NULL);
}

/*
 * Write out a formula worksheet cell with a numeric result.
 */
STATIC void
_write_formula_str_cell(lxw_worksheet *self, lxw_cell *cell)
{
    lxw_xml_data_element(self->file, "f", cell->u.string, NULL);
    lxw_xml_data_element(self->file, "v", cell->user_data2, NULL);
}

/*
 * Write out an array formula worksheet cell with a numeric result.
 */
STATIC void
_write_array_formula_num_cell(lxw_worksheet *self, lxw_cell *cell)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char data[LXW_ATTR_32];

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("t", "array");
    LXW_PUSH_ATTRIBUTES_STR("ref", cell->user_data1);

    lxw_sprintf_dbl(data, cell->formula_result);

    lxw_xml_data_element(self->file, "f", cell->u.string, &attributes);
    lxw_xml_data_element(self->file, "v", data, NULL);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write out a boolean worksheet cell.
 */
STATIC void
_write_boolean_cell(lxw_worksheet *self, lxw_cell *cell)
{
    char data[LXW_ATTR_32];

    if (cell->u.number == 0.0)
        data[0] = '0';
    else
        data[0] = '1';

    data[1] = '\0';

    lxw_xml_data_element(self->file, "v", data, NULL);
}

/*
 * Write out a error worksheet cell.
 */
STATIC void
_write_error_cell(lxw_worksheet *self)
{
    lxw_xml_data_element(self->file, "v", "#VALUE!", NULL);
}

/*
 * Calculate the "spans" attribute of the <row> tag. This is an XLSX
 * optimization and isn't strictly required. However, it makes comparing
 * files easier.
 *
 * The span is the same for each block of 16 rows.
 */
STATIC void
_calculate_spans(struct lxw_row *row, char *span, int32_t *block_num)
{
    lxw_cell *cell_min = RB_MIN(lxw_table_cells, row->cells);
    lxw_cell *cell_max = RB_MAX(lxw_table_cells, row->cells);
    lxw_col_t span_col_min = cell_min->col_num;
    lxw_col_t span_col_max = cell_max->col_num;
    lxw_col_t col_min;
    lxw_col_t col_max;
    *block_num = row->row_num / 16;

    row = RB_NEXT(lxw_table_rows, root, row);

    while (row && (int32_t) (row->row_num / 16) == *block_num) {

        if (!RB_EMPTY(row->cells)) {
            cell_min = RB_MIN(lxw_table_cells, row->cells);
            cell_max = RB_MAX(lxw_table_cells, row->cells);
            col_min = cell_min->col_num;
            col_max = cell_max->col_num;

            if (col_min < span_col_min)
                span_col_min = col_min;

            if (col_max > span_col_max)
                span_col_max = col_max;
        }

        row = RB_NEXT(lxw_table_rows, root, row);
    }

    lxw_snprintf(span, LXW_MAX_CELL_RANGE_LENGTH,
                 "%d:%d", span_col_min + 1, span_col_max + 1);
}

/*
 * Write out a generic worksheet cell.
 */
STATIC void
_write_cell(lxw_worksheet *self, lxw_cell *cell, lxw_format *row_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char range[LXW_MAX_CELL_NAME_LENGTH] = { 0 };
    lxw_row_t row_num = cell->row_num;
    lxw_col_t col_num = cell->col_num;
    int32_t style_index = 0;

    lxw_rowcol_to_cell(range, row_num, col_num);

    if (cell->format) {
        style_index = lxw_format_get_xf_index(cell->format);
    }
    else if (row_format) {
        style_index = lxw_format_get_xf_index(row_format);
    }
    else if (col_num < self->col_formats_max && self->col_formats[col_num]) {
        style_index = lxw_format_get_xf_index(self->col_formats[col_num]);
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
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r", range);

    if (style_index)
        LXW_PUSH_ATTRIBUTES_INT("s", style_index);

    if (cell->type == FORMULA_CELL) {
        /* If user_data2 is set then the formula has a string result. */
        if (cell->user_data2)
            LXW_PUSH_ATTRIBUTES_STR("t", "str");

        lxw_xml_start_tag(self->file, "c", &attributes);

        if (cell->user_data2)
            _write_formula_str_cell(self, cell);
        else
            _write_formula_num_cell(self, cell);

        lxw_xml_end_tag(self->file, "c");
    }
    else if (cell->type == BLANK_CELL) {
        if (cell->format)
            lxw_xml_empty_tag(self->file, "c", &attributes);
    }
    else if (cell->type == BOOLEAN_CELL) {
        LXW_PUSH_ATTRIBUTES_STR("t", "b");
        lxw_xml_start_tag(self->file, "c", &attributes);
        _write_boolean_cell(self, cell);
        lxw_xml_end_tag(self->file, "c");
    }
    else if (cell->type == ARRAY_FORMULA_CELL) {
        lxw_xml_start_tag(self->file, "c", &attributes);
        _write_array_formula_num_cell(self, cell);
        lxw_xml_end_tag(self->file, "c");
    }
    else if (cell->type == DYNAMIC_ARRAY_FORMULA_CELL) {
        LXW_PUSH_ATTRIBUTES_STR("cm", "1");
        lxw_xml_start_tag(self->file, "c", &attributes);
        _write_array_formula_num_cell(self, cell);
        lxw_xml_end_tag(self->file, "c");
    }
    else if (cell->type == ERROR_CELL) {
        LXW_PUSH_ATTRIBUTES_STR("t", "e");
        LXW_PUSH_ATTRIBUTES_DBL("vm", cell->u.number);
        lxw_xml_start_tag(self->file, "c", &attributes);
        _write_error_cell(self);
        lxw_xml_end_tag(self->file, "c");
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write out the worksheet data as a series of rows and cells.
 */
STATIC void
_worksheet_write_rows(lxw_worksheet *self)
{
    lxw_row *row;
    lxw_cell *cell;
    int32_t block_num = -1;
    char spans[LXW_MAX_CELL_RANGE_LENGTH] = { 0 };

    RB_FOREACH(row, lxw_table_rows, self->table) {

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
                RB_FOREACH(cell, lxw_table_cells, row->cells) {
                    _write_cell(self, cell, row->format);
                }

                lxw_xml_end_tag(self->file, "row");
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
lxw_worksheet_write_single_row(lxw_worksheet *self)
{
    lxw_row *row = self->optimize_row;
    lxw_col_t col;

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

        lxw_xml_end_tag(self->file, "row");
    }

    /* Reset the row. */
    row->height = LXW_DEF_ROW_HEIGHT;
    row->format = NULL;
    row->hidden = LXW_FALSE;
    row->level = 0;
    row->collapsed = LXW_FALSE;
    row->data_changed = LXW_FALSE;
    row->row_changed = LXW_FALSE;
}

/* Process a header/footer image and store it in the correct slot. */
lxw_error
_worksheet_set_header_footer_image(lxw_worksheet *self, const char *filename,
                                   uint8_t image_position)
{
    FILE *image_stream;
    const char *description;
    lxw_object_properties *object_props;
    char *image_strings[] = { "LH", "CH", "RH", "LF", "CF", "RF" };

    /* Not all slots will have image files. */
    if (!filename)
        return LXW_NO_ERROR;

    /* Check that the image file exists and can be opened. */
    image_stream = lxw_fopen(filename, "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Use the filename as the default description, like Excel. */
    description = lxw_basename(filename);
    if (!description) {
        LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                         "couldn't get basename for file: %s.", filename);
        fclose(image_stream);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup(filename);
    object_props->description = lxw_strdup(description);
    object_props->stream = image_stream;

    /* Set VML image position string based on the header/footer/position. */
    object_props->image_position = lxw_strdup(image_strings[image_position]);

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        *self->header_footer_objs[image_position] = object_props;
        self->has_header_vml = LXW_TRUE;
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Write the <col> element.
 */
STATIC void
_worksheet_write_col_info(lxw_worksheet *self, lxw_col_options *options)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    double width = options->width;
    uint8_t has_custom_width = LXW_TRUE;
    int32_t xf_index = 0;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;

    /* Get the format index. */
    if (options->format) {
        xf_index = lxw_format_get_xf_index(options->format);
    }

    /* Check if width is the Excel default. */
    if (width == LXW_DEF_COL_WIDTH) {

        /* The default col width changes to 0 for hidden columns. */
        if (options->hidden)
            width = 0;
        else
            has_custom_width = LXW_FALSE;

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

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("min", 1 + options->firstcol);
    LXW_PUSH_ATTRIBUTES_INT("max", 1 + options->lastcol);
    LXW_PUSH_ATTRIBUTES_DBL("width", width);

    if (xf_index)
        LXW_PUSH_ATTRIBUTES_INT("style", xf_index);

    if (options->hidden)
        LXW_PUSH_ATTRIBUTES_STR("hidden", "1");

    if (has_custom_width)
        LXW_PUSH_ATTRIBUTES_STR("customWidth", "1");

    if (options->level)
        LXW_PUSH_ATTRIBUTES_INT("outlineLevel", options->level);

    if (options->collapsed)
        LXW_PUSH_ATTRIBUTES_STR("collapsed", "1");

    lxw_xml_empty_tag(self->file, "col", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cols> element and <col> sub elements.
 */
STATIC void
_worksheet_write_cols(lxw_worksheet *self)
{
    lxw_col_t col;

    if (!self->col_size_changed)
        return;

    lxw_xml_start_tag(self->file, "cols", NULL);

    for (col = 0; col < self->col_options_max; col++) {
        if (self->col_options[col])
            _worksheet_write_col_info(self, self->col_options[col]);
    }

    lxw_xml_end_tag(self->file, "cols");
}

/*
 * Write the <mergeCell> element.
 */
STATIC void
_worksheet_write_merge_cell(lxw_worksheet *self,
                            lxw_merged_range *merged_range)
{

    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char ref[LXW_MAX_CELL_RANGE_LENGTH];

    LXW_INIT_ATTRIBUTES();

    /* Convert the merge dimensions to a cell range. */
    lxw_rowcol_to_range(ref, merged_range->first_row, merged_range->first_col,
                        merged_range->last_row, merged_range->last_col);

    LXW_PUSH_ATTRIBUTES_STR("ref", ref);

    lxw_xml_empty_tag(self->file, "mergeCell", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <mergeCells> element.
 */
STATIC void
_worksheet_write_merge_cells(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_merged_range *merged_range;

    if (self->merged_range_count) {
        LXW_INIT_ATTRIBUTES();

        LXW_PUSH_ATTRIBUTES_INT("count", self->merged_range_count);

        lxw_xml_start_tag(self->file, "mergeCells", &attributes);

        STAILQ_FOREACH(merged_range, self->merged_ranges, list_pointers) {
            _worksheet_write_merge_cell(self, merged_range);
        }
        lxw_xml_end_tag(self->file, "mergeCells");

        LXW_FREE_ATTRIBUTES();
    }
}

/*
 * Write the <oddHeader> element.
 */
STATIC void
_worksheet_write_odd_header(lxw_worksheet *self)
{
    lxw_xml_data_element(self->file, "oddHeader", self->header, NULL);
}

/*
 * Write the <oddFooter> element.
 */
STATIC void
_worksheet_write_odd_footer(lxw_worksheet *self)
{
    lxw_xml_data_element(self->file, "oddFooter", self->footer, NULL);
}

/*
 * Write the <headerFooter> element.
 */
STATIC void
_worksheet_write_header_footer(lxw_worksheet *self)
{
    if (!self->header_footer_changed)
        return;

    lxw_xml_start_tag(self->file, "headerFooter", NULL);

    if (self->header)
        _worksheet_write_odd_header(self);

    if (self->footer)
        _worksheet_write_odd_footer(self);

    lxw_xml_end_tag(self->file, "headerFooter");
}

/*
 * Write the <pageSetUpPr> element.
 */
STATIC void
_worksheet_write_page_set_up_pr(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!self->fit_page)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("fitToPage", "1");

    lxw_xml_empty_tag(self->file, "pageSetUpPr", &attributes);

    LXW_FREE_ATTRIBUTES();

}

/*
 * Write the <tabColor> element.
 */
STATIC void
_worksheet_write_tab_color(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rgb_str[LXW_ATTR_32];

    if (self->tab_color == LXW_COLOR_UNSET)
        return;

    lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X",
                 self->tab_color & LXW_COLOR_MASK);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

    lxw_xml_empty_tag(self->file, "tabColor", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <outlinePr> element.
 */
STATIC void
_worksheet_write_outline_pr(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!self->outline_changed)
        return;

    LXW_INIT_ATTRIBUTES();

    if (self->outline_style)
        LXW_PUSH_ATTRIBUTES_STR("applyStyles", "1");

    if (!self->outline_below)
        LXW_PUSH_ATTRIBUTES_STR("summaryBelow", "0");

    if (!self->outline_right)
        LXW_PUSH_ATTRIBUTES_STR("summaryRight", "0");

    if (!self->outline_on)
        LXW_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");

    lxw_xml_empty_tag(self->file, "outlinePr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <sheetPr> element for Sheet level properties.
 */
STATIC void
_worksheet_write_sheet_pr(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!self->fit_page
        && !self->filter_on
        && self->tab_color == LXW_COLOR_UNSET
        && !self->outline_changed
        && !self->vba_codename && !self->is_chartsheet) {
        return;
    }

    LXW_INIT_ATTRIBUTES();

    if (self->vba_codename)
        LXW_PUSH_ATTRIBUTES_STR("codeName", self->vba_codename);

    if (self->filter_on)
        LXW_PUSH_ATTRIBUTES_STR("filterMode", "1");

    if (self->fit_page || self->tab_color != LXW_COLOR_UNSET
        || self->outline_changed) {
        lxw_xml_start_tag(self->file, "sheetPr", &attributes);
        _worksheet_write_tab_color(self);
        _worksheet_write_outline_pr(self);
        _worksheet_write_page_set_up_pr(self);
        lxw_xml_end_tag(self->file, "sheetPr");
    }
    else {
        lxw_xml_empty_tag(self->file, "sheetPr", &attributes);
    }

    LXW_FREE_ATTRIBUTES();

}

/*
 * Write the <brk> element.
 */
STATIC void
_worksheet_write_brk(lxw_worksheet *self, uint32_t id, uint32_t max)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("id", id);
    LXW_PUSH_ATTRIBUTES_INT("max", max);
    LXW_PUSH_ATTRIBUTES_STR("man", "1");

    lxw_xml_empty_tag(self->file, "brk", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <rowBreaks> element.
 */
STATIC void
_worksheet_write_row_breaks(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint16_t count = self->hbreaks_count;
    uint16_t i;

    if (!count)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("count", count);
    LXW_PUSH_ATTRIBUTES_INT("manualBreakCount", count);

    lxw_xml_start_tag(self->file, "rowBreaks", &attributes);

    for (i = 0; i < count; i++)
        _worksheet_write_brk(self, self->hbreaks[i], LXW_COL_MAX - 1);

    lxw_xml_end_tag(self->file, "rowBreaks");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <colBreaks> element.
 */
STATIC void
_worksheet_write_col_breaks(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint16_t count = self->vbreaks_count;
    uint16_t i;

    if (!count)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("count", count);
    LXW_PUSH_ATTRIBUTES_INT("manualBreakCount", count);

    lxw_xml_start_tag(self->file, "colBreaks", &attributes);

    for (i = 0; i < count; i++)
        _worksheet_write_brk(self, self->vbreaks[i], LXW_ROW_MAX - 1);

    lxw_xml_end_tag(self->file, "colBreaks");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <filter> element.
 */
STATIC void
_worksheet_write_filter(lxw_worksheet *self, const char *str, double num,
                        uint8_t criteria)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (criteria == LXW_FILTER_CRITERIA_BLANKS)
        return;

    LXW_INIT_ATTRIBUTES();

    if (str)
        LXW_PUSH_ATTRIBUTES_STR("val", str);
    else
        LXW_PUSH_ATTRIBUTES_DBL("val", num);

    lxw_xml_empty_tag(self->file, "filter", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element as simple equality.
 */
STATIC void
_worksheet_write_filter_standard(lxw_worksheet *self,
                                 lxw_filter_rule_obj *filter)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (filter->has_blanks) {
        LXW_PUSH_ATTRIBUTES_STR("blank", "1");
    }

    if (filter->type == LXW_FILTER_TYPE_SINGLE && filter->has_blanks) {
        lxw_xml_empty_tag(self->file, "filters", &attributes);

    }
    else {
        lxw_xml_start_tag(self->file, "filters", &attributes);

        /* Write the filter element. */
        if (filter->type == LXW_FILTER_TYPE_SINGLE) {
            _worksheet_write_filter(self, filter->value1_string,
                                    filter->value1, filter->criteria1);
        }
        else if (filter->type == LXW_FILTER_TYPE_AND
                 || filter->type == LXW_FILTER_TYPE_OR) {
            _worksheet_write_filter(self, filter->value1_string,
                                    filter->value1, filter->criteria1);
            _worksheet_write_filter(self, filter->value2_string,
                                    filter->value2, filter->criteria2);
        }

        lxw_xml_end_tag(self->file, "filters");
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <customFilter> element.
 */
STATIC void
_worksheet_write_custom_filter(lxw_worksheet *self, const char *str,
                               double num, uint8_t criteria)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (criteria == LXW_FILTER_CRITERIA_NOT_EQUAL_TO)
        LXW_PUSH_ATTRIBUTES_STR("operator", "notEqual");
    if (criteria == LXW_FILTER_CRITERIA_GREATER_THAN)
        LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
    else if (criteria == LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO)
        LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
    else if (criteria == LXW_FILTER_CRITERIA_LESS_THAN)
        LXW_PUSH_ATTRIBUTES_STR("operator", "lessThan");
    else if (criteria == LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO)
        LXW_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");

    if (str)
        LXW_PUSH_ATTRIBUTES_STR("val", str);
    else
        LXW_PUSH_ATTRIBUTES_DBL("val", num);

    lxw_xml_empty_tag(self->file, "customFilter", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element as a list.
 */
STATIC void
_worksheet_write_filter_list(lxw_worksheet *self, lxw_filter_rule_obj *filter)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint16_t i;

    LXW_INIT_ATTRIBUTES();

    if (filter->has_blanks) {
        LXW_PUSH_ATTRIBUTES_STR("blank", "1");
    }

    lxw_xml_start_tag(self->file, "filters", &attributes);

    for (i = 0; i < filter->num_list_filters; i++) {
        _worksheet_write_filter(self, filter->list[i], 0, 0);
    }

    lxw_xml_end_tag(self->file, "filters");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element.
 */
STATIC void
_worksheet_write_filter_custom(lxw_worksheet *self,
                               lxw_filter_rule_obj *filter)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (filter->type == LXW_FILTER_TYPE_AND)
        LXW_PUSH_ATTRIBUTES_STR("and", "1");

    lxw_xml_start_tag(self->file, "customFilters", &attributes);

    /* Write the filter element. */
    if (filter->type == LXW_FILTER_TYPE_SINGLE) {
        _worksheet_write_custom_filter(self, filter->value1_string,
                                       filter->value1, filter->criteria1);
    }
    else if (filter->type == LXW_FILTER_TYPE_AND
             || filter->type == LXW_FILTER_TYPE_OR) {
        _worksheet_write_custom_filter(self, filter->value1_string,
                                       filter->value1, filter->criteria1);
        _worksheet_write_custom_filter(self, filter->value2_string,
                                       filter->value2, filter->criteria2);
    }

    lxw_xml_end_tag(self->file, "customFilters");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <filterColumn> element.
 */
STATIC void
_worksheet_write_filter_column(lxw_worksheet *self,
                               lxw_filter_rule_obj *filter)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!filter)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("colId", filter->col_num);

    lxw_xml_start_tag(self->file, "filterColumn", &attributes);

    if (filter->list)
        _worksheet_write_filter_list(self, filter);
    else if (filter->is_custom)
        _worksheet_write_filter_custom(self, filter);
    else
        _worksheet_write_filter_standard(self, filter);

    lxw_xml_end_tag(self->file, "filterColumn");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <autoFilter> element.
 */
STATIC void
_worksheet_write_auto_filter(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char range[LXW_MAX_CELL_RANGE_LENGTH];
    uint16_t i;

    if (!self->autofilter.in_use)
        return;

    lxw_rowcol_to_range(range,
                        self->autofilter.first_row,
                        self->autofilter.first_col,
                        self->autofilter.last_row, self->autofilter.last_col);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", range);

    if (self->autofilter.has_rules) {
        lxw_xml_start_tag(self->file, "autoFilter", &attributes);

        for (i = 0; i < self->num_filter_rules; i++)
            _worksheet_write_filter_column(self, self->filter_rules[i]);

        lxw_xml_end_tag(self->file, "autoFilter");

    }
    else {
        lxw_xml_empty_tag(self->file, "autoFilter", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <hyperlink> element for external links.
 */
STATIC void
_worksheet_write_hyperlink_external(lxw_worksheet *self, lxw_row_t row_num,
                                    lxw_col_t col_num, const char *location,
                                    const char *tooltip, uint16_t id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char ref[LXW_MAX_CELL_NAME_LENGTH];
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_rowcol_to_cell(ref, row_num, col_num);

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", ref);
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    if (location)
        LXW_PUSH_ATTRIBUTES_STR("location", location);

    if (tooltip)
        LXW_PUSH_ATTRIBUTES_STR("tooltip", tooltip);

    lxw_xml_empty_tag(self->file, "hyperlink", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <hyperlink> element for internal links.
 */
STATIC void
_worksheet_write_hyperlink_internal(lxw_worksheet *self, lxw_row_t row_num,
                                    lxw_col_t col_num, const char *location,
                                    const char *display, const char *tooltip)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char ref[LXW_MAX_CELL_NAME_LENGTH];

    lxw_rowcol_to_cell(ref, row_num, col_num);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", ref);

    if (location)
        LXW_PUSH_ATTRIBUTES_STR("location", location);

    if (tooltip)
        LXW_PUSH_ATTRIBUTES_STR("tooltip", tooltip);

    if (display)
        LXW_PUSH_ATTRIBUTES_STR("display", display);

    lxw_xml_empty_tag(self->file, "hyperlink", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Process any stored hyperlinks in row/col order and write the <hyperlinks>
 * element. The attributes are different for internal and external links.
 */
STATIC void
_worksheet_write_hyperlinks(lxw_worksheet *self)
{

    lxw_row *row;
    lxw_cell *link;
    lxw_rel_tuple *relationship;

    if (RB_EMPTY(self->hyperlinks))
        return;

    /* Write the hyperlink elements. */
    lxw_xml_start_tag(self->file, "hyperlinks", NULL);

    RB_FOREACH(row, lxw_table_rows, self->hyperlinks) {

        RB_FOREACH(link, lxw_table_cells, row->cells) {

            if (link->type == HYPERLINK_URL
                || link->type == HYPERLINK_EXTERNAL) {

                self->rel_count++;

                relationship = calloc(1, sizeof(lxw_rel_tuple));
                GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

                relationship->type = lxw_strdup("/hyperlink");
                GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

                relationship->target = lxw_strdup(link->u.string);
                GOTO_LABEL_ON_MEM_ERROR(relationship->target, mem_error);

                relationship->target_mode = lxw_strdup("External");
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

    lxw_xml_end_tag(self->file, "hyperlinks");
    return;

mem_error:
    if (relationship) {
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
    lxw_xml_end_tag(self->file, "hyperlinks");
}

/*
 * Write the <sheetProtection> element.
 */
STATIC void
_worksheet_write_sheet_protection(lxw_worksheet *self,
                                  lxw_protection_obj *protect)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!protect->is_configured)
        return;

    LXW_INIT_ATTRIBUTES();

    if (*protect->hash)
        LXW_PUSH_ATTRIBUTES_STR("password", protect->hash);

    if (!protect->no_sheet)
        LXW_PUSH_ATTRIBUTES_INT("sheet", 1);

    if (!protect->no_content)
        LXW_PUSH_ATTRIBUTES_INT("content", 1);

    if (!protect->objects)
        LXW_PUSH_ATTRIBUTES_INT("objects", 1);

    if (!protect->scenarios)
        LXW_PUSH_ATTRIBUTES_INT("scenarios", 1);

    if (protect->format_cells)
        LXW_PUSH_ATTRIBUTES_INT("formatCells", 0);

    if (protect->format_columns)
        LXW_PUSH_ATTRIBUTES_INT("formatColumns", 0);

    if (protect->format_rows)
        LXW_PUSH_ATTRIBUTES_INT("formatRows", 0);

    if (protect->insert_columns)
        LXW_PUSH_ATTRIBUTES_INT("insertColumns", 0);

    if (protect->insert_rows)
        LXW_PUSH_ATTRIBUTES_INT("insertRows", 0);

    if (protect->insert_hyperlinks)
        LXW_PUSH_ATTRIBUTES_INT("insertHyperlinks", 0);

    if (protect->delete_columns)
        LXW_PUSH_ATTRIBUTES_INT("deleteColumns", 0);

    if (protect->delete_rows)
        LXW_PUSH_ATTRIBUTES_INT("deleteRows", 0);

    if (protect->no_select_locked_cells)
        LXW_PUSH_ATTRIBUTES_INT("selectLockedCells", 1);

    if (protect->sort)
        LXW_PUSH_ATTRIBUTES_INT("sort", 0);

    if (protect->autofilter)
        LXW_PUSH_ATTRIBUTES_INT("autoFilter", 0);

    if (protect->pivot_tables)
        LXW_PUSH_ATTRIBUTES_INT("pivotTables", 0);

    if (protect->no_select_unlocked_cells)
        LXW_PUSH_ATTRIBUTES_INT("selectUnlockedCells", 1);

    lxw_xml_empty_tag(self->file, "sheetProtection", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <legacyDrawing> element.
 */
STATIC void
_worksheet_write_legacy_drawing(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_vml)
        return;
    else
        self->rel_count++;

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", self->rel_count);
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "legacyDrawing", &attributes);

    LXW_FREE_ATTRIBUTES();

}

/*
 * Write the <legacyDrawingHF> element.
 */
STATIC void
_worksheet_write_legacy_drawing_hf(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_header_vml)
        return;
    else
        self->rel_count++;

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", self->rel_count);
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "legacyDrawingHF", &attributes);

    LXW_FREE_ATTRIBUTES();

}

/*
 * Write the <picture> element.
 */
STATIC void
_worksheet_write_picture(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    if (!self->has_background_image)
        return;
    else
        self->rel_count++;

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", self->rel_count);
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "picture", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <drawing> element.
 */
STATIC void
_worksheet_write_drawing(lxw_worksheet *self, uint16_t id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "drawing", &attributes);

    LXW_FREE_ATTRIBUTES();

}

/*
 * Write the <drawing> elements.
 */
STATIC void
_worksheet_write_drawings(lxw_worksheet *self)
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
_worksheet_write_formula1_num(lxw_worksheet *self, double number)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, number);

    lxw_xml_data_element(self->file, "formula1", data, NULL);
}

/*
 * Write the <formula1> element for strings/formulas.
 */
STATIC void
_worksheet_write_formula1_str(lxw_worksheet *self, char *str)
{
    lxw_xml_data_element(self->file, "formula1", str, NULL);
}

/*
 * Write the <formula2> element for numbers.
 */
STATIC void
_worksheet_write_formula2_num(lxw_worksheet *self, double number)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, number);

    lxw_xml_data_element(self->file, "formula2", data, NULL);
}

/*
 * Write the <formula2> element for strings/formulas.
 */
STATIC void
_worksheet_write_formula2_str(lxw_worksheet *self, char *str)
{
    lxw_xml_data_element(self->file, "formula2", str, NULL);
}

/*
 * Write the <dataValidation> element.
 */
STATIC void
_worksheet_write_data_validation(lxw_worksheet *self,
                                 lxw_data_val_obj *validation)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint8_t is_between = 0;

    LXW_INIT_ATTRIBUTES();

    switch (validation->validate) {
        case LXW_VALIDATION_TYPE_INTEGER:
        case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
            LXW_PUSH_ATTRIBUTES_STR("type", "whole");
            break;
        case LXW_VALIDATION_TYPE_DECIMAL:
        case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
            LXW_PUSH_ATTRIBUTES_STR("type", "decimal");
            break;
        case LXW_VALIDATION_TYPE_LIST:
        case LXW_VALIDATION_TYPE_LIST_FORMULA:
            LXW_PUSH_ATTRIBUTES_STR("type", "list");
            break;
        case LXW_VALIDATION_TYPE_DATE:
        case LXW_VALIDATION_TYPE_DATE_FORMULA:
        case LXW_VALIDATION_TYPE_DATE_NUMBER:
            LXW_PUSH_ATTRIBUTES_STR("type", "date");
            break;
        case LXW_VALIDATION_TYPE_TIME:
        case LXW_VALIDATION_TYPE_TIME_FORMULA:
        case LXW_VALIDATION_TYPE_TIME_NUMBER:
            LXW_PUSH_ATTRIBUTES_STR("type", "time");
            break;
        case LXW_VALIDATION_TYPE_LENGTH:
        case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
            LXW_PUSH_ATTRIBUTES_STR("type", "textLength");
            break;
        case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
            LXW_PUSH_ATTRIBUTES_STR("type", "custom");
            break;
    }

    switch (validation->criteria) {
        case LXW_VALIDATION_CRITERIA_EQUAL_TO:
            LXW_PUSH_ATTRIBUTES_STR("operator", "equal");
            break;
        case LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO:
            LXW_PUSH_ATTRIBUTES_STR("operator", "notEqual");
            break;
        case LXW_VALIDATION_CRITERIA_LESS_THAN:
            LXW_PUSH_ATTRIBUTES_STR("operator", "lessThan");
            break;
        case LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO:
            LXW_PUSH_ATTRIBUTES_STR("operator", "lessThanOrEqual");
            break;
        case LXW_VALIDATION_CRITERIA_GREATER_THAN:
            LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThan");
            break;
        case LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
            LXW_PUSH_ATTRIBUTES_STR("operator", "greaterThanOrEqual");
            break;
        case LXW_VALIDATION_CRITERIA_BETWEEN:
            /* Between is the default for 2 formulas and isn't added. */
            is_between = 1;
            break;
        case LXW_VALIDATION_CRITERIA_NOT_BETWEEN:
            is_between = 1;
            LXW_PUSH_ATTRIBUTES_STR("operator", "notBetween");
            break;
    }

    if (validation->error_type == LXW_VALIDATION_ERROR_TYPE_WARNING)
        LXW_PUSH_ATTRIBUTES_STR("errorStyle", "warning");

    if (validation->error_type == LXW_VALIDATION_ERROR_TYPE_INFORMATION)
        LXW_PUSH_ATTRIBUTES_STR("errorStyle", "information");

    if (validation->ignore_blank)
        LXW_PUSH_ATTRIBUTES_INT("allowBlank", 1);

    if (validation->dropdown == LXW_VALIDATION_OFF)
        LXW_PUSH_ATTRIBUTES_INT("showDropDown", 1);

    if (validation->show_input)
        LXW_PUSH_ATTRIBUTES_INT("showInputMessage", 1);

    if (validation->show_error)
        LXW_PUSH_ATTRIBUTES_INT("showErrorMessage", 1);

    if (validation->error_title)
        LXW_PUSH_ATTRIBUTES_STR("errorTitle", validation->error_title);

    if (validation->error_message)
        LXW_PUSH_ATTRIBUTES_STR("error", validation->error_message);

    if (validation->input_title)
        LXW_PUSH_ATTRIBUTES_STR("promptTitle", validation->input_title);

    if (validation->input_message)
        LXW_PUSH_ATTRIBUTES_STR("prompt", validation->input_message);

    LXW_PUSH_ATTRIBUTES_STR("sqref", validation->sqref);

    if (validation->validate == LXW_VALIDATION_TYPE_ANY)
        lxw_xml_empty_tag(self->file, "dataValidation", &attributes);
    else
        lxw_xml_start_tag(self->file, "dataValidation", &attributes);

    /* Write the formula1 and formula2 elements. */
    switch (validation->validate) {
        case LXW_VALIDATION_TYPE_INTEGER:
        case LXW_VALIDATION_TYPE_DECIMAL:
        case LXW_VALIDATION_TYPE_LENGTH:
        case LXW_VALIDATION_TYPE_DATE:
        case LXW_VALIDATION_TYPE_TIME:
        case LXW_VALIDATION_TYPE_DATE_NUMBER:
        case LXW_VALIDATION_TYPE_TIME_NUMBER:
            _worksheet_write_formula1_num(self, validation->value_number);
            if (is_between)
                _worksheet_write_formula2_num(self,
                                              validation->maximum_number);
            break;
        case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
        case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
        case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
        case LXW_VALIDATION_TYPE_DATE_FORMULA:
        case LXW_VALIDATION_TYPE_TIME_FORMULA:
        case LXW_VALIDATION_TYPE_LIST:
        case LXW_VALIDATION_TYPE_LIST_FORMULA:
        case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
            _worksheet_write_formula1_str(self, validation->value_formula);
            if (is_between)
                _worksheet_write_formula2_str(self,
                                              validation->maximum_formula);
            break;
    }

    if (validation->validate != LXW_VALIDATION_TYPE_ANY)
        lxw_xml_end_tag(self->file, "dataValidation");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <dataValidations> element.
 */
STATIC void
_worksheet_write_data_validations(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_data_val_obj *data_validation;

    if (self->num_validations == 0)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("count", self->num_validations);

    lxw_xml_start_tag(self->file, "dataValidations", &attributes);

    STAILQ_FOREACH(data_validation, self->data_validations, list_pointers) {
        /* Write the dataValidation element. */
        _worksheet_write_data_validation(self, data_validation);
    }

    lxw_xml_end_tag(self->file, "dataValidations");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <formula> element for strings.
 */
STATIC void
_worksheet_write_formula_str(lxw_worksheet *self, char *data)
{
    lxw_xml_data_element(self->file, "formula", data, NULL);
}

/*
 * Write the <formula> element for numbers.
 */
STATIC void
_worksheet_write_formula_num(lxw_worksheet *self, double num)
{
    char data[LXW_ATTR_32];

    lxw_sprintf_dbl(data, num);
    lxw_xml_data_element(self->file, "formula", data, NULL);
}

/*
 * Write the <ext> element.
 */
STATIC void
_worksheet_write_ext(lxw_worksheet *self, char *uri)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_x_14[] =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:x14", xmlns_x_14);
    LXW_PUSH_ATTRIBUTES_STR("uri", uri);

    lxw_xml_start_tag(self->file, "ext", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <extLst> dataBar extension element.
 */
STATIC void
_worksheet_write_data_bar_ext(lxw_worksheet *self,
                              lxw_cond_format_obj *cond_format)
{
    /* Create a pseudo GUID for each unique Excel 2010 data bar. */
    cond_format->guid = calloc(1, LXW_GUID_LENGTH);
    lxw_snprintf(cond_format->guid, LXW_GUID_LENGTH,
                 "{DA7ABA51-AAAA-BBBB-%04X-%012X}",
                 self->index + 1, ++self->data_bar_2010_index);

    lxw_xml_start_tag(self->file, "extLst", NULL);

    _worksheet_write_ext(self, "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}");

    lxw_xml_data_element(self->file, "x14:id", cond_format->guid, NULL);

    lxw_xml_end_tag(self->file, "ext");
    lxw_xml_end_tag(self->file, "extLst");
}

/*
 * Write the <color> element.
 */
STATIC void
_worksheet_write_color(lxw_worksheet *self, lxw_color_t color)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rgb[LXW_ATTR_32];

    lxw_snprintf(rgb, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("rgb", rgb);

    lxw_xml_empty_tag(self->file, "color", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfvo> element for strings.
 */
STATIC void
_worksheet_write_cfvo_str(lxw_worksheet *self, uint8_t rule_type,
                          char *value, uint8_t data_bar_2010)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM)
        LXW_PUSH_ATTRIBUTES_STR("type", "min");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER)
        LXW_PUSH_ATTRIBUTES_STR("type", "num");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT)
        LXW_PUSH_ATTRIBUTES_STR("type", "percent");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE)
        LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA)
        LXW_PUSH_ATTRIBUTES_STR("type", "formula");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM)
        LXW_PUSH_ATTRIBUTES_STR("type", "max");

    if (!data_bar_2010 || (rule_type != LXW_CONDITIONAL_RULE_TYPE_MINIMUM
                           && rule_type != LXW_CONDITIONAL_RULE_TYPE_MAXIMUM))
        LXW_PUSH_ATTRIBUTES_STR("val", value);

    lxw_xml_empty_tag(self->file, "cfvo", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfvo> element for numbers.
 */
STATIC void
_worksheet_write_cfvo_num(lxw_worksheet *self, uint8_t rule_type,
                          double value, uint8_t data_bar_2010)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM)
        LXW_PUSH_ATTRIBUTES_STR("type", "min");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER)
        LXW_PUSH_ATTRIBUTES_STR("type", "num");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT)
        LXW_PUSH_ATTRIBUTES_STR("type", "percent");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE)
        LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA)
        LXW_PUSH_ATTRIBUTES_STR("type", "formula");
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM)
        LXW_PUSH_ATTRIBUTES_STR("type", "max");

    if (!data_bar_2010 || (rule_type != LXW_CONDITIONAL_RULE_TYPE_MINIMUM
                           && rule_type != LXW_CONDITIONAL_RULE_TYPE_MAXIMUM))
        LXW_PUSH_ATTRIBUTES_DBL("val", value);

    lxw_xml_empty_tag(self->file, "cfvo", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <iconSet> element.
 */
STATIC void
_worksheet_write_icon_set(lxw_worksheet *self,
                          lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
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
    uint8_t percent = LXW_CONDITIONAL_RULE_TYPE_PERCENT;
    uint8_t style = cond_format->icon_style;

    LXW_INIT_ATTRIBUTES();

    if (style != LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED)
        LXW_PUSH_ATTRIBUTES_STR("iconSet", icon_set[style]);

    if (cond_format->reverse_icons == LXW_TRUE)
        LXW_PUSH_ATTRIBUTES_STR("reverse", "1");

    if (cond_format->icons_only == LXW_TRUE)
        LXW_PUSH_ATTRIBUTES_STR("showValue", "0");

    lxw_xml_start_tag(self->file, "iconSet", &attributes);

    if (style < LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED) {
        _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 33, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 67, LXW_FALSE);
    }

    if (style >= LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED
        && style < LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED) {
        _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 25, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 50, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 75, LXW_FALSE);
    }

    if (style >= LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED
        && style <= LXW_CONDITIONAL_ICONS_5_QUARTERS) {
        _worksheet_write_cfvo_num(self, percent, 0, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 20, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 40, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 60, LXW_FALSE);
        _worksheet_write_cfvo_num(self, percent, 80, LXW_FALSE);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for data bar rules.
 */
STATIC void
_worksheet_write_cf_rule_icons(lxw_worksheet *self,
                               lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    _worksheet_write_icon_set(self, cond_format);

    lxw_xml_end_tag(self->file, "iconSet");
    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <dataBar> element.
 */
STATIC void
_worksheet_write_data_bar(lxw_worksheet *self,
                          lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (cond_format->bar_only)
        LXW_PUSH_ATTRIBUTES_STR("showValue", "0");

    lxw_xml_start_tag(self->file, "dataBar", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for data bar rules.
 */
STATIC void
_worksheet_write_cf_rule_data_bar(lxw_worksheet *self,
                                  lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxw_xml_start_tag(self->file, "cfRule", &attributes);

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

    lxw_xml_end_tag(self->file, "dataBar");

    if (cond_format->data_bar_2010)
        _worksheet_write_data_bar_ext(self, cond_format);

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for 2 and 3 color scale rules.
 */
STATIC void
_worksheet_write_cf_rule_color_scale(lxw_worksheet *self,
                                     lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);
    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);
    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    lxw_xml_start_tag(self->file, "colorScale", NULL);

    if (cond_format->min_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->min_rule_type,
                                  cond_format->min_value_string, LXW_FALSE);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->min_rule_type,
                                  cond_format->min_value, LXW_FALSE);
    }

    if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
        if (cond_format->mid_value_string) {
            _worksheet_write_cfvo_str(self, cond_format->mid_rule_type,
                                      cond_format->mid_value_string,
                                      LXW_FALSE);
        }
        else {
            _worksheet_write_cfvo_num(self, cond_format->mid_rule_type,
                                      cond_format->mid_value, LXW_FALSE);
        }
    }

    if (cond_format->max_value_string) {
        _worksheet_write_cfvo_str(self, cond_format->max_rule_type,
                                  cond_format->max_value_string, LXW_FALSE);
    }
    else {
        _worksheet_write_cfvo_num(self, cond_format->max_rule_type,
                                  cond_format->max_value, LXW_FALSE);
    }

    _worksheet_write_color(self, cond_format->min_color);

    if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE)
        _worksheet_write_color(self, cond_format->mid_color);

    _worksheet_write_color(self, cond_format->max_color);

    lxw_xml_end_tag(self->file, "colorScale");
    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for formula rules.
 */
STATIC void
_worksheet_write_cf_rule_formula(lxw_worksheet *self,
                                 lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    _worksheet_write_formula_str(self, cond_format->min_value_string);

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for top and bottom rules.
 */
STATIC void
_worksheet_write_cf_rule_top(lxw_worksheet *self,
                             lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    if (cond_format->criteria ==
        LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT)
        LXW_PUSH_ATTRIBUTES_INT("percent", 1);

    if (cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM)
        LXW_PUSH_ATTRIBUTES_INT("bottom", 1);

    /* Rank must be an int in the range 1-1000 . */
    if (cond_format->min_value < 1.0 || cond_format->min_value > 1000.0)
        LXW_PUSH_ATTRIBUTES_DBL("rank", 10);
    else
        LXW_PUSH_ATTRIBUTES_DBL("rank", (uint16_t) cond_format->min_value);

    lxw_xml_empty_tag(self->file, "cfRule", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for unique/duplicate rules.
 */
STATIC void
_worksheet_write_cf_rule_duplicate(lxw_worksheet *self,
                                   lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    /* Set the attributes common to all rule types. */
    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    lxw_xml_empty_tag(self->file, "cfRule", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for averages rules.
 */
STATIC void
_worksheet_write_cf_rule_average(lxw_worksheet *self,
                                 lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint8_t criteria = cond_format->criteria;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
        LXW_PUSH_ATTRIBUTES_INT("aboveAverage", 0);

    if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL)
        LXW_PUSH_ATTRIBUTES_INT("equalAverage", 1);

    if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW)
        LXW_PUSH_ATTRIBUTES_INT("stdDev", 1);

    if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW)
        LXW_PUSH_ATTRIBUTES_INT("stdDev", 2);

    if (criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE
        || criteria == LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW)
        LXW_PUSH_ATTRIBUTES_INT("stdDev", 3);

    lxw_xml_empty_tag(self->file, "cfRule", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for time_period rules.
 */
STATIC void
_worksheet_write_cf_rule_time_period(lxw_worksheet *self,
                                     lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char formula[LXW_MAX_ATTRIBUTE_LENGTH];
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

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    pos = criteria - LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY;
    LXW_PUSH_ATTRIBUTES_STR("timePeriod", time_periods[pos]);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()-1", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "FLOOR(%s,1)=TODAY()+1", first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),"
                     "TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,"
                     "ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),"
                     "ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY())-1,OR(YEAR(%s)=YEAR("
                     "TODAY()),AND(MONTH(%s)=1,YEAR(A1)=YEAR(TODAY())-1)))",
                     first_cell, first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))",
                     first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {
        lxw_snprintf(formula, LXW_MAX_ATTRIBUTE_LENGTH,
                     "AND(MONTH(%s)=MONTH(TODAY())+1,OR(YEAR(%s)=YEAR("
                     "TODAY()),AND(MONTH(%s)=12,YEAR(%s)=YEAR(TODAY())+1)))",
                     first_cell, first_cell, first_cell, first_cell);
        _worksheet_write_formula_str(self, formula);
    }

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for blanks/no_blanks, errors/no_errors rules.
 */
STATIC void
_worksheet_write_cf_rule_blanks(lxw_worksheet *self,
                                lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char formula[LXW_ATTR_32];
    uint8_t type = cond_format->type;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    if (type == LXW_CONDITIONAL_TYPE_BLANKS) {
        lxw_snprintf(formula, LXW_ATTR_32, "LEN(TRIM(%s))=0",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXW_CONDITIONAL_TYPE_NO_BLANKS) {
        lxw_snprintf(formula, LXW_ATTR_32, "LEN(TRIM(%s))>0",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXW_CONDITIONAL_TYPE_ERRORS) {
        lxw_snprintf(formula, LXW_ATTR_32, "ISERROR(%s)",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (type == LXW_CONDITIONAL_TYPE_NO_ERRORS) {
        lxw_snprintf(formula, LXW_ATTR_32, "NOT(ISERROR(%s))",
                     cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for text rules.
 */
STATIC void
_worksheet_write_cf_rule_text(lxw_worksheet *self,
                              lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint8_t pos;
    char formula[LXW_ATTR_32 * 2];
    char *operators[] = {
        "containsText",
        "notContains",
        "beginsWith",
        "endsWith",
    };
    uint8_t criteria = cond_format->criteria;

    LXW_INIT_ATTRIBUTES();

    if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING)
        LXW_PUSH_ATTRIBUTES_STR("type", "containsText");
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING)
        LXW_PUSH_ATTRIBUTES_STR("type", "notContainsText");
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH)
        LXW_PUSH_ATTRIBUTES_STR("type", "beginsWith");
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH)
        LXW_PUSH_ATTRIBUTES_STR("type", "endsWith");

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    pos = criteria - LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING;
    LXW_PUSH_ATTRIBUTES_STR("operator", operators[pos]);

    LXW_PUSH_ATTRIBUTES_STR("text", cond_format->min_value_string);

    lxw_xml_start_tag(self->file, "cfRule", &attributes);

    if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING) {
        lxw_snprintf(formula, LXW_ATTR_32 * 2,
                     "NOT(ISERROR(SEARCH(\"%s\",%s)))",
                     cond_format->min_value_string, cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING) {
        lxw_snprintf(formula, LXW_ATTR_32 * 2,
                     "ISERROR(SEARCH(\"%s\",%s))",
                     cond_format->min_value_string, cond_format->first_cell);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH) {
        lxw_snprintf(formula, LXW_ATTR_32 * 2,
                     "LEFT(%s,%d)=\"%s\"",
                     cond_format->first_cell,
                     (uint16_t) strlen(cond_format->min_value_string),
                     cond_format->min_value_string);
        _worksheet_write_formula_str(self, formula);
    }
    else if (criteria == LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {
        lxw_snprintf(formula, LXW_ATTR_32 * 2,
                     "RIGHT(%s,%d)=\"%s\"",
                     cond_format->first_cell,
                     (uint16_t) strlen(cond_format->min_value_string),
                     cond_format->min_value_string);
        _worksheet_write_formula_str(self, formula);
    }

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element for cell rules.
 */
STATIC void
_worksheet_write_cf_rule_cell(lxw_worksheet *self,
                              lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
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

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("type", cond_format->type_string);

    if (cond_format->dxf_index != LXW_PROPERTY_UNSET)
        LXW_PUSH_ATTRIBUTES_INT("dxfId", cond_format->dxf_index);

    LXW_PUSH_ATTRIBUTES_INT("priority", cond_format->dxf_priority);

    if (cond_format->stop_if_true)
        LXW_PUSH_ATTRIBUTES_INT("stopIfTrue", 1);

    LXW_PUSH_ATTRIBUTES_STR("operator", operators[cond_format->criteria]);

    lxw_xml_start_tag(self->file, "cfRule", &attributes);

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

    lxw_xml_end_tag(self->file, "cfRule");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <cfRule> element.
 */
STATIC void
_worksheet_write_cf_rule(lxw_worksheet *self,
                         lxw_cond_format_obj *cond_format)
{
    if (cond_format->type == LXW_CONDITIONAL_TYPE_CELL) {

        _worksheet_write_cf_rule_cell(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TEXT) {

        _worksheet_write_cf_rule_text(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {

        _worksheet_write_cf_rule_time_period(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_DUPLICATE
             || cond_format->type == LXW_CONDITIONAL_TYPE_UNIQUE) {

        _worksheet_write_cf_rule_duplicate(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_AVERAGE) {

        _worksheet_write_cf_rule_average(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {

        _worksheet_write_cf_rule_top(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_BLANKS
             || cond_format->type == LXW_CONDITIONAL_TYPE_NO_BLANKS
             || cond_format->type == LXW_CONDITIONAL_TYPE_ERRORS
             || cond_format->type == LXW_CONDITIONAL_TYPE_NO_ERRORS) {

        _worksheet_write_cf_rule_blanks(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_FORMULA) {

        _worksheet_write_cf_rule_formula(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_2_COLOR_SCALE
             || cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {

        _worksheet_write_cf_rule_color_scale(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_DATA_BAR) {

        _worksheet_write_cf_rule_data_bar(self, cond_format);
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_ICON_SETS) {

        _worksheet_write_cf_rule_icons(self, cond_format);
    }

}

/*
 * Write the <conditionalFormatting> element.
 */
STATIC void
_worksheet_write_conditional_formatting(lxw_worksheet *self,
                                        lxw_cond_format_hash_element *element)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_cond_format_obj *cond_format;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("sqref", element->sqref);

    lxw_xml_start_tag(self->file, "conditionalFormatting", &attributes);

    STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
        /* Write the cfRule element. */
        _worksheet_write_cf_rule(self, cond_format);
    }

    lxw_xml_end_tag(self->file, "conditionalFormatting");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the conditional formatting> elements.
 */
STATIC void
_worksheet_write_conditional_formats(lxw_worksheet *self)
{
    lxw_cond_format_hash_element *element;
    lxw_cond_format_hash_element *next_element;

    for (element = RB_MIN(lxw_cond_format_hash, self->conditional_formats);
         element; element = next_element) {

        _worksheet_write_conditional_formatting(self, element);

        next_element =
            RB_NEXT(lxw_cond_format_hash, self->conditional_formats, element);
    }
}

/*
 * Write the <x14:xxxColor> elements for data bar conditional formats.
 */
STATIC void
_worksheet_write_x14_color(lxw_worksheet *self, char *type, lxw_color_t color)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rgb[LXW_ATTR_32];

    lxw_snprintf(rgb, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("rgb", rgb);
    lxw_xml_empty_tag(self->file, type, &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:cfvo> element.
 */
STATIC void
_worksheet_write_x14_cfvo(lxw_worksheet *self, uint8_t rule_type,
                          double number, char *string)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char data[LXW_ATTR_32];
    uint8_t has_value = LXW_FALSE;

    LXW_INIT_ATTRIBUTES();

    if (!string)
        lxw_sprintf_dbl(data, number);

    if (rule_type == LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN) {
        LXW_PUSH_ATTRIBUTES_STR("type", "autoMin");
        has_value = LXW_FALSE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MINIMUM) {
        LXW_PUSH_ATTRIBUTES_STR("type", "min");
        has_value = LXW_FALSE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_NUMBER) {
        LXW_PUSH_ATTRIBUTES_STR("type", "num");
        has_value = LXW_TRUE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENT) {
        LXW_PUSH_ATTRIBUTES_STR("type", "percent");
        has_value = LXW_TRUE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_PERCENTILE) {
        LXW_PUSH_ATTRIBUTES_STR("type", "percentile");
        has_value = LXW_TRUE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_FORMULA) {
        LXW_PUSH_ATTRIBUTES_STR("type", "formula");
        has_value = LXW_TRUE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        LXW_PUSH_ATTRIBUTES_STR("type", "max");
        has_value = LXW_FALSE;
    }
    else if (rule_type == LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX) {
        LXW_PUSH_ATTRIBUTES_STR("type", "autoMax");
        has_value = LXW_FALSE;
    }

    if (has_value) {
        lxw_xml_start_tag(self->file, "x14:cfvo", &attributes);

        if (string)
            lxw_xml_data_element(self->file, "xm:f", string, NULL);
        else
            lxw_xml_data_element(self->file, "xm:f", data, NULL);

        lxw_xml_end_tag(self->file, "x14:cfvo");
    }
    else {
        lxw_xml_empty_tag(self->file, "x14:cfvo", &attributes);
    }
    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:dataBar> element.
 */
STATIC void
_worksheet_write_x14_data_bar(lxw_worksheet *self,
                              lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char min_length[] = "0";
    char max_length[] = "100";
    char border[] = "1";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("minLength", min_length);
    LXW_PUSH_ATTRIBUTES_STR("maxLength", max_length);

    if (!cond_format->bar_no_border)
        LXW_PUSH_ATTRIBUTES_STR("border", border);

    if (cond_format->bar_solid)
        LXW_PUSH_ATTRIBUTES_STR("gradient", "0");

    if (cond_format->bar_direction ==
        LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT)
        LXW_PUSH_ATTRIBUTES_STR("direction", "rightToLeft");

    if (cond_format->bar_direction ==
        LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT)
        LXW_PUSH_ATTRIBUTES_STR("direction", "leftToRight");

    if (cond_format->bar_negative_color_same)
        LXW_PUSH_ATTRIBUTES_STR("negativeBarColorSameAsPositive", "1");

    if (!cond_format->bar_no_border
        && !cond_format->bar_negative_border_color_same)
        LXW_PUSH_ATTRIBUTES_STR("negativeBarBorderColorSameAsPositive", "0");

    if (cond_format->bar_axis_position == LXW_CONDITIONAL_BAR_AXIS_MIDPOINT)
        LXW_PUSH_ATTRIBUTES_STR("axisPosition", "middle");

    if (cond_format->bar_axis_position == LXW_CONDITIONAL_BAR_AXIS_NONE)
        LXW_PUSH_ATTRIBUTES_STR("axisPosition", "none");

    lxw_xml_start_tag(self->file, "x14:dataBar", &attributes);

    if (cond_format->auto_min)
        cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN;

    _worksheet_write_x14_cfvo(self, cond_format->min_rule_type,
                              cond_format->min_value,
                              cond_format->min_value_string);

    if (cond_format->auto_max)
        cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX;

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

    if (cond_format->bar_axis_position != LXW_CONDITIONAL_BAR_AXIS_NONE)
        _worksheet_write_x14_color(self, "x14:axisColor",
                                   cond_format->bar_axis_color);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x14:cfRule> element.
 */
STATIC void
_worksheet_write_x14_cf_rule(lxw_worksheet *self,
                             lxw_cond_format_obj *cond_format)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("type", "dataBar");
    LXW_PUSH_ATTRIBUTES_STR("id", cond_format->guid);

    lxw_xml_start_tag(self->file, "x14:cfRule", &attributes);

    /* Write the x14:dataBar element. */
    _worksheet_write_x14_data_bar(self, cond_format);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xm:sqref> element.
 */
STATIC void
_worksheet_write_xm_sqref(lxw_worksheet *self,
                          lxw_cond_format_obj *cond_format)
{
    lxw_xml_data_element(self->file, "xm:sqref", cond_format->sqref, NULL);
}

/*
 * Write the <conditionalFormatting> element.
 */
STATIC void
_worksheet_write_conditional_formatting_2010(lxw_worksheet *self, lxw_cond_format_hash_element
                                             *element)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_cond_format_obj *cond_format;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:xm",
                            "http://schemas.microsoft.com/office/excel/2006/main");

    STAILQ_FOREACH(cond_format, element->cond_formats, list_pointers) {
        if (!cond_format->data_bar_2010)
            continue;

        lxw_xml_start_tag(self->file, "x14:conditionalFormatting",
                          &attributes);

        _worksheet_write_x14_cf_rule(self, cond_format);

        lxw_xml_end_tag(self->file, "x14:dataBar");
        lxw_xml_end_tag(self->file, "x14:cfRule");
        _worksheet_write_xm_sqref(self, cond_format);
        lxw_xml_end_tag(self->file, "x14:conditionalFormatting");
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <extLst> element for Excel 2010 conditional formatting data bars.
 */
STATIC void
_worksheet_write_ext_list_data_bars(lxw_worksheet *self)
{
    lxw_cond_format_hash_element *element;
    lxw_cond_format_hash_element *next_element;

    _worksheet_write_ext(self, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
    lxw_xml_start_tag(self->file, "x14:conditionalFormattings", NULL);

    for (element = RB_MIN(lxw_cond_format_hash, self->conditional_formats);
         element; element = next_element) {

        _worksheet_write_conditional_formatting_2010(self, element);

        next_element =
            RB_NEXT(lxw_cond_format_hash, self->conditional_formats, element);
    }

    lxw_xml_end_tag(self->file, "x14:conditionalFormattings");
    lxw_xml_end_tag(self->file, "ext");
}

/*
 * Write the <extLst> element.
 */
STATIC void
_worksheet_write_ext_list(lxw_worksheet *self)
{
    if (self->data_bar_2010_index == 0)
        return;

    lxw_xml_start_tag(self->file, "extLst", NULL);

    _worksheet_write_ext_list_data_bars(self);

    lxw_xml_end_tag(self->file, "extLst");
}

/*
 * Write the <ignoredError> element.
 */
STATIC void
_worksheet_write_ignored_error(lxw_worksheet *self, char *ignore_error,
                               char *range)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("sqref", range);
    LXW_PUSH_ATTRIBUTES_STR(ignore_error, "1");

    lxw_xml_empty_tag(self->file, "ignoredError", &attributes);

    LXW_FREE_ATTRIBUTES();
}

lxw_error
_validate_conditional_icons(lxw_conditional_format *user)
{
    if (user->icon_style > LXW_CONDITIONAL_ICONS_5_QUARTERS) {

        LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                         "For type = LXW_CONDITIONAL_TYPE_ICON_SETS, "
                         "invalid icon_style (%d).", user->icon_style);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXW_NO_ERROR;
    }
}

lxw_error
_validate_conditional_data_bar(lxw_worksheet *self,
                               lxw_cond_format_obj *cond_format,
                               lxw_conditional_format *user_options)
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

        cond_format->data_bar_2010 = LXW_TRUE;
        self->excel_version = 2010;
    }

    if (min_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
        && min_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->min_rule_type = min_rule_type;
        cond_format->min_value = user_options->min_value;
        cond_format->min_value_string =
            lxw_strdup_formula(user_options->min_value_string);
    }
    else {
        cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_MINIMUM;
        cond_format->min_value = 0;
    }

    if (max_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
        && max_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->max_rule_type = max_rule_type;
        cond_format->max_value = user_options->max_value;
        cond_format->max_value_string =
            lxw_strdup_formula(user_options->max_value_string);
    }
    else {
        cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_MAXIMUM;
        cond_format->max_value = 0;
    }

    if (cond_format->data_bar_2010) {
        if (min_rule_type == LXW_CONDITIONAL_RULE_TYPE_NONE)
            cond_format->auto_min = LXW_TRUE;
        if (max_rule_type == LXW_CONDITIONAL_RULE_TYPE_NONE)
            cond_format->auto_max = LXW_TRUE;
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

    if (user_options->bar_color != LXW_COLOR_UNSET)
        cond_format->bar_color = user_options->bar_color;
    else
        cond_format->bar_color = 0x638EC6;

    if (user_options->bar_negative_color != LXW_COLOR_UNSET)
        cond_format->bar_negative_color = user_options->bar_negative_color;
    else
        cond_format->bar_negative_color = 0xFF0000;

    if (user_options->bar_border_color != LXW_COLOR_UNSET)
        cond_format->bar_border_color = user_options->bar_border_color;
    else
        cond_format->bar_border_color = cond_format->bar_color;

    if (user_options->bar_negative_border_color != LXW_COLOR_UNSET)
        cond_format->bar_negative_border_color =
            user_options->bar_negative_border_color;
    else
        cond_format->bar_negative_border_color = 0xFF0000;

    if (user_options->bar_axis_color != LXW_COLOR_UNSET)
        cond_format->bar_axis_color = user_options->bar_axis_color;
    else
        cond_format->bar_axis_color = 0x000000;

    return LXW_NO_ERROR;
}

lxw_error
_validate_conditional_scale(lxw_cond_format_obj *cond_format,
                            lxw_conditional_format *user_options)
{
    uint8_t min_rule_type = user_options->min_rule_type;
    uint8_t mid_rule_type = user_options->mid_rule_type;
    uint8_t max_rule_type = user_options->max_rule_type;

    if (min_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
        && min_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->min_rule_type = min_rule_type;
        cond_format->min_value = user_options->min_value;
        cond_format->min_value_string =
            lxw_strdup_formula(user_options->min_value_string);
    }
    else {
        cond_format->min_rule_type = LXW_CONDITIONAL_RULE_TYPE_MINIMUM;
        cond_format->min_value = 0;
    }

    if (max_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
        && max_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
        cond_format->max_rule_type = max_rule_type;
        cond_format->max_value = user_options->max_value;
        cond_format->max_value_string =
            lxw_strdup_formula(user_options->max_value_string);
    }
    else {
        cond_format->max_rule_type = LXW_CONDITIONAL_RULE_TYPE_MAXIMUM;
        cond_format->max_value = 0;
    }

    if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
        if (mid_rule_type > LXW_CONDITIONAL_RULE_TYPE_MINIMUM
            && mid_rule_type < LXW_CONDITIONAL_RULE_TYPE_MAXIMUM) {
            cond_format->mid_rule_type = mid_rule_type;
            cond_format->mid_value = user_options->mid_value;
            cond_format->mid_value_string =
                lxw_strdup_formula(user_options->mid_value_string);
        }
        else {
            cond_format->mid_rule_type = LXW_CONDITIONAL_RULE_TYPE_PERCENTILE;
            cond_format->mid_value = 50;
        }
    }

    if (user_options->min_color != LXW_COLOR_UNSET)
        cond_format->min_color = user_options->min_color;
    else
        cond_format->min_color = 0xFF7128;

    if (user_options->max_color != LXW_COLOR_UNSET)
        cond_format->max_color = user_options->max_color;
    else
        cond_format->max_color = 0xFFEF9C;

    if (cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {
        if (user_options->min_color == LXW_COLOR_UNSET)
            cond_format->min_color = 0xF8696B;

        if (user_options->mid_color != LXW_COLOR_UNSET)
            cond_format->mid_color = user_options->mid_color;
        else
            cond_format->mid_color = 0xFFEB84;

        if (user_options->max_color == LXW_COLOR_UNSET)
            cond_format->max_color = 0x63BE7B;
    }

    return LXW_NO_ERROR;
}

lxw_error
_validate_conditional_top(lxw_cond_format_obj *cond_format,
                          lxw_conditional_format *user_options)
{
    /* Restrict the range of rank values to Excel's allowed range. */
    if (user_options->criteria ==
        LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT) {
        if (user_options->value < 0.0 || user_options->value > 100.0) {

            LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                             "For type = LXW_CONDITIONAL_TYPE_TOP/BOTTOM, "
                             "top/bottom percent (%g%%) must by in range 0-100",
                             user_options->value);

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }
    else {
        if (user_options->value < 1.0 || user_options->value > 1000.0) {

            LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                             "For type = LXW_CONDITIONAL_TYPE_TOP/BOTTOM, "
                             "top/bottom items (%g) must by in range 1-1000",
                             user_options->value);

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }

    cond_format->min_value = (uint16_t) user_options->value;

    return LXW_NO_ERROR;
}

lxw_error
_validate_conditional_average(lxw_conditional_format *user)
{
    if (user->criteria < LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE ||
        user->criteria > LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW) {

        LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                         "For type = LXW_CONDITIONAL_TYPE_AVERAGE, "
                         "invalid criteria value (%d).", user->criteria);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXW_NO_ERROR;
    }
}

lxw_error
_validate_conditional_time_period(lxw_conditional_format *user)
{
    if (user->criteria < LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY ||
        user->criteria > LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH) {

        LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                         "For type = LXW_CONDITIONAL_TYPE_TIME_PERIOD, "
                         "invalid criteria value (%d).", user->criteria);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXW_NO_ERROR;
    }
}

lxw_error
_validate_conditional_text(lxw_cond_format_obj *cond_format,
                           lxw_conditional_format *user_options)
{
    if (!user_options->value_string) {

        LXW_WARN_FORMAT("worksheet_conditional_format_cell()/_range(): "
                        "For type = LXW_CONDITIONAL_TYPE_TEXT, "
                        "value_string can not be NULL. "
                        "Text must be specified.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (strlen(user_options->value_string) >= LXW_MAX_ATTRIBUTE_LENGTH) {

        LXW_WARN_FORMAT2("worksheet_conditional_format_cell()/_range(): "
                         "For type = LXW_CONDITIONAL_TYPE_TEXT, "
                         "value_string length (%d) must be less than %d.",
                         (uint16_t) strlen(user_options->value_string),
                         LXW_MAX_ATTRIBUTE_LENGTH);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (user_options->criteria < LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING ||
        user_options->criteria > LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH) {

        LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                         "For type = LXW_CONDITIONAL_TYPE_TEXT, "
                         "invalid criteria value (%d).",
                         user_options->criteria);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    cond_format->min_value_string =
        lxw_strdup_formula(user_options->value_string);

    return LXW_NO_ERROR;
}

lxw_error
_validate_conditional_formula(lxw_cond_format_obj *cond_format,
                              lxw_conditional_format *user_options)
{
    if (!user_options->value_string) {

        LXW_WARN_FORMAT("worksheet_conditional_format_cell()/_range(): "
                        "For type = LXW_CONDITIONAL_TYPE_FORMULA, "
                        "value_string can not be NULL. "
                        "Formula must be specified.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    cond_format->min_value_string =
        lxw_strdup_formula(user_options->value_string);

    return LXW_NO_ERROR;
}

lxw_error
_validate_conditional_cell(lxw_cond_format_obj *cond_format,
                           lxw_conditional_format *user_options)
{
    cond_format->min_value = user_options->value;
    cond_format->min_value_string =
        lxw_strdup_formula(user_options->value_string);

    if (cond_format->criteria == LXW_CONDITIONAL_CRITERIA_BETWEEN
        || cond_format->criteria == LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN) {
        cond_format->has_max = LXW_TRUE;
        cond_format->min_value = user_options->min_value;
        cond_format->max_value = user_options->max_value;
        cond_format->min_value_string =
            lxw_strdup_formula(user_options->min_value_string);
        cond_format->max_value_string =
            lxw_strdup_formula(user_options->max_value_string);
    }

    return LXW_NO_ERROR;
}

/* Check that the correct criteria and used with the correct conditional format. */
lxw_error
_validate_conditional_criteria(lxw_cond_format_obj *cond_format)
{
    uint8_t criteria_mismatch = LXW_FALSE;

    if (cond_format->type == LXW_CONDITIONAL_TYPE_CELL) {
        switch (cond_format->criteria) {
            case LXW_CONDITIONAL_CRITERIA_EQUAL_TO:
            case LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO:
            case LXW_CONDITIONAL_CRITERIA_GREATER_THAN:
            case LXW_CONDITIONAL_CRITERIA_LESS_THAN:
            case LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO:
            case LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO:
            case LXW_CONDITIONAL_CRITERIA_BETWEEN:
            case LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN:
                criteria_mismatch = LXW_FALSE;
                break;
            default:
                criteria_mismatch = LXW_TRUE;
        }
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {
        switch (cond_format->criteria) {
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH:
            case LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH:
                criteria_mismatch = LXW_FALSE;
                break;
            default:
                criteria_mismatch = LXW_TRUE;
        }
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TEXT) {
        switch (cond_format->criteria) {
            case LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING:
            case LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING:
            case LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH:
            case LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH:
                criteria_mismatch = LXW_FALSE;
                break;
            default:
                criteria_mismatch = LXW_TRUE;
        }
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_AVERAGE) {
        switch (cond_format->criteria) {
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE:
            case LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW:
                criteria_mismatch = LXW_FALSE;
                break;
            default:
                criteria_mismatch = LXW_TRUE;
        }
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {
        switch (cond_format->criteria) {
            case LXW_CONDITIONAL_CRITERIA_NONE:
            case LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT:
                criteria_mismatch = LXW_FALSE;
                break;
            default:
                criteria_mismatch = LXW_TRUE;
        }
    }
    else {
        /* Any other conditional type should have a zero criteria. */
        cond_format->criteria = LXW_CONDITIONAL_CRITERIA_NONE;
    }

    if (criteria_mismatch) {
        LXW_WARN_FORMAT2("worksheet_conditional_format_cell()/_range(): "
                         "LXW_CONDITIONAL_CRITERIA_* = %d is not valid for "
                         "LXW_CONDITIONAL_TYPE_* = %d", cond_format->criteria,
                         cond_format->type);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }
    else {
        return LXW_NO_ERROR;
    }
}

/*
 * Write the <ignoredErrors> element.
 */
STATIC void
_worksheet_write_ignored_errors(lxw_worksheet *self)
{
    if (!self->has_ignore_errors)
        return;

    lxw_xml_start_tag(self->file, "ignoredErrors", NULL);

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

    lxw_xml_end_tag(self->file, "ignoredErrors");
}

/*
 * Write the <tablePart> element.
 */
STATIC void
_worksheet_write_table_part(lxw_worksheet *self, uint16_t id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "tablePart", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <tableParts> element.
 */
STATIC void
_worksheet_write_table_parts(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_table_obj *table_obj;

    if (!self->table_count)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("count", self->table_count);

    lxw_xml_start_tag(self->file, "tableParts", &attributes);

    STAILQ_FOREACH(table_obj, self->table_objs, list_pointers) {
        self->rel_count++;

        /* Write the tablePart element. */
        _worksheet_write_table_part(self, self->rel_count);
    }

    lxw_xml_end_tag(self->file, "tableParts");

    LXW_FREE_ATTRIBUTES();
}

/*
 * External functions to call intern XML methods shared with chartsheet.
 */
void
lxw_worksheet_write_sheet_views(lxw_worksheet *self)
{
    _worksheet_write_sheet_views(self);
}

void
lxw_worksheet_write_page_margins(lxw_worksheet *self)
{
    _worksheet_write_page_margins(self);
}

void
lxw_worksheet_write_drawings(lxw_worksheet *self)
{
    _worksheet_write_drawings(self);
}

void
lxw_worksheet_write_sheet_protection(lxw_worksheet *self,
                                     lxw_protection_obj *protect)
{
    _worksheet_write_sheet_protection(self, protect);
}

void
lxw_worksheet_write_sheet_pr(lxw_worksheet *self)
{
    _worksheet_write_sheet_pr(self);
}

void
lxw_worksheet_write_page_setup(lxw_worksheet *self)
{
    _worksheet_write_page_setup(self);
}

void
lxw_worksheet_write_header_footer(lxw_worksheet *self)
{
    _worksheet_write_header_footer(self);
}

/*
 * Assemble and write the XML file.
 */
void
lxw_worksheet_assemble_xml_file(lxw_worksheet *self)
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
    lxw_xml_end_tag(self->file, "worksheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Write a number to a cell in Excel.
 */
lxw_error
worksheet_write_number(lxw_worksheet *self,
                       lxw_row_t row_num,
                       lxw_col_t col_num, double value, lxw_format *format)
{
    lxw_cell *cell;
    lxw_error err;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    cell = _new_number_cell(row_num, col_num, value, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a string to an Excel file.
 */
lxw_error
worksheet_write_string(lxw_worksheet *self,
                       lxw_row_t row_num,
                       lxw_col_t col_num, const char *string,
                       lxw_format *format)
{
    lxw_cell *cell;
    int32_t string_id;
    char *string_copy;
    struct sst_element *sst_element;
    lxw_error err;

    if (!string || !*string) {
        /* Treat a NULL or empty string with formatting as a blank cell. */
        /* Null strings without formats should be ignored.      */
        if (format)
            return worksheet_write_blank(self, row_num, col_num, format);
        else
            return LXW_NO_ERROR;
    }

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    if (lxw_utf8_strlen(string) > LXW_STR_MAX)
        return LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED;

    if (!self->optimize) {
        /* Get the SST element and string id. */
        sst_element = lxw_get_sst_index(self->sst, string, LXW_FALSE);

        if (!sst_element)
            return LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id,
                                sst_element->string, format);
    }
    else {
        /* Look for and escape control chars in the string. */
        if (lxw_has_control_characters(string)) {
            string_copy = lxw_escape_control_characters(string);
        }
        else {
            string_copy = lxw_strdup(string);
        }
        cell = _new_inline_string_cell(row_num, col_num, string_copy, format);
    }

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a formula with a numerical result to a cell in Excel.
 */
lxw_error
worksheet_write_formula_num(lxw_worksheet *self,
                            lxw_row_t row_num,
                            lxw_col_t col_num,
                            const char *formula,
                            lxw_format *format, double result)
{
    lxw_cell *cell;
    char *formula_copy;
    lxw_error err;

    if (!formula)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_str_is_empty(formula))
        return LXW_ERROR_PARAMETER_IS_EMPTY;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Strip leading "=" from formula. */
    if (formula[0] == '=')
        formula_copy = lxw_strdup(formula + 1);
    else
        formula_copy = lxw_strdup(formula);

    cell = _new_formula_cell(row_num, col_num, formula_copy, format);
    cell->formula_result = result;

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a formula with a string result to a cell in Excel.
 */
lxw_error
worksheet_write_formula_str(lxw_worksheet *self,
                            lxw_row_t row_num,
                            lxw_col_t col_num,
                            const char *formula,
                            lxw_format *format, const char *result)
{
    lxw_cell *cell;
    char *formula_copy;
    lxw_error err;

    if (!formula)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_str_is_empty(formula))
        return LXW_ERROR_PARAMETER_IS_EMPTY;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Strip leading "=" from formula. */
    if (formula[0] == '=')
        formula_copy = lxw_strdup(formula + 1);
    else
        formula_copy = lxw_strdup(formula);

    cell = _new_formula_cell(row_num, col_num, formula_copy, format);
    cell->user_data2 = lxw_strdup(result);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a formula with a default result to a cell in Excel .
 */
lxw_error
worksheet_write_formula(lxw_worksheet *self,
                        lxw_row_t row_num,
                        lxw_col_t col_num, const char *formula,
                        lxw_format *format)
{
    return worksheet_write_formula_num(self, row_num, col_num, formula,
                                       format, 0);
}

/*
 * Internal shared function for various array formula functions.
 */
lxw_error
_store_array_formula(lxw_worksheet *self,
                     lxw_row_t first_row,
                     lxw_col_t first_col,
                     lxw_row_t last_row,
                     lxw_col_t last_col,
                     const char *formula, lxw_format *format, double result,
                     uint8_t is_dynamic)
{
    lxw_cell *cell;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    char *formula_copy;
    char *range;
    lxw_error err;

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
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_str_is_empty(formula))
        return LXW_ERROR_PARAMETER_IS_EMPTY;

    /* Check that row and col are valid and store max and min values. */
    err = _check_dimensions(self, first_row, first_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Define the array range. */
    range = calloc(1, LXW_MAX_CELL_RANGE_LENGTH);
    RETURN_ON_MEM_ERROR(range, LXW_ERROR_MEMORY_MALLOC_FAILED);

    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(range, first_row, first_col);
    else
        lxw_rowcol_to_range(range, first_row, first_col, last_row, last_col);

    /* Copy and trip leading "{=" from formula. */
    if (formula[0] == '{')
        if (strlen(formula) >= 2 && formula[1] == '=')
            formula_copy = lxw_strdup(formula + 2);
        else
            formula_copy = lxw_strdup(formula + 1);
    else
        formula_copy = lxw_strdup_formula(formula);

    /* Strip trailing "}" from formula. */
    if (strlen(formula_copy) > 0
        && formula_copy[strlen(formula_copy) - 1] == '}') {
        formula_copy[strlen(formula_copy) - 1] = '\0';
    }

    /* Check for empty formula that started as {=}. */
    if (lxw_str_is_empty(formula_copy)) {
        free(formula_copy);
        free(range);
        return LXW_ERROR_PARAMETER_IS_EMPTY;
    }

    /* Create a new array formula cell object. */
    cell = _new_array_formula_cell(first_row, first_col,
                                   formula_copy, range, format, is_dynamic);

    cell->formula_result = result;

    _insert_cell(self, first_row, first_col, cell);

    if (is_dynamic)
        self->has_dynamic_functions = LXW_TRUE;

    /* Pad out the rest of the area with formatted zeroes. */
    if (!self->optimize) {
        for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
            for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
                if (tmp_row == first_row && tmp_col == first_col)
                    continue;

                worksheet_write_number(self, tmp_row, tmp_col, 0, format);
            }
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Write an array formula with a numerical result to a cell in Excel.
 */
lxw_error
worksheet_write_array_formula_num(lxw_worksheet *self,
                                  lxw_row_t first_row,
                                  lxw_col_t first_col,
                                  lxw_row_t last_row,
                                  lxw_col_t last_col,
                                  const char *formula,
                                  lxw_format *format, double result)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, result,
                                LXW_FALSE);
}

/*
 * Write an array formula with a default result to a cell in Excel.
 */
lxw_error
worksheet_write_array_formula(lxw_worksheet *self,
                              lxw_row_t first_row,
                              lxw_col_t first_col,
                              lxw_row_t last_row,
                              lxw_col_t last_col,
                              const char *formula, lxw_format *format)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, 0,
                                LXW_FALSE);
}

/*
 * Write a single cell dynamic array formula with a default result to a cell.
 */
lxw_error
worksheet_write_dynamic_formula(lxw_worksheet *self,
                                lxw_row_t row,
                                lxw_col_t col,
                                const char *formula, lxw_format *format)
{
    return _store_array_formula(self, row, col, row, col, formula, format, 0,
                                LXW_TRUE);
}

/*
 * Write a single cell dynamic array formula with a numerical result to a cell.
 */
lxw_error
worksheet_write_dynamic_formula_num(lxw_worksheet *self,
                                    lxw_row_t row,
                                    lxw_col_t col,
                                    const char *formula,
                                    lxw_format *format, double result)
{
    return _store_array_formula(self, row, col, row, col, formula, format,
                                result, LXW_TRUE);
}

/*
 * Write a dynamic array formula with a numerical result to a cell in Excel.
 */
lxw_error
worksheet_write_dynamic_array_formula_num(lxw_worksheet *self,
                                          lxw_row_t first_row,
                                          lxw_col_t first_col,
                                          lxw_row_t last_row,
                                          lxw_col_t last_col,
                                          const char *formula,
                                          lxw_format *format, double result)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, result,
                                LXW_TRUE);
}

/*
 * Write a dynamic array formula with a default result to a cell in Excel.
 */
lxw_error
worksheet_write_dynamic_array_formula(lxw_worksheet *self,
                                      lxw_row_t first_row,
                                      lxw_col_t first_col,
                                      lxw_row_t last_row,
                                      lxw_col_t last_col,
                                      const char *formula, lxw_format *format)
{
    return _store_array_formula(self, first_row, first_col,
                                last_row, last_col, formula, format, 0,
                                LXW_TRUE);
}

/*
 * Write a blank cell with a format to a cell in Excel.
 */
lxw_error
worksheet_write_blank(lxw_worksheet *self,
                      lxw_row_t row_num, lxw_col_t col_num,
                      lxw_format *format)
{
    lxw_cell *cell;
    lxw_error err;

    /* Blank cells without formatting are ignored by Excel. */
    if (!format)
        return LXW_NO_ERROR;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    cell = _new_blank_cell(row_num, col_num, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a boolean cell with a format to a cell in Excel.
 */
lxw_error
worksheet_write_boolean(lxw_worksheet *self,
                        lxw_row_t row_num, lxw_col_t col_num,
                        int value, lxw_format *format)
{
    lxw_cell *cell;
    lxw_error err;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    cell = _new_boolean_cell(row_num, col_num, value, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a date and or time to a cell in Excel.
 */
lxw_error
worksheet_write_datetime(lxw_worksheet *self,
                         lxw_row_t row_num,
                         lxw_col_t col_num, lxw_datetime *datetime,
                         lxw_format *format)
{
    lxw_cell *cell;
    double excel_date;
    lxw_error err;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    err = lxw_datetime_validate(datetime);
    if (err)
        return err;

    excel_date =
        lxw_datetime_to_excel_date_with_epoch(datetime, self->use_1904_epoch);

    cell = _new_number_cell(row_num, col_num, excel_date, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a date and or time to a cell in Excel.
 */
lxw_error
worksheet_write_unixtime(lxw_worksheet *self,
                         lxw_row_t row_num,
                         lxw_col_t col_num,
                         int64_t unixtime, lxw_format *format)
{
    lxw_cell *cell;
    double excel_date;
    lxw_error err;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    excel_date =
        lxw_unixtime_to_excel_date_with_epoch(unixtime, self->use_1904_epoch);

    cell = _new_number_cell(row_num, col_num, excel_date, format);

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a hyperlink/url to an Excel file.
 */
lxw_error
worksheet_write_url_opt(lxw_worksheet *self,
                        lxw_row_t row_num,
                        lxw_col_t col_num, const char *url,
                        lxw_format *user_format, const char *string,
                        const char *tooltip)
{
    lxw_cell *link;
    char *string_copy = NULL;
    char *url_copy = NULL;
    char *url_external = NULL;
    char *url_string = NULL;
    char *tooltip_copy = NULL;
    char *found_string;
    char *tmp_string = NULL;
    lxw_format *format = NULL;
    size_t string_size;
    size_t i;
    lxw_error err = LXW_ERROR_MEMORY_MALLOC_FAILED;
    enum cell_types link_type = HYPERLINK_URL;

    if (!url || !*url)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* Check the Excel limit of URLS per worksheet. */
    if (self->hlink_count > LXW_MAX_NUMBER_URLS) {
        LXW_WARN("worksheet_write_url()/_opt(): URL ignored since it exceeds "
                 "the maximum number of allowed worksheet URLs (65530).");
        return LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED;
    }

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Reset default error condition. */
    err = LXW_ERROR_MEMORY_MALLOC_FAILED;

    /* Set the URI scheme from internal links. */
    found_string = strstr(url, "internal:");
    if (found_string)
        link_type = HYPERLINK_INTERNAL;

    /* Set the URI scheme from external links. */
    found_string = strstr(url, "external:");
    if (found_string)
        link_type = HYPERLINK_EXTERNAL;

    if (string) {
        string_copy = lxw_strdup(string);
        GOTO_LABEL_ON_MEM_ERROR(string_copy, mem_error);
    }
    else {
        if (link_type == HYPERLINK_URL) {
            /* Strip the mailto header. */
            found_string = strstr(url, "mailto:");
            if (found_string)
                string_copy = lxw_strdup(url + sizeof("mailto"));
            else
                string_copy = lxw_strdup(url);
        }
        else {
            string_copy = lxw_strdup(url + sizeof("__ternal"));
        }
        GOTO_LABEL_ON_MEM_ERROR(string_copy, mem_error);
    }

    if (url) {
        if (link_type == HYPERLINK_URL)
            url_copy = lxw_strdup(url);
        else
            url_copy = lxw_strdup(url + sizeof("__ternal"));

        GOTO_LABEL_ON_MEM_ERROR(url_copy, mem_error);
    }

    if (tooltip) {
        tooltip_copy = lxw_strdup(tooltip);
        GOTO_LABEL_ON_MEM_ERROR(tooltip_copy, mem_error);
    }

    if (link_type == HYPERLINK_INTERNAL) {
        url_string = lxw_strdup(string_copy);
        GOTO_LABEL_ON_MEM_ERROR(url_string, mem_error);
    }

    /* Split url into the link and optional anchor/location. */
    found_string = strchr(url_copy, '#');

    if (found_string) {
        free(url_string);
        url_string = lxw_strdup(found_string + 1);
        GOTO_LABEL_ON_MEM_ERROR(url_string, mem_error);

        *found_string = '\0';
    }

    /* Escape the URL. */
    if (link_type == HYPERLINK_URL || link_type == HYPERLINK_EXTERNAL) {
        tmp_string = lxw_escape_url_characters(url_copy, LXW_FALSE);
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

            lxw_snprintf(url_external, string_size, "file:///%s", url_copy);

        }

        /* Convert a ./dir/file.xlsx link to dir/file.xlsx. */
        found_string = strstr(url_copy, ".\\");
        if (found_string == url_copy)
            memmove(url_copy, url_copy + 2, strlen(url_copy) - 1);

        if (url_external) {
            free(url_copy);
            url_copy = lxw_strdup(url_external);
            GOTO_LABEL_ON_MEM_ERROR(url_copy, mem_error);

            free(url_external);
            url_external = NULL;
        }

    }

    /* Check if URL exceeds Excel's length limit. */
    if (lxw_utf8_strlen(url_copy) > self->max_url_length) {
        LXW_WARN_FORMAT2("worksheet_write_url()/_opt(): URL exceeds "
                         "Excel's allowable length of %d characters: %s",
                         self->max_url_length, url_copy);
        err = LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED;
        goto mem_error;
    }

    /* Use the default URL format if none is specified. */
    if (!user_format)
        format = self->default_url_format;
    else
        format = user_format;

    if (!self->storing_embedded_image) {
        err =
            worksheet_write_string(self, row_num, col_num, string_copy,
                                   format);
        if (err)
            goto mem_error;
    }

    /* Reset default error condition. */
    err = LXW_ERROR_MEMORY_MALLOC_FAILED;

    link = _new_hyperlink_cell(row_num, col_num, link_type, url_copy,
                               url_string, tooltip_copy);
    GOTO_LABEL_ON_MEM_ERROR(link, mem_error);

    _insert_hyperlink(self, row_num, col_num, link);

    free(string_copy);
    self->hlink_count++;
    return LXW_NO_ERROR;

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
lxw_error
worksheet_write_url(lxw_worksheet *self,
                    lxw_row_t row_num,
                    lxw_col_t col_num, const char *url, lxw_format *format)
{
    return worksheet_write_url_opt(self, row_num, col_num, url, format, NULL,
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
lxw_error
worksheet_write_rich_string(lxw_worksheet *self,
                            lxw_row_t row_num,
                            lxw_col_t col_num,
                            lxw_rich_string_tuple *rich_strings[],
                            lxw_format *format)
{
    lxw_cell *cell;
    int32_t string_id;
    struct sst_element *sst_element;
    lxw_error err;
    uint8_t i;
    long file_size;
    char *rich_string = NULL;
    const char *string_copy = NULL;
    lxw_styles *styles = NULL;
    lxw_format *default_format = NULL;
    lxw_rich_string_tuple *rich_string_tuple = NULL;
    FILE *tmpfile;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Iterate through rich string fragments to check for input errors. */
    i = 0;
    err = LXW_NO_ERROR;
    while ((rich_string_tuple = rich_strings[i++]) != NULL) {

        /* Check for NULL or empty strings. */
        if (!rich_string_tuple->string || !*rich_string_tuple->string) {
            err = LXW_ERROR_PARAMETER_VALIDATION;
        }
    }

    /* If there are less than 2 fragments it isn't a rich string. */
    if (i <= 2)
        err = LXW_ERROR_PARAMETER_VALIDATION;

    if (err)
        return err;

    /* Create a tmp file for the styles object. */
    tmpfile = lxw_get_filehandle(&rich_string, NULL, self->tmpdir);
    if (!tmpfile)
        return LXW_ERROR_CREATING_TMPFILE;

    /* Create a temp styles object for writing the font data. */
    styles = lxw_styles_new();
    GOTO_LABEL_ON_MEM_ERROR(styles, mem_error);
    styles->file = tmpfile;

    /* Create a default format for non-formatted text. */
    default_format = lxw_format_new();
    GOTO_LABEL_ON_MEM_ERROR(default_format, mem_error);

    /* Iterate through the rich string fragments and write each one out. */
    i = 0;
    while ((rich_string_tuple = rich_strings[i++]) != NULL) {
        lxw_xml_start_tag(tmpfile, "r", NULL);

        if (rich_string_tuple->format) {
            /* Write the user defined font format. */
            lxw_styles_write_rich_font(styles, rich_string_tuple->format);
        }
        else {
            /* Write a default font format. Except for the first fragment. */
            if (i > 1)
                lxw_styles_write_rich_font(styles, default_format);
        }

        lxw_styles_write_string_fragment(styles, rich_string_tuple->string);
        lxw_xml_end_tag(tmpfile, "r");
    }

    /* Free the temp objects. */
    lxw_styles_free(styles);
    lxw_format_free(default_format);

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
            return LXW_ERROR_READING_TMPFILE;
        }
    }

    /* Close the temp file. */
    fclose(tmpfile);

    if (lxw_utf8_strlen(rich_string) > LXW_STR_MAX) {
        free((void *) rich_string);
        return LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED;
    }

    if (!self->optimize) {
        /* Get the SST element and string id. */
        sst_element = lxw_get_sst_index(self->sst, rich_string, LXW_TRUE);
        free((void *) rich_string);

        if (!sst_element)
            return LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id,
                                sst_element->string, format);
    }
    else {
        /* Look for and escape control chars in the string. */
        if (lxw_has_control_characters(rich_string)) {
            string_copy = lxw_escape_control_characters(rich_string);
            free((void *) rich_string);
        }
        else {
            string_copy = rich_string;
        }
        cell = _new_inline_rich_string_cell(row_num, col_num, string_copy,
                                            format);
    }

    _insert_cell(self, row_num, col_num, cell);

    return LXW_NO_ERROR;

mem_error:
    lxw_styles_free(styles);
    lxw_format_free(default_format);
    fclose(tmpfile);

    return LXW_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Write a comment to a worksheet cell in Excel.
 */
lxw_error
worksheet_write_comment_opt(lxw_worksheet *self,
                            lxw_row_t row_num, lxw_col_t col_num,
                            const char *text, lxw_comment_options *options)
{
    lxw_cell *cell;
    lxw_error err;
    lxw_vml_obj *comment;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    if (!text)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_str_is_empty(text))
        return LXW_ERROR_PARAMETER_IS_EMPTY;

    if (lxw_utf8_strlen(text) > LXW_STR_MAX)
        return LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED;

    comment = calloc(1, sizeof(lxw_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(comment, mem_error);

    comment->text = lxw_strdup(text);
    GOTO_LABEL_ON_MEM_ERROR(comment->text, mem_error);

    comment->row = row_num;
    comment->col = col_num;

    cell = _new_comment_cell(row_num, col_num, comment);
    GOTO_LABEL_ON_MEM_ERROR(cell, mem_error);

    _insert_comment(self, row_num, col_num, cell);

    /* Set user and default parameters for the comment. */
    _get_comment_params(comment, options);

    self->has_vml = LXW_TRUE;
    self->has_comments = LXW_TRUE;

    /* Insert a placeholder in the cell RB table in the same position so
     * that the worksheet row "spans" calculations are correct. */
    _insert_cell_placeholder(self, row_num, col_num);

    return LXW_NO_ERROR;

mem_error:
    if (comment)
        _free_vml_object(comment);

    return LXW_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Write a comment to a worksheet cell in Excel.
 */
lxw_error
worksheet_write_comment(lxw_worksheet *self,
                        lxw_row_t row_num, lxw_col_t col_num,
                        const char *string)
{
    return worksheet_write_comment_opt(self, row_num, col_num, string, NULL);
}

/*
 * Set the properties of a single column or a range of columns with options.
 */
lxw_error
worksheet_set_column_opt(lxw_worksheet *self,
                         lxw_col_t firstcol,
                         lxw_col_t lastcol,
                         double width,
                         lxw_format *format,
                         lxw_row_col_options *user_options)
{
    lxw_col_options *copied_options;
    uint8_t ignore_row = LXW_TRUE;
    uint8_t ignore_col = LXW_TRUE;
    uint8_t hidden = LXW_FALSE;
    uint8_t level = 0;
    uint8_t collapsed = LXW_FALSE;
    lxw_col_t col;
    lxw_error err;

    if (user_options) {
        hidden = user_options->hidden;
        level = user_options->level;
        collapsed = user_options->collapsed;
    }

    /* Ensure second col is larger than first. */
    if (firstcol > lastcol) {
        lxw_col_t tmp = firstcol;
        firstcol = lastcol;
        lastcol = tmp;
    }

    /* Ensure that the cols are valid and store max and min values.
     * NOTE: The check shouldn't modify the row dimensions and should only
     *       modify the column dimensions in certain cases. */
    if (format != NULL || (width != LXW_DEF_COL_WIDTH && hidden))
        ignore_col = LXW_FALSE;

    err = _check_dimensions(self, 0, firstcol, ignore_row, ignore_col);

    if (!err)
        err = _check_dimensions(self, 0, lastcol, ignore_row, ignore_col);

    if (err)
        return err;

    /* Resize the col_options array if required. */
    if (firstcol >= self->col_options_max) {
        lxw_col_t col_tmp;
        lxw_col_t old_size = self->col_options_max;
        lxw_col_t new_size = _next_power_of_two(firstcol + 1);
        lxw_col_options **new_ptr = realloc(self->col_options,
                                            new_size *
                                            sizeof(lxw_col_options *));

        if (new_ptr) {
            for (col_tmp = old_size; col_tmp < new_size; col_tmp++)
                new_ptr[col_tmp] = NULL;

            self->col_options = new_ptr;
            self->col_options_max = new_size;
        }
        else {
            return LXW_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Resize the col_formats array if required. */
    if (lastcol >= self->col_formats_max) {
        lxw_col_t col;
        lxw_col_t old_size = self->col_formats_max;
        lxw_col_t new_size = _next_power_of_two(lastcol + 1);
        lxw_format **new_ptr = realloc(self->col_formats,
                                       new_size * sizeof(lxw_format *));

        if (new_ptr) {
            for (col = old_size; col < new_size; col++)
                new_ptr[col] = NULL;

            self->col_formats = new_ptr;
            self->col_formats_max = new_size;
        }
        else {
            return LXW_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Store the column options. */
    copied_options = calloc(1, sizeof(lxw_col_options));
    RETURN_ON_MEM_ERROR(copied_options, LXW_ERROR_MEMORY_MALLOC_FAILED);

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
    self->col_size_changed = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set the properties of a single column or a range of columns.
 */
lxw_error
worksheet_set_column(lxw_worksheet *self,
                     lxw_col_t firstcol,
                     lxw_col_t lastcol, double width, lxw_format *format)
{
    return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    NULL);
}

/*
 * Set the properties of a single column or a range of columns, with the
 * width in pixels.
 */
lxw_error
worksheet_set_column_pixels(lxw_worksheet *self,
                            lxw_col_t firstcol,
                            lxw_col_t lastcol,
                            uint32_t pixels, lxw_format *format)
{
    double width = _pixels_to_width(pixels);

    return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    NULL);
}

/*
 * Set the properties of a single column or a range of columns with options,
 * with the width in pixels.
 */
lxw_error
worksheet_set_column_pixels_opt(lxw_worksheet *self,
                                lxw_col_t firstcol,
                                lxw_col_t lastcol,
                                uint32_t pixels,
                                lxw_format *format,
                                lxw_row_col_options *user_options)
{
    double width = _pixels_to_width(pixels);

    return worksheet_set_column_opt(self, firstcol, lastcol, width, format,
                                    user_options);
}

/*
 * Set the properties of a row with options.
 */
lxw_error
worksheet_set_row_opt(lxw_worksheet *self,
                      lxw_row_t row_num,
                      double height,
                      lxw_format *format, lxw_row_col_options *user_options)
{

    lxw_col_t min_col;
    uint8_t hidden = LXW_FALSE;
    uint8_t level = 0;
    uint8_t collapsed = LXW_FALSE;
    lxw_row *row;
    lxw_error err;

    if (user_options) {
        hidden = user_options->hidden;
        level = user_options->level;
        collapsed = user_options->collapsed;
    }

    /* Use minimum col in _check_dimensions(). */
    if (self->dim_colmin != LXW_COL_MAX)
        min_col = self->dim_colmin;
    else
        min_col = 0;

    err = _check_dimensions(self, row_num, min_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* If the height is 0 the row is hidden and the height is the default. */
    if (height == 0) {
        hidden = LXW_TRUE;
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
    row->row_changed = LXW_TRUE;

    if (height != self->default_row_height)
        row->height_changed = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set the properties of a row.
 */
lxw_error
worksheet_set_row(lxw_worksheet *self,
                  lxw_row_t row_num, double height, lxw_format *format)
{
    return worksheet_set_row_opt(self, row_num, height, format, NULL);
}

/*
 * Set the properties of a row, with the height in pixels.
 */
lxw_error
worksheet_set_row_pixels(lxw_worksheet *self,
                         lxw_row_t row_num, uint32_t pixels,
                         lxw_format *format)
{
    double height = _pixels_to_height(pixels);

    return worksheet_set_row_opt(self, row_num, height, format, NULL);
}

/*
 * Set the properties of a row with options, with the height in pixels.
 */
lxw_error
worksheet_set_row_pixels_opt(lxw_worksheet *self,
                             lxw_row_t row_num,
                             uint32_t pixels,
                             lxw_format *format,
                             lxw_row_col_options *user_options)
{
    double height = _pixels_to_height(pixels);

    return worksheet_set_row_opt(self, row_num, height, format, user_options);
}

/*
 * Merge a range of cells. The first cell should contain the data and the others
 * should be blank. All cells should contain the same format.
 */
lxw_error
worksheet_merge_range(lxw_worksheet *self, lxw_row_t first_row,
                      lxw_col_t first_col, lxw_row_t last_row,
                      lxw_col_t last_col, const char *string,
                      lxw_format *format)
{
    lxw_merged_range *merged_range;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;

    /* Excel doesn't allow a single cell to be merged */
    if (first_row == last_row && first_col == last_col)
        return LXW_ERROR_PARAMETER_VALIDATION;

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
    err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Store the merge range. */
    merged_range = calloc(1, sizeof(lxw_merged_range));
    RETURN_ON_MEM_ERROR(merged_range, LXW_ERROR_MEMORY_MALLOC_FAILED);

    merged_range->first_row = first_row;
    merged_range->first_col = first_col;
    merged_range->last_row = last_row;
    merged_range->last_col = last_col;

    STAILQ_INSERT_TAIL(self->merged_ranges, merged_range, list_pointers);
    self->merged_range_count++;

    /* Write the first cell */
    worksheet_write_string(self, first_row, first_col, string, format);

    /* Pad out the rest of the area with formatted blank cells. */
    for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
        for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
            if (tmp_row == first_row && tmp_col == first_col)
                continue;

            worksheet_write_blank(self, tmp_row, tmp_col, format);
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Set the autofilter area in the worksheet.
 */
lxw_error
worksheet_autofilter(lxw_worksheet *self, lxw_row_t first_row,
                     lxw_col_t first_col, lxw_row_t last_row,
                     lxw_col_t last_col)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;
    lxw_filter_rule_obj **filter_rules;
    lxw_col_t num_filter_rules;

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
    err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Create a array to hold filter rules. */
    self->autofilter.in_use = LXW_FALSE;
    self->autofilter.has_rules = LXW_FALSE;
    _free_filter_rules(self);
    num_filter_rules = last_col - first_col + 1;
    filter_rules = calloc(num_filter_rules, sizeof(lxw_filter_rule_obj *));
    RETURN_ON_MEM_ERROR(filter_rules, LXW_ERROR_MEMORY_MALLOC_FAILED);

    self->autofilter.in_use = LXW_TRUE;
    self->autofilter.first_row = first_row;
    self->autofilter.first_col = first_col;
    self->autofilter.last_row = last_row;
    self->autofilter.last_col = last_col;

    self->filter_rules = filter_rules;
    self->num_filter_rules = num_filter_rules;

    return LXW_NO_ERROR;
}

/*
 * Set a autofilter rule for a filter column.
 */
lxw_error
worksheet_filter_column(lxw_worksheet *self, lxw_col_t col,
                        lxw_filter_rule *rule)
{
    lxw_filter_rule_obj *rule_obj;
    uint16_t rule_index;

    if (!rule) {
        LXW_WARN("worksheet_filter_column(): rule parameter cannot be NULL");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXW_FALSE) {
        LXW_WARN("worksheet_filter_column(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use worksheet_autofilter() first.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXW_WARN_FORMAT3("worksheet_filter_column(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);

    rule_obj->col_num = rule_index;
    rule_obj->type = LXW_FILTER_TYPE_SINGLE;
    rule_obj->criteria1 = rule->criteria;
    rule_obj->value1 = rule->value;

    if (rule_obj->criteria1 != LXW_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value1_string = lxw_strdup(rule->value_string);
    }
    else {
        rule_obj->criteria1 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value1_string = lxw_strdup(" ");
    }

    if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXW_TRUE;

    _set_custom_filter(rule_obj);

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXW_TRUE;
    self->autofilter.has_rules = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set two autofilter rules for a filter column.
 */
lxw_error
worksheet_filter_column2(lxw_worksheet *self, lxw_col_t col,
                         lxw_filter_rule *rule1, lxw_filter_rule *rule2,
                         uint8_t and_or)
{
    lxw_filter_rule_obj *rule_obj;
    uint16_t rule_index;

    if (!rule1 || !rule2) {
        LXW_WARN("worksheet_filter_column2(): rule parameter cannot be NULL");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXW_FALSE) {
        LXW_WARN("worksheet_filter_column2(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use worksheet_autofilter() first.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXW_WARN_FORMAT3("worksheet_filter_column2(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);

    if (and_or == LXW_FILTER_AND)
        rule_obj->type = LXW_FILTER_TYPE_AND;
    else
        rule_obj->type = LXW_FILTER_TYPE_OR;

    rule_obj->col_num = rule_index;

    rule_obj->criteria1 = rule1->criteria;
    rule_obj->value1 = rule1->value;

    rule_obj->criteria2 = rule2->criteria;
    rule_obj->value2 = rule2->value;

    if (rule_obj->criteria1 != LXW_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value1_string = lxw_strdup(rule1->value_string);
    }
    else {
        rule_obj->criteria1 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value1_string = lxw_strdup(" ");
    }

    if (rule_obj->criteria2 != LXW_FILTER_CRITERIA_NON_BLANKS) {
        rule_obj->value2_string = lxw_strdup(rule2->value_string);
    }
    else {
        rule_obj->criteria2 = LXW_FILTER_CRITERIA_NOT_EQUAL_TO;
        rule_obj->value2_string = lxw_strdup(" ");
    }

    if (rule_obj->criteria1 == LXW_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXW_TRUE;

    if (rule_obj->criteria2 == LXW_FILTER_CRITERIA_BLANKS)
        rule_obj->has_blanks = LXW_TRUE;

    _set_custom_filter(rule_obj);

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXW_TRUE;
    self->autofilter.has_rules = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set two autofilter rules for a filter column.
 */
lxw_error
worksheet_filter_list(lxw_worksheet *self, lxw_col_t col, const char **list)
{
    lxw_filter_rule_obj *rule_obj;
    uint16_t rule_index;
    uint8_t has_blanks = LXW_FALSE;
    uint16_t num_filters = 0;
    uint16_t input_list_index;
    uint16_t rule_obj_list_index;
    const char *str;
    char **tmp_list;

    if (!list) {
        LXW_WARN("worksheet_filter_list(): list parameter cannot be NULL");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (self->autofilter.in_use == LXW_FALSE) {
        LXW_WARN("worksheet_filter_list(): "
                 "Worksheet autofilter range hasn't been defined. "
                 "Use worksheet_autofilter() first.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    if (col < self->autofilter.first_col || col > self->autofilter.last_col) {
        LXW_WARN_FORMAT3("worksheet_filter_list(): "
                         "Column '%d' is outside autofilter range "
                         "'%d <= col <= %d'.", col,
                         self->autofilter.first_col,
                         self->autofilter.last_col);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Count the number of non "Blanks" strings in the input list. */
    input_list_index = 0;
    while ((str = list[input_list_index]) != NULL) {
        if (strncmp(str, "Blanks", 6) == 0)
            has_blanks = LXW_TRUE;
        else
            num_filters++;

        input_list_index++;
    }

    /* There should be at least one filter string. */
    if (num_filters == 0) {
        LXW_WARN("worksheet_filter_list(): "
                 "list must have at least 1 non-blanks item.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Free any previous rule in the column slot. */
    rule_index = col - self->autofilter.first_col;
    _free_filter_rule(self->filter_rules[rule_index]);

    /* Create a new rule and copy user input. */
    rule_obj = calloc(1, sizeof(lxw_filter_rule_obj));
    RETURN_ON_MEM_ERROR(rule_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);

    tmp_list = calloc(num_filters + 1, sizeof(char *));
    GOTO_LABEL_ON_MEM_ERROR(tmp_list, mem_error);

    /* Copy input list (without any "Blanks" command) to an internal list. */
    input_list_index = 0;
    rule_obj_list_index = 0;
    while ((str = list[input_list_index]) != NULL) {
        if (strncmp(str, "Blanks", 6) != 0) {
            tmp_list[rule_obj_list_index] = lxw_strdup(str);
            rule_obj_list_index++;
        }

        input_list_index++;
    }

    rule_obj->list = tmp_list;
    rule_obj->num_list_filters = num_filters;
    rule_obj->is_custom = LXW_FALSE;
    rule_obj->col_num = rule_index;
    rule_obj->type = LXW_FILTER_TYPE_STRING_LIST;
    rule_obj->has_blanks = has_blanks;

    self->filter_rules[rule_index] = rule_obj;
    self->filter_on = LXW_TRUE;
    self->autofilter.has_rules = LXW_TRUE;

    return LXW_NO_ERROR;

mem_error:
    free(rule_obj);
    return LXW_ERROR_MEMORY_MALLOC_FAILED;

}

/*
 * Add an Excel table to the worksheet.
 */
lxw_error
worksheet_add_table(lxw_worksheet *self, lxw_row_t first_row,
                    lxw_col_t first_col, lxw_row_t last_row,
                    lxw_col_t last_col, lxw_table_options *user_options)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_col_t num_cols;
    lxw_error err;
    lxw_table_obj *table_obj;
    lxw_table_column **columns;

    if (self->optimize) {
        LXW_WARN_FORMAT("worksheet_add_table(): "
                        "worksheet tables aren't supported in "
                        "'constant_memory' mode");
        return LXW_ERROR_FEATURE_NOT_SUPPORTED;
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
    err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
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
    table_obj = calloc(1, sizeof(lxw_table_obj));
    RETURN_ON_MEM_ERROR(table_obj, LXW_ERROR_MEMORY_MALLOC_FAILED);

    columns = calloc(num_cols, sizeof(lxw_table_column *));
    GOTO_LABEL_ON_MEM_ERROR(columns, error);

    table_obj->columns = columns;
    table_obj->num_cols = num_cols;
    table_obj->first_row = first_row;
    table_obj->first_col = first_col;
    table_obj->last_row = last_row;
    table_obj->last_col = last_col;

    err = _set_default_table_columns(table_obj);
    if (err)
        goto error;

    /* Create the table range. */
    lxw_rowcol_to_range(table_obj->sqref,
                        first_row, first_col, last_row, last_col);
    lxw_rowcol_to_range(table_obj->filter_sqref,
                        first_row, first_col, last_row, last_col);

    /* Validate and copy user options to an internal object. */
    if (user_options) {

        _check_and_copy_table_style(table_obj, user_options);

        table_obj->total_row = user_options->total_row;
        table_obj->last_column = user_options->last_column;
        table_obj->first_column = user_options->first_column;
        table_obj->no_autofilter = user_options->no_autofilter;
        table_obj->no_header_row = user_options->no_header_row;
        table_obj->no_banded_rows = user_options->no_banded_rows;
        table_obj->banded_columns = user_options->banded_columns;

        if (user_options->no_header_row)
            table_obj->no_autofilter = LXW_TRUE;

        if (user_options->columns) {
            err = _set_custom_table_columns(table_obj, user_options);
            if (err)
                goto error;
        }

        if (user_options->total_row) {
            lxw_rowcol_to_range(table_obj->filter_sqref,
                                first_row, first_col, last_row - 1, last_col);
        }

        if (user_options->name) {
            table_obj->name = lxw_strdup(user_options->name);
            if (!table_obj->name) {
                err = LXW_ERROR_MEMORY_MALLOC_FAILED;
                goto error;
            }
        }
    }

    _write_table_column_data(self, table_obj);

    STAILQ_INSERT_TAIL(self->table_objs, table_obj, list_pointers);
    self->table_count++;

    return LXW_NO_ERROR;

error:
    _free_worksheet_table(table_obj);
    return err;

}

/*
 * Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
void
worksheet_select(lxw_worksheet *self)
{
    self->selected = LXW_TRUE;

    /* Selected worksheet can't be hidden. */
    self->hidden = LXW_FALSE;
}

/*
 * Set this worksheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
void
worksheet_activate(lxw_worksheet *self)
{
    self->selected = LXW_TRUE;
    self->active = LXW_TRUE;

    /* Active worksheet can't be hidden. */
    self->hidden = LXW_FALSE;

    *self->active_sheet = self->index;
}

/*
 * Set this worksheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
void
worksheet_set_first_sheet(lxw_worksheet *self)
{
    /* Active worksheet can't be hidden. */
    self->hidden = LXW_FALSE;

    *self->first_sheet = self->index;
}

/*
 * Hide this worksheet.
 */
void
worksheet_hide(lxw_worksheet *self)
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
 * Set which cell or cells are selected in a worksheet.
 */
lxw_error
worksheet_set_selection(lxw_worksheet *self,
                        lxw_row_t first_row, lxw_col_t first_col,
                        lxw_row_t last_row, lxw_col_t last_col)
{
    lxw_selection *selection;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;
    char active_cell[LXW_MAX_CELL_RANGE_LENGTH];
    char sqref[LXW_MAX_CELL_RANGE_LENGTH];

    /* Only allow selection to be set once to avoid freeing/re-creating it. */
    if (!STAILQ_EMPTY(self->selections))
        return LXW_ERROR_PARAMETER_VALIDATION;

    /* Excel doesn't set a selection for cell A1 since it is the default. */
    if (first_row == 0 && first_col == 0 && last_row == 0 && last_col == 0)
        return LXW_NO_ERROR;

    selection = calloc(1, sizeof(lxw_selection));
    RETURN_ON_MEM_ERROR(selection, LXW_ERROR_MEMORY_MALLOC_FAILED);

    /* Check that row and col are valid without storing. */
    err = _check_dimensions(self, first_row, first_col, LXW_TRUE, LXW_TRUE);
    if (err) {
        free(selection);
        return err;
    }

    err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
    if (err) {
        free(selection);
        return err;
    }

    /* Set the cell range selection. Do this before swapping max/min to  */
    /* allow the selection direction to be reversed. */
    lxw_rowcol_to_cell(active_cell, first_row, first_col);

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
        lxw_rowcol_to_cell(sqref, first_row, first_col);
    else
        lxw_rowcol_to_range(sqref, first_row, first_col, last_row, last_col);

    lxw_strcpy(selection->pane, "");
    lxw_strcpy(selection->active_cell, active_cell);
    lxw_strcpy(selection->sqref, sqref);

    STAILQ_INSERT_TAIL(self->selections, selection, list_pointers);

    return LXW_NO_ERROR;
}

/*
 * Set the first visible cell at the top left of the worksheet.
 */
void
worksheet_set_top_left_cell(lxw_worksheet *self, lxw_row_t row, lxw_col_t col)
{
    if (row == 0 && col == 0)
        return;

    lxw_rowcol_to_cell(self->top_left_cell, row, col);
}

/*
 * Set panes and mark them as frozen. With extra options.
 */
void
worksheet_freeze_panes_opt(lxw_worksheet *self,
                           lxw_row_t first_row, lxw_col_t first_col,
                           lxw_row_t top_row, lxw_col_t left_col,
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
worksheet_freeze_panes(lxw_worksheet *self,
                       lxw_row_t first_row, lxw_col_t first_col)
{
    worksheet_freeze_panes_opt(self, first_row, first_col,
                               first_row, first_col, 0);
}

/*
 * Set panes and mark them as split.With extra options.
 */
void
worksheet_split_panes_opt(lxw_worksheet *self,
                          double y_split, double x_split,
                          lxw_row_t top_row, lxw_col_t left_col)
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
worksheet_split_panes(lxw_worksheet *self, double y_split, double x_split)
{
    worksheet_split_panes_opt(self, y_split, x_split, 0, 0);
}

/*
 * Set the page orientation as portrait.
 */
void
worksheet_set_portrait(lxw_worksheet *self)
{
    self->orientation = LXW_PORTRAIT;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Set the page orientation as landscape.
 */
void
worksheet_set_landscape(lxw_worksheet *self)
{
    self->orientation = LXW_LANDSCAPE;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Set the page view mode for Mac Excel.
 */
void
worksheet_set_page_view(lxw_worksheet *self)
{
    self->page_view = LXW_TRUE;
}

/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
void
worksheet_set_paper(lxw_worksheet *self, uint8_t paper_size)
{
    if (paper_size > 118) {
        LXW_WARN_FORMAT1("worksheet_set_paper(): invalid paper size: %d. "
                         "Valid range is 0-118", paper_size);
        return;
    }

    self->paper_size = paper_size;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Set the order in which pages are printed.
 */
void
worksheet_print_across(lxw_worksheet *self)
{
    self->page_order = LXW_PRINT_ACROSS;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Set all the page margins in inches.
 */
void
worksheet_set_margins(lxw_worksheet *self, double left, double right,
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
lxw_error
worksheet_set_header_opt(lxw_worksheet *self, const char *string,
                         lxw_header_footer_options *options)
{
    lxw_error err;
    char *tmp_header;
    char *found_string;
    char *offset_string;
    uint8_t placeholder_count = 0;
    uint8_t image_count = 0;

    if (!string) {
        LXW_WARN_FORMAT("worksheet_set_header_opt/footer_opt(): "
                        "header/footer string cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxw_utf8_strlen(string) > LXW_HEADER_FOOTER_MAX) {
        LXW_WARN_FORMAT("worksheet_set_header_opt/footer_opt(): "
                        "header/footer string exceeds Excel's limit of "
                        "255 characters.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    tmp_header = lxw_strdup(string);
    RETURN_ON_MEM_ERROR(tmp_header, LXW_ERROR_MEMORY_MALLOC_FAILED);

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
        LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                         "the number of &G/&[Picture] placeholders in option "
                         "string \"%s\" does not match the number of supplied "
                         "images.", string);

        free(tmp_header);
        return LXW_ERROR_PARAMETER_VALIDATION;
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
            LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                             "the number of &G/&[Picture] placeholders in option "
                             "string \"%s\" does not match the number of supplied "
                             "images.", string);

            free(tmp_header);
            return LXW_ERROR_PARAMETER_VALIDATION;
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
    self->header_footer_changed = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set the page footer caption and options.
 */
lxw_error
worksheet_set_footer_opt(lxw_worksheet *self, const char *string,
                         lxw_header_footer_options *options)
{
    lxw_error err;
    char *tmp_footer;
    char *found_string;
    char *offset_string;
    uint8_t placeholder_count = 0;
    uint8_t image_count = 0;

    if (!string) {
        LXW_WARN_FORMAT("worksheet_set_header_opt/footer_opt(): "
                        "header/footer string cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (lxw_utf8_strlen(string) > LXW_HEADER_FOOTER_MAX) {
        LXW_WARN_FORMAT("worksheet_set_header_opt/footer_opt(): "
                        "header/footer string exceeds Excel's limit of "
                        "255 characters.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    tmp_footer = lxw_strdup(string);
    RETURN_ON_MEM_ERROR(tmp_footer, LXW_ERROR_MEMORY_MALLOC_FAILED);

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
        LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                         "the number of &G/&[Picture] placeholders in option "
                         "string \"%s\" does not match the number of supplied "
                         "images.", string);

        free(tmp_footer);
        return LXW_ERROR_PARAMETER_VALIDATION;
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
            LXW_WARN_FORMAT1("worksheet_set_header_opt/footer_opt(): "
                             "the number of &G/&[Picture] placeholders in option "
                             "string \"%s\" does not match the number of supplied "
                             "images.", string);

            free(tmp_footer);
            return LXW_ERROR_PARAMETER_VALIDATION;
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
    self->header_footer_changed = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Set the page header caption.
 */
lxw_error
worksheet_set_header(lxw_worksheet *self, const char *string)
{
    return worksheet_set_header_opt(self, string, NULL);
}

/*
 * Set the page footer caption.
 */
lxw_error
worksheet_set_footer(lxw_worksheet *self, const char *string)
{
    return worksheet_set_footer_opt(self, string, NULL);
}

/*
 * Set the option to show/hide gridlines on the screen and the printed page.
 */
void
worksheet_gridlines(lxw_worksheet *self, uint8_t option)
{
    if (option == LXW_HIDE_ALL_GRIDLINES) {
        self->print_gridlines = 0;
        self->screen_gridlines = 0;
    }

    if (option & LXW_SHOW_SCREEN_GRIDLINES) {
        self->screen_gridlines = 1;
    }

    if (option & LXW_SHOW_PRINT_GRIDLINES) {
        self->print_gridlines = 1;
        self->print_options_changed = 1;
    }
}

/*
 * Center the page horizontally.
 */
void
worksheet_center_horizontally(lxw_worksheet *self)
{
    self->print_options_changed = 1;
    self->hcenter = 1;
}

/*
 * Center the page horizontally.
 */
void
worksheet_center_vertically(lxw_worksheet *self)
{
    self->print_options_changed = 1;
    self->vcenter = 1;
}

/*
 * Set the option to print the row and column headers on the printed page.
 */
void
worksheet_print_row_col_headers(lxw_worksheet *self)
{
    self->print_headers = 1;
    self->print_options_changed = 1;
}

/*
 * Set the rows to repeat at the top of each printed page.
 */
lxw_error
worksheet_repeat_rows(lxw_worksheet *self, lxw_row_t first_row,
                      lxw_row_t last_row)
{
    lxw_row_t tmp_row;
    lxw_error err;

    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }

    err = _check_dimensions(self, last_row, 0, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    self->repeat_rows.in_use = LXW_TRUE;
    self->repeat_rows.first_row = first_row;
    self->repeat_rows.last_row = last_row;

    return LXW_NO_ERROR;
}

/*
 * Set the columns to repeat at the left hand side of each printed page.
 */
lxw_error
worksheet_repeat_columns(lxw_worksheet *self, lxw_col_t first_col,
                         lxw_col_t last_col)
{
    lxw_col_t tmp_col;
    lxw_error err;

    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    err = _check_dimensions(self, last_col, 0, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    self->repeat_cols.in_use = LXW_TRUE;
    self->repeat_cols.first_col = first_col;
    self->repeat_cols.last_col = last_col;

    return LXW_NO_ERROR;
}

/*
 * Set the print area in the current worksheet.
 */
lxw_error
worksheet_print_area(lxw_worksheet *self, lxw_row_t first_row,
                     lxw_col_t first_col, lxw_row_t last_row,
                     lxw_col_t last_col)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;

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

    err = _check_dimensions(self, last_row, last_col, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    /* Ignore max area since it is the same as no print area in Excel. */
    if (first_row == 0 && first_col == 0 && last_row == LXW_ROW_MAX - 1
        && last_col == LXW_COL_MAX - 1) {
        return LXW_NO_ERROR;
    }

    self->print_area.in_use = LXW_TRUE;
    self->print_area.first_row = first_row;
    self->print_area.last_row = last_row;
    self->print_area.first_col = first_col;
    self->print_area.last_col = last_col;

    return LXW_NO_ERROR;
}

/* Store the vertical and horizontal number of pages that will define the
 * maximum area printed.
 */
void
worksheet_fit_to_pages(lxw_worksheet *self, uint16_t width, uint16_t height)
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
worksheet_set_start_page(lxw_worksheet *self, uint16_t start_page)
{
    self->page_start = start_page;
}

/*
 * Set the scale factor for the printed page.
 */
void
worksheet_set_print_scale(lxw_worksheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400)
        return;

    /* Turn off "fit to page" option. */
    self->fit_page = LXW_FALSE;

    self->print_scale = scale;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Set the print in black and white option.
 */
void
worksheet_print_black_and_white(lxw_worksheet *self)
{
    self->black_white = LXW_TRUE;
    self->page_setup_changed = LXW_TRUE;
}

/*
 * Store the horizontal page breaks on a worksheet.
 */
lxw_error
worksheet_set_h_pagebreaks(lxw_worksheet *self, lxw_row_t hbreaks[])
{
    uint16_t count = 0;

    if (hbreaks == NULL)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    while (hbreaks[count])
        count++;

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXW_BREAKS_MAX)
        count = LXW_BREAKS_MAX;

    self->hbreaks = calloc(count, sizeof(lxw_row_t));
    RETURN_ON_MEM_ERROR(self->hbreaks, LXW_ERROR_MEMORY_MALLOC_FAILED);
    memcpy(self->hbreaks, hbreaks, count * sizeof(lxw_row_t));
    self->hbreaks_count = count;

    return LXW_NO_ERROR;
}

/*
 * Store the vertical page breaks on a worksheet.
 */
lxw_error
worksheet_set_v_pagebreaks(lxw_worksheet *self, lxw_col_t vbreaks[])
{
    uint16_t count = 0;

    if (vbreaks == NULL)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    while (vbreaks[count])
        count++;

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXW_BREAKS_MAX)
        count = LXW_BREAKS_MAX;

    self->vbreaks = calloc(count, sizeof(lxw_col_t));
    RETURN_ON_MEM_ERROR(self->vbreaks, LXW_ERROR_MEMORY_MALLOC_FAILED);
    memcpy(self->vbreaks, vbreaks, count * sizeof(lxw_col_t));
    self->vbreaks_count = count;

    return LXW_NO_ERROR;
}

/*
 * Set the worksheet zoom factor.
 */
void
worksheet_set_zoom(lxw_worksheet *self, uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400) {
        LXW_WARN("worksheet_set_zoom(): "
                 "Zoom factor scale outside range: 10 <= zoom <= 400.");
        return;
    }

    self->zoom = scale;
}

/*
 * Hide cell zero values.
 */
void
worksheet_hide_zero(lxw_worksheet *self)
{
    self->show_zeros = LXW_FALSE;
}

/*
 * Display the worksheet right to left for some eastern versions of Excel.
 */
void
worksheet_right_to_left(lxw_worksheet *self)
{
    self->right_to_left = LXW_TRUE;
}

/*
 * Set the color of the worksheet tab.
 */
void
worksheet_set_tab_color(lxw_worksheet *self, lxw_color_t color)
{
    self->tab_color = color;
}

/*
 * Set the worksheet protection flags to prevent modification of worksheet
 * objects.
 */
void
worksheet_protect(lxw_worksheet *self, const char *password,
                  lxw_protection *options)
{
    struct lxw_protection_obj *protect = &self->protection;

    /* Copy any user parameters to the internal structure. */
    if (options) {
        protect->no_select_locked_cells = options->no_select_locked_cells;
        protect->no_select_unlocked_cells = options->no_select_unlocked_cells;
        protect->format_cells = options->format_cells;
        protect->format_columns = options->format_columns;
        protect->format_rows = options->format_rows;
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
        uint16_t hash = lxw_hash_password(password);
        lxw_snprintf(protect->hash, 5, "%X", hash);
    }

    protect->no_sheet = LXW_FALSE;
    protect->no_content = LXW_TRUE;
    protect->is_configured = LXW_TRUE;
}

/*
 * Set the worksheet properties for outlines and grouping.
 */
void
worksheet_outline_settings(lxw_worksheet *self,
                           uint8_t visible,
                           uint8_t symbols_below,
                           uint8_t symbols_right, uint8_t auto_style)
{
    self->outline_on = visible;
    self->outline_below = symbols_below;
    self->outline_right = symbols_right;
    self->outline_style = auto_style;

    self->outline_changed = LXW_TRUE;
}

/*
 * Set the default row properties
 */
void
worksheet_set_default_row(lxw_worksheet *self, double height,
                          uint8_t hide_unused_rows)
{
    if (height < 0)
        height = self->default_row_height;

    if (height != self->default_row_height) {
        self->default_row_height = height;
        self->row_size_changed = LXW_TRUE;
    }

    if (hide_unused_rows)
        self->default_row_zeroed = LXW_TRUE;

    self->default_row_set = LXW_TRUE;
}

/*
 * Insert an image with options into the worksheet.
 */
lxw_error
worksheet_insert_image_opt(lxw_worksheet *self,
                           lxw_row_t row_num, lxw_col_t col_num,
                           const char *filename,
                           lxw_image_options *user_options)
{
    FILE *image_stream;
    const char *description;
    lxw_object_properties *object_props;

    if (!filename) {
        LXW_WARN("worksheet_insert_image()/_opt(): "
                 "filename must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxw_fopen(filename, "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Use the filename as the default description, like Excel. */
    description = lxw_basename(filename);
    if (!description) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "couldn't get basename for file: %s.", filename);
        fclose(image_stream);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (user_options) {
        object_props->x_offset = user_options->x_offset;
        object_props->y_offset = user_options->y_offset;
        object_props->x_scale = user_options->x_scale;
        object_props->y_scale = user_options->y_scale;
        object_props->url = lxw_strdup(user_options->url);
        object_props->tip = lxw_strdup(user_options->tip);
        object_props->object_position = user_options->object_position;
        object_props->decorative = user_options->decorative;

        if (user_options->description)
            description = user_options->description;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup(filename);
    object_props->description = lxw_strdup(description);
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image into the worksheet.
 */
lxw_error
worksheet_insert_image(lxw_worksheet *self,
                       lxw_row_t row_num, lxw_col_t col_num,
                       const char *filename)
{
    return worksheet_insert_image_opt(self, row_num, col_num, filename, NULL);
}

/*
 * Insert an image buffer, with options, into the worksheet.
 */
lxw_error
worksheet_insert_image_buffer_opt(lxw_worksheet *self,
                                  lxw_row_t row_num,
                                  lxw_col_t col_num,
                                  const unsigned char *image_buffer,
                                  size_t image_size,
                                  lxw_image_options *user_options)
{
    FILE *image_stream;
    lxw_object_properties *object_props;

    if (!image_size) {
        LXW_WARN("worksheet_insert_image_buffer()/_opt(): "
                 "size must be non-zero.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxw_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXW_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXW_TRUE;
    }

    if (user_options) {
        object_props->x_offset = user_options->x_offset;
        object_props->y_offset = user_options->y_offset;
        object_props->x_scale = user_options->x_scale;
        object_props->y_scale = user_options->y_scale;
        object_props->url = lxw_strdup(user_options->url);
        object_props->tip = lxw_strdup(user_options->tip);
        object_props->object_position = user_options->object_position;
        object_props->description = lxw_strdup(user_options->description);
        object_props->decorative = user_options->decorative;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->image_props, object_props, list_pointers);
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image buffer into the worksheet.
 */
lxw_error
worksheet_insert_image_buffer(lxw_worksheet *self,
                              lxw_row_t row_num,
                              lxw_col_t col_num,
                              const unsigned char *image_buffer,
                              size_t image_size)
{
    return worksheet_insert_image_buffer_opt(self, row_num, col_num,
                                             image_buffer, image_size, NULL);
}

/*
 * Embed an image with options into the worksheet.
 */
lxw_error
worksheet_embed_image_opt(lxw_worksheet *self,
                          lxw_row_t row_num, lxw_col_t col_num,
                          const char *filename,
                          lxw_image_options *user_options)
{
    FILE *image_stream;
    lxw_object_properties *object_props;
    lxw_error err;

    if (!filename) {
        LXW_WARN("worksheet_embed_image()/_opt(): "
                 "filename must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxw_fopen(filename, "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_embed_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check and store the cell dimensions. */
    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err) {
        fclose(image_stream);
        return err;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* We only copy/use a limited number of options for embedded images. */
    if (user_options) {
        if (user_options->cell_format)
            object_props->format = user_options->cell_format;

        /* The url for embedded images is written as a cell url. */
        if (user_options->url) {
            if (!user_options->cell_format)
                object_props->format = self->default_url_format;

            self->storing_embedded_image = LXW_TRUE;
            err = worksheet_write_url(self,
                                      row_num,
                                      col_num,
                                      user_options->url,
                                      object_props->format);
            if (err) {
                _free_object_properties(object_props);
                fclose(image_stream);
                return err;
            }

            self->storing_embedded_image = LXW_FALSE;
        }

        object_props->decorative = user_options->decorative;
        if (user_options->description)
            object_props->description = lxw_strdup(user_options->description);
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup(filename);
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->embedded_image_props, object_props,
                           list_pointers);
        fclose(image_stream);

        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Embed an image into the worksheet.
 */
lxw_error
worksheet_embed_image(lxw_worksheet *self,
                      lxw_row_t row_num, lxw_col_t col_num,
                      const char *filename)
{
    return worksheet_embed_image_opt(self, row_num, col_num, filename, NULL);
}

/*
 * Embed an image buffer, with options, into the worksheet.
 */
lxw_error
worksheet_embed_image_buffer_opt(lxw_worksheet *self,
                                 lxw_row_t row_num,
                                 lxw_col_t col_num,
                                 const unsigned char *image_buffer,
                                 size_t image_size,
                                 lxw_image_options *user_options)
{
    FILE *image_stream;
    lxw_object_properties *object_props;
    lxw_error err;

    if (!image_size) {
        LXW_WARN("worksheet_embed_image_buffer()/_opt(): "
                 "size must be non-zero.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. For embedded images we really only
     * need the image type. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxw_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXW_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Check and store the cell dimensions. */
    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXW_TRUE;
    }

    /* We only copy/use a limited number of options for embedded images. */
    if (user_options) {
        if (user_options->cell_format)
            object_props->format = user_options->cell_format;

        /* The url for embedded images is written as a cell url. */
        if (user_options->url) {
            if (!user_options->cell_format)
                object_props->format = self->default_url_format;

            self->storing_embedded_image = LXW_TRUE;
            err = worksheet_write_url(self,
                                      row_num,
                                      col_num,
                                      user_options->url,
                                      object_props->format);
            if (err) {
                _free_object_properties(object_props);
                fclose(image_stream);
                return err;
            }

            self->storing_embedded_image = LXW_FALSE;
        }

        object_props->decorative = user_options->decorative;
        if (user_options->description)
            object_props->description = lxw_strdup(user_options->description);
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->row = row_num;
    object_props->col = col_num;

    if (object_props->x_scale == 0.0)
        object_props->x_scale = 1;

    if (object_props->y_scale == 0.0)
        object_props->y_scale = 1;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->embedded_image_props, object_props,
                           list_pointers);
        fclose(image_stream);

        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image buffer into the worksheet.
 */
lxw_error
worksheet_embed_image_buffer(lxw_worksheet *self,
                             lxw_row_t row_num,
                             lxw_col_t col_num,
                             const unsigned char *image_buffer,
                             size_t image_size)
{
    return worksheet_embed_image_buffer_opt(self, row_num, col_num,
                                            image_buffer, image_size, NULL);
}

/*
 * Set an image as a worksheet background.
 */
lxw_error
worksheet_set_background(lxw_worksheet *self, const char *filename)
{
    FILE *image_stream;
    lxw_object_properties *object_props;

    if (!filename) {
        LXW_WARN("worksheet_set_background(): "
                 "filename must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = lxw_fopen(filename, "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_set_background(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup(filename);
    object_props->stream = image_stream;
    object_props->is_background = LXW_TRUE;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        _free_object_properties(self->background_image);
        self->background_image = object_props;
        self->has_background_image = LXW_TRUE;
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Set an image buffer as a worksheet background.
 */
lxw_error
worksheet_set_background_buffer(lxw_worksheet *self,
                                const unsigned char *image_buffer,
                                size_t image_size)
{
    FILE *image_stream;
    lxw_object_properties *object_props;

    if (!image_size) {
        LXW_WARN("worksheet_set_background(): " "size must be non-zero.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Write the image buffer to a file (preferably in memory) so we can read
     * the dimensions like an ordinary file. */
#ifdef USE_FMEMOPEN
    image_stream = fmemopen((void *) image_buffer, image_size, "rb");

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;
#else
    image_stream = lxw_tmpfile(self->tmpdir);

    if (!image_stream)
        return LXW_ERROR_CREATING_TMPFILE;

    if (fwrite(image_buffer, 1, image_size, image_stream) != image_size) {
        fclose(image_stream);
        return LXW_ERROR_CREATING_TMPFILE;
    }

    rewind(image_stream);
#endif

    /* Create a new object to hold the image properties. */
    object_props = calloc(1, sizeof(lxw_object_properties));
    if (!object_props) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    /* Store the image data in the properties structure. */
    object_props->image_buffer = calloc(1, image_size);
    if (!object_props->image_buffer) {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }
    else {
        memcpy(object_props->image_buffer, image_buffer, image_size);
        object_props->image_buffer_size = image_size;
        object_props->is_image_buffer = LXW_TRUE;
    }

    /* Copy other options or set defaults. */
    object_props->filename = lxw_strdup("image_buffer");
    object_props->stream = image_stream;
    object_props->is_background = LXW_TRUE;

    if (_get_image_properties(object_props) == LXW_NO_ERROR) {
        _free_object_properties(self->background_image);
        self->background_image = object_props;
        self->has_background_image = LXW_TRUE;
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        _free_object_properties(object_props);
        fclose(image_stream);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an chart into the worksheet.
 */
lxw_error
worksheet_insert_chart_opt(lxw_worksheet *self,
                           lxw_row_t row_num, lxw_col_t col_num,
                           lxw_chart *chart, lxw_chart_options *user_options)
{
    lxw_object_properties *object_props;
    lxw_chart_series *series;

    if (!chart) {
        LXW_WARN("worksheet_insert_chart()/_opt(): chart must be non-NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the chart isn't being used more than once. */
    if (chart->in_use) {
        LXW_WARN("worksheet_insert_chart()/_opt(): the same chart object "
                 "cannot be inserted in a worksheet more than once.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a data series. */
    if (STAILQ_EMPTY(chart->series_list)) {
        LXW_WARN
            ("worksheet_insert_chart()/_opt(): chart must have a series.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a 'values' series. */
    STAILQ_FOREACH(series, chart->series_list, list_pointers) {
        if (!series->values->formula && !series->values->sheetname) {
            LXW_WARN("worksheet_insert_chart()/_opt(): chart must have a "
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
        object_props->object_position = user_options->object_position;
        object_props->description = lxw_strdup(user_options->description);
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

    STAILQ_INSERT_TAIL(self->chart_data, object_props, list_pointers);

    chart->in_use = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Insert an image into the worksheet.
 */
lxw_error
worksheet_insert_chart(lxw_worksheet *self,
                       lxw_row_t row_num, lxw_col_t col_num, lxw_chart *chart)
{
    return worksheet_insert_chart_opt(self, row_num, col_num, chart, NULL);
}

/*
 * Add a data validation to a worksheet, for a range. Ironically this requires
 * a lot of validation of the user input.
 */
lxw_error
worksheet_data_validation_range(lxw_worksheet *self, lxw_row_t first_row,
                                lxw_col_t first_col,
                                lxw_row_t last_row,
                                lxw_col_t last_col,
                                lxw_data_validation *validation)
{
    lxw_data_val_obj *copy;
    uint8_t is_between = LXW_FALSE;
    uint8_t is_formula = LXW_FALSE;
    uint8_t has_criteria = LXW_TRUE;
    lxw_error err;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    size_t length;

    /* No action is required for validation type 'any' unless there are
     * input messages to display.*/
    if (validation->validate == LXW_VALIDATION_TYPE_ANY
        && !(validation->input_title || validation->input_message)) {

        return LXW_NO_ERROR;
    }

    /* Check for formula types. */
    switch (validation->validate) {
        case LXW_VALIDATION_TYPE_INTEGER_FORMULA:
        case LXW_VALIDATION_TYPE_DECIMAL_FORMULA:
        case LXW_VALIDATION_TYPE_LIST_FORMULA:
        case LXW_VALIDATION_TYPE_LENGTH_FORMULA:
        case LXW_VALIDATION_TYPE_DATE_FORMULA:
        case LXW_VALIDATION_TYPE_TIME_FORMULA:
        case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
            is_formula = LXW_TRUE;
            break;
    }

    /* Check for types without a criteria. */
    switch (validation->validate) {
        case LXW_VALIDATION_TYPE_LIST:
        case LXW_VALIDATION_TYPE_LIST_FORMULA:
        case LXW_VALIDATION_TYPE_ANY:
        case LXW_VALIDATION_TYPE_CUSTOM_FORMULA:
            has_criteria = LXW_FALSE;
            break;
    }

    /* Check that a validation parameter has been specified
     * except for 'list', 'any' and 'custom'. */
    if (has_criteria && validation->criteria == LXW_VALIDATION_CRITERIA_NONE) {

        LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
                        "criteria parameter must be specified.");
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check for "between" criteria so we can do additional checks. */
    if (has_criteria
        && (validation->criteria == LXW_VALIDATION_CRITERIA_BETWEEN
            || validation->criteria == LXW_VALIDATION_CRITERIA_NOT_BETWEEN)) {

        is_between = LXW_TRUE;
    }

    /* Check that formula values are non NULL. */
    if (is_formula) {
        if (is_between) {
            if (!validation->minimum_formula) {
                LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
                                "minimum_formula parameter cannot be NULL.");
                return LXW_ERROR_PARAMETER_VALIDATION;
            }
            if (!validation->maximum_formula) {
                LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
                                "maximum_formula parameter cannot be NULL.");
                return LXW_ERROR_PARAMETER_VALIDATION;
            }
        }
        else {
            if (!validation->value_formula) {
                LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
                                "formula parameter cannot be NULL.");
                return LXW_ERROR_PARAMETER_VALIDATION;
            }
        }
    }

    /* Check Excel limitations on input strings. */
    if (validation->input_title) {
        length = lxw_utf8_strlen(validation->input_title);
        if (length > LXW_VALIDATION_MAX_TITLE_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
                             "input_title length > Excel limit of %d.",
                             LXW_VALIDATION_MAX_TITLE_LENGTH);
            return LXW_ERROR_32_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->error_title) {
        length = lxw_utf8_strlen(validation->error_title);
        if (length > LXW_VALIDATION_MAX_TITLE_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
                             "error_title length > Excel limit of %d.",
                             LXW_VALIDATION_MAX_TITLE_LENGTH);
            return LXW_ERROR_32_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->input_message) {
        length = lxw_utf8_strlen(validation->input_message);
        if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
                             "input_message length > Excel limit of %d.",
                             LXW_VALIDATION_MAX_STRING_LENGTH);
            return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->error_message) {
        length = lxw_utf8_strlen(validation->error_message);
        if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
                             "error_message length > Excel limit of %d.",
                             LXW_VALIDATION_MAX_STRING_LENGTH);
            return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
        }
    }

    if (validation->validate == LXW_VALIDATION_TYPE_LIST) {
        length = _validation_list_length(validation->value_list);

        if (length == 0) {
            LXW_WARN_FORMAT("worksheet_data_validation_cell()/_range(): "
                            "list parameters cannot be zero.");
            return LXW_ERROR_PARAMETER_VALIDATION;
        }

        if (length > LXW_VALIDATION_MAX_STRING_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_data_validation_cell()/_range(): "
                             "list length with commas > Excel limit of %d.",
                             LXW_VALIDATION_MAX_STRING_LENGTH);
            return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
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
    err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
    if (err)
        return err;

    /* Create a copy of the parameters from the user data validation. */
    copy = calloc(1, sizeof(lxw_data_val_obj));
    GOTO_LABEL_ON_MEM_ERROR(copy, mem_error);

    /* Create the data validation range. */
    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(copy->sqref, first_row, first_col);
    else
        lxw_rowcol_to_range(copy->sqref, first_row, first_col, last_row,
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
        copy->input_title = lxw_strdup_formula(validation->input_title);
        GOTO_LABEL_ON_MEM_ERROR(copy->input_title, mem_error);
    }

    if (validation->input_message) {
        copy->input_message = lxw_strdup_formula(validation->input_message);
        GOTO_LABEL_ON_MEM_ERROR(copy->input_message, mem_error);
    }

    if (validation->error_title) {
        copy->error_title = lxw_strdup_formula(validation->error_title);
        GOTO_LABEL_ON_MEM_ERROR(copy->error_title, mem_error);
    }

    if (validation->error_message) {
        copy->error_message = lxw_strdup_formula(validation->error_message);
        GOTO_LABEL_ON_MEM_ERROR(copy->error_message, mem_error);
    }

    /* Copy the formula strings. */
    if (is_formula) {
        if (is_between) {
            copy->value_formula =
                lxw_strdup_formula(validation->minimum_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
            copy->maximum_formula =
                lxw_strdup_formula(validation->maximum_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->maximum_formula, mem_error);
        }
        else {
            copy->value_formula =
                lxw_strdup_formula(validation->value_formula);
            GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
        }
    }

    /* Copy the validation list as a csv string. */
    if (validation->validate == LXW_VALIDATION_TYPE_LIST) {
        copy->value_formula = _validation_list_to_csv(validation->value_list);
        GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
    }

    if (validation->validate == LXW_VALIDATION_TYPE_DATE
        || validation->validate == LXW_VALIDATION_TYPE_TIME) {
        if (is_between) {
            copy->value_number =
                lxw_datetime_to_excel_date_with_epoch
                (&validation->minimum_datetime, self->use_1904_epoch);
            copy->maximum_number =
                lxw_datetime_to_excel_date_with_epoch
                (&validation->maximum_datetime, self->use_1904_epoch);
        }
        else {
            copy->value_number =
                lxw_datetime_to_excel_date_with_epoch
                (&validation->value_datetime, self->use_1904_epoch);
        }
    }

    /* These options are on by default so we can't take plain booleans. */
    copy->ignore_blank = validation->ignore_blank ^ 1;
    copy->show_input = validation->show_input ^ 1;
    copy->show_error = validation->show_error ^ 1;

    STAILQ_INSERT_TAIL(self->data_validations, copy, list_pointers);

    self->num_validations++;

    return LXW_NO_ERROR;

mem_error:
    _free_data_validation(copy);
    return LXW_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Add a data validation to a worksheet, for a cell.
 */
lxw_error
worksheet_data_validation_cell(lxw_worksheet *self, lxw_row_t row,
                               lxw_col_t col, lxw_data_validation *validation)
{
    return worksheet_data_validation_range(self, row, col,
                                           row, col, validation);
}

/*
 * Add a conditional format to a worksheet, for a range.
 */
lxw_error
worksheet_conditional_format_range(lxw_worksheet *self, lxw_row_t first_row,
                                   lxw_col_t first_col,
                                   lxw_row_t last_row,
                                   lxw_col_t last_col,
                                   lxw_conditional_format *user_options)
{
    lxw_cond_format_obj *cond_format;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err = LXW_NO_ERROR;
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
    err = _check_dimensions(self, last_row, last_col, LXW_TRUE, LXW_TRUE);
    if (err)
        return err;

    /* Check the validation type is in correct enum range. */
    if (user_options->type <= LXW_CONDITIONAL_TYPE_NONE ||
        user_options->type >= LXW_CONDITIONAL_TYPE_LAST) {

        LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                         "invalid type value (%d).", user_options->type);

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a copy of the parameters from the user data validation. */
    cond_format = calloc(1, sizeof(lxw_cond_format_obj));
    GOTO_LABEL_ON_MEM_ERROR(cond_format, error);

    /* Create the data validation range. */
    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(cond_format->sqref, first_row, first_col);
    else
        lxw_rowcol_to_range(cond_format->sqref, first_row, first_col,
                            last_row, last_col);

    /* Store the first cell string for text and date rules. */
    lxw_rowcol_to_cell(cond_format->first_cell, first_row, first_col);

    /* Overwrite the sqref range with a user supplied set of ranges. */
    if (user_options->multi_range) {

        if (strlen(user_options->multi_range) >= LXW_MAX_ATTRIBUTE_LENGTH) {
            LXW_WARN_FORMAT1("worksheet_conditional_format_cell()/_range(): "
                             "multi_range >= limit = %d.",
                             LXW_MAX_ATTRIBUTE_LENGTH);
            err = LXW_ERROR_PARAMETER_VALIDATION;
            goto error;
        }

        LXW_ATTRIBUTE_COPY(cond_format->sqref, user_options->multi_range);
    }

    /* Get the conditional format dxf format index. */
    if (user_options->format)
        cond_format->dxf_index =
            lxw_format_get_dxf_index(user_options->format);
    else
        cond_format->dxf_index = LXW_PROPERTY_UNSET;

    /* Set some common option for all validation types. */
    cond_format->type = user_options->type;
    cond_format->criteria = user_options->criteria;
    cond_format->stop_if_true = user_options->stop_if_true;
    cond_format->type_string = lxw_strdup(type_strings[cond_format->type]);

    /* Check that the criteria matches the conditional type. */
    err = _validate_conditional_criteria(cond_format);
    if (err)
        goto error;

    /* Validate the user input for various types of rules. */
    if (user_options->type == LXW_CONDITIONAL_TYPE_CELL
        || cond_format->type == LXW_CONDITIONAL_TYPE_DUPLICATE
        || cond_format->type == LXW_CONDITIONAL_TYPE_UNIQUE) {

        err = _validate_conditional_cell(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXW_CONDITIONAL_TYPE_TEXT) {

        err = _validate_conditional_text(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXW_CONDITIONAL_TYPE_TIME_PERIOD) {

        err = _validate_conditional_time_period(user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXW_CONDITIONAL_TYPE_AVERAGE) {

        err = _validate_conditional_average(user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_TOP
             || cond_format->type == LXW_CONDITIONAL_TYPE_BOTTOM) {

        err = _validate_conditional_top(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (user_options->type == LXW_CONDITIONAL_TYPE_FORMULA) {

        err = _validate_conditional_formula(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXW_CONDITIONAL_2_COLOR_SCALE
             || cond_format->type == LXW_CONDITIONAL_3_COLOR_SCALE) {

        err = _validate_conditional_scale(cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXW_CONDITIONAL_DATA_BAR) {

        err = _validate_conditional_data_bar(self, cond_format, user_options);
        if (err)
            goto error;
    }
    else if (cond_format->type == LXW_CONDITIONAL_TYPE_ICON_SETS) {

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
        return LXW_NO_ERROR;

error:
    _free_cond_format(cond_format);
    return err;
}

/*
 * Add a conditional format to a worksheet, for a cell.
 */
lxw_error
worksheet_conditional_format_cell(lxw_worksheet *self,
                                  lxw_row_t row,
                                  lxw_col_t col,
                                  lxw_conditional_format *options)
{
    return worksheet_conditional_format_range(self, row, col,
                                              row, col, options);
}

/*
 * Insert a button object into the worksheet.
 */
lxw_error
worksheet_insert_button(lxw_worksheet *self, lxw_row_t row_num,
                        lxw_col_t col_num, lxw_button_options *options)
{
    lxw_error err;
    lxw_vml_obj *button;

    err = _check_dimensions(self, row_num, col_num, LXW_TRUE, LXW_TRUE);
    if (err)
        return err;

    button = calloc(1, sizeof(lxw_vml_obj));
    GOTO_LABEL_ON_MEM_ERROR(button, mem_error);

    button->row = row_num;
    button->col = col_num;

    /* Set user and default parameters for the button. */
    err = _get_button_params(button, 1 + self->num_buttons, options);
    if (err)
        goto mem_error;

    /* Calculate the worksheet position of the button. */
    _worksheet_position_vml_object(self, button);

    self->has_vml = LXW_TRUE;
    self->has_buttons = LXW_TRUE;
    self->num_buttons++;

    STAILQ_INSERT_TAIL(self->button_objs, button, list_pointers);

    return LXW_NO_ERROR;

mem_error:
    if (button)
        _free_vml_object(button);

    return LXW_ERROR_MEMORY_MALLOC_FAILED;
}

/*
 * Set the VBA name for the worksheet.
 */
lxw_error
worksheet_set_vba_name(lxw_worksheet *self, const char *name)
{
    if (!name) {
        LXW_WARN("worksheet_set_vba_name(): " "name must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    self->vba_codename = lxw_strdup(name);

    return LXW_NO_ERROR;
}

/*
 * Set the default author of the cell comments.
 */
void
worksheet_set_comments_author(lxw_worksheet *self, const char *author)
{
    self->comment_author = lxw_strdup(author);
}

/*
 * Make any comments in the worksheet visible, unless explicitly hidden.
 */
void
worksheet_show_comments(lxw_worksheet *self)
{
    self->comment_display_default = LXW_COMMENT_DISPLAY_VISIBLE;
}

/*
 * Ignore various Excel errors/warnings in a worksheet for user defined ranges.
 */
lxw_error
worksheet_ignore_errors(lxw_worksheet *self, uint8_t type, const char *range)
{
    if (!range) {
        LXW_WARN("worksheet_ignore_errors(): " "'range' must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (type <= 0 || type >= LXW_IGNORE_LAST_OPTION) {
        LXW_WARN("worksheet_ignore_errors(): " "unknown option in 'type'.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Set the ranges to be ignored. */
    if (type == LXW_IGNORE_NUMBER_STORED_AS_TEXT) {
        free(self->ignore_number_stored_as_text);
        self->ignore_number_stored_as_text = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_EVAL_ERROR) {
        free(self->ignore_eval_error);
        self->ignore_eval_error = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_FORMULA_DIFFERS) {
        free(self->ignore_formula_differs);
        self->ignore_formula_differs = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_FORMULA_RANGE) {
        free(self->ignore_formula_range);
        self->ignore_formula_range = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_FORMULA_UNLOCKED) {
        free(self->ignore_formula_unlocked);
        self->ignore_formula_unlocked = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_EMPTY_CELL_REFERENCE) {
        free(self->ignore_empty_cell_reference);
        self->ignore_empty_cell_reference = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_LIST_DATA_VALIDATION) {
        free(self->ignore_list_data_validation);
        self->ignore_list_data_validation = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_CALCULATED_COLUMN) {
        free(self->ignore_calculated_column);
        self->ignore_calculated_column = lxw_strdup(range);
    }
    else if (type == LXW_IGNORE_TWO_DIGIT_TEXT_YEAR) {
        free(self->ignore_two_digit_text_year);
        self->ignore_two_digit_text_year = lxw_strdup(range);
    }

    self->has_ignore_errors = LXW_TRUE;

    return LXW_NO_ERROR;
}

/*
 * Write an error cell for versions of Excel that don't support embedded images.
 */
void
worksheet_set_error_cell(lxw_worksheet *self,
                         lxw_object_properties *object_props, uint32_t ref_id)
{
    lxw_row_t row_num = object_props->row;
    lxw_col_t col_num = object_props->col;

    lxw_cell *cell =
        _new_error_cell(row_num, col_num, ref_id, object_props->format);
    _insert_cell(self, row_num, col_num, cell);

}
