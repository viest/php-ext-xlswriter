/*****************************************************************************
 * worksheet - A library for creating Excel XLSX worksheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <ctype.h>

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/worksheet.h"
#include "xlsxwriter/format.h"
#include "xlsxwriter/utility.h"
#include "xlsxwriter/relationships.h"

#define LXW_STR_MAX                      32767
#define LXW_BUFFER_SIZE                  4096
#define LXW_PORTRAIT                     1
#define LXW_LANDSCAPE                    0
#define LXW_PRINT_ACROSS                 1
#define LXW_VALIDATION_MAX_TITLE_LENGTH  32
#define LXW_VALIDATION_MAX_STRING_LENGTH 255

/*
 * Forward declarations.
 */
STATIC void _worksheet_write_rows(lxw_worksheet *self);
STATIC int _row_cmp(lxw_row *row1, lxw_row *row2);
STATIC int _cell_cmp(lxw_cell *cell1, lxw_cell *cell2);

#ifndef __clang_analyzer__
LXW_RB_GENERATE_ROW(lxw_table_rows, lxw_row, tree_pointers, _row_cmp);
LXW_RB_GENERATE_CELL(lxw_table_cells, lxw_cell, tree_pointers, _cell_cmp);
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
    lxw_row row;

    row.row_num = row_num;

    return RB_FIND(lxw_table_rows, self->table, &row);
}

/*
 * Find but don't create a cell object for a given row object and col number.
 */
lxw_cell *
lxw_worksheet_find_cell(lxw_row *row, lxw_col_t col_num)
{
    lxw_cell cell;

    if (!row)
        return NULL;

    cell.col_num = col_num;

    return RB_FIND(lxw_table_cells, row->cells, &cell);
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

    /* Initialize the cached rows. */
    worksheet->table->cached_row_num = LXW_ROW_MAX + 1;
    worksheet->hyperlinks->cached_row_num = LXW_ROW_MAX + 1;

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

    worksheet->image_data = calloc(1, sizeof(struct lxw_image_data));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->image_data, mem_error);
    STAILQ_INIT(worksheet->image_data);

    worksheet->chart_data = calloc(1, sizeof(struct lxw_chart_data));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->chart_data, mem_error);
    STAILQ_INIT(worksheet->chart_data);

    worksheet->selections = calloc(1, sizeof(struct lxw_selections));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->selections, mem_error);
    STAILQ_INIT(worksheet->selections);

    worksheet->data_validations =
        calloc(1, sizeof(struct lxw_data_validations));
    GOTO_LABEL_ON_MEM_ERROR(worksheet->data_validations, mem_error);
    STAILQ_INIT(worksheet->data_validations);

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

    if (init_data && init_data->optimize) {
        FILE *tmpfile;

        tmpfile = lxw_tmpfile(init_data->tmpdir);
        if (!tmpfile) {
            LXW_ERROR("Error creating tmpfile() for worksheet in "
                      "'constant_memory' mode.");
            goto mem_error;
        }

        worksheet->optimize_tmpfile = tmpfile;
        GOTO_LABEL_ON_MEM_ERROR(worksheet->optimize_tmpfile, mem_error);
        worksheet->file = worksheet->optimize_tmpfile;
    }

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
    }

    return worksheet;

mem_error:
    lxw_worksheet_free(worksheet);
    return NULL;
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
        && cell->type != BLANK_CELL && cell->type != BOOLEAN_CELL) {

        free(cell->u.string);
    }

    free(cell->user_data1);
    free(cell->user_data2);

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
_free_image_options(lxw_image_options *image)
{
    if (!image)
        return;

    free(image->filename);
    free(image->short_name);
    free(image->extension);
    free(image->url);
    free(image->tip);
    free(image);
}

/*
 * Free a worksheet data_validation.
 */
STATIC void
_free_data_validation(lxw_data_validation *data_validation)
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
 * Free a worksheet object.
 */
void
lxw_worksheet_free(lxw_worksheet *worksheet)
{
    lxw_row *row;
    lxw_row *next_row;
    lxw_col_t col;
    lxw_merged_range *merged_range;
    lxw_image_options *image_options;
    lxw_selection *selection;
    lxw_data_validation *data_validation;
    lxw_rel_tuple *relationship;

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

    if (worksheet->merged_ranges) {
        while (!STAILQ_EMPTY(worksheet->merged_ranges)) {
            merged_range = STAILQ_FIRST(worksheet->merged_ranges);
            STAILQ_REMOVE_HEAD(worksheet->merged_ranges, list_pointers);
            free(merged_range);
        }

        free(worksheet->merged_ranges);
    }

    if (worksheet->image_data) {
        while (!STAILQ_EMPTY(worksheet->image_data)) {
            image_options = STAILQ_FIRST(worksheet->image_data);
            STAILQ_REMOVE_HEAD(worksheet->image_data, list_pointers);
            _free_image_options(image_options);
        }

        free(worksheet->image_data);
    }

    if (worksheet->chart_data) {
        while (!STAILQ_EMPTY(worksheet->chart_data)) {
            image_options = STAILQ_FIRST(worksheet->chart_data);
            STAILQ_REMOVE_HEAD(worksheet->chart_data, list_pointers);
            _free_image_options(image_options);
        }

        free(worksheet->chart_data);
    }

    if (worksheet->selections) {
        while (!STAILQ_EMPTY(worksheet->selections)) {
            selection = STAILQ_FIRST(worksheet->selections);
            STAILQ_REMOVE_HEAD(worksheet->selections, list_pointers);
            free(selection);
        }

        free(worksheet->selections);
    }

    if (worksheet->data_validations) {
        while (!STAILQ_EMPTY(worksheet->data_validations)) {
            data_validation = STAILQ_FIRST(worksheet->data_validations);
            STAILQ_REMOVE_HEAD(worksheet->data_validations, list_pointers);
            _free_data_validation(data_validation);
        }

        free(worksheet->data_validations);
    }

    /* TODO. Add function for freeing the relationship lists. */
    while (!STAILQ_EMPTY(worksheet->external_hyperlinks)) {
        relationship = STAILQ_FIRST(worksheet->external_hyperlinks);
        STAILQ_REMOVE_HEAD(worksheet->external_hyperlinks, list_pointers);
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
    free(worksheet->external_hyperlinks);

    while (!STAILQ_EMPTY(worksheet->external_drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->external_drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->external_drawing_links, list_pointers);
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
    free(worksheet->external_drawing_links);

    while (!STAILQ_EMPTY(worksheet->drawing_links)) {
        relationship = STAILQ_FIRST(worksheet->drawing_links);
        STAILQ_REMOVE_HEAD(worksheet->drawing_links, list_pointers);
        free(relationship->type);
        free(relationship->target);
        free(relationship->target_mode);
        free(relationship);
    }
    free(worksheet->drawing_links);

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
    free(worksheet->name);
    free(worksheet->quoted_name);

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
                        char *range, lxw_format *format)
{
    lxw_cell *cell = calloc(1, sizeof(lxw_cell));
    RETURN_ON_MEM_ERROR(cell, cell);

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = ARRAY_FORMULA_CELL;
    cell->format = format;
    cell->u.string = formula;
    cell->user_data1 = range;

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
 * Insert a hyperlink object into the hyperlink list.
 */
STATIC void
_insert_hyperlink(lxw_worksheet *self, lxw_row_t row_num, lxw_col_t col_num,
                  lxw_cell *link)
{
    lxw_row *row = _get_row_list(self->hyperlinks, row_num);

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
 * Hash a worksheet password. Based on the algorithm provided by Daniel Rentz
 * of OpenOffice.
 */
STATIC uint16_t
_hash_password(const char *password)
{
    size_t count;
    uint8_t i;
    uint16_t hash = 0x0000;

    count = strlen(password);

    for (i = 0; i < count; i++) {
        uint32_t low_15;
        uint32_t high_15;
        uint32_t letter = password[i] << (i + 1);

        low_15 = letter & 0x7fff;
        high_15 = letter & (0x7fff << 15);
        high_15 = high_15 >> 15;
        letter = low_15 | high_15;

        hash ^= letter;
    }

    hash ^= count;
    hash ^= 0xCE4B;

    return hash;
}

/*
 * Simple replacement for libgen.h basename() for compatibility with MSVC. It
 * handles forward and back slashes. It doesn't copy exactly the return
 * format of basename().
 */
char *
lxw_basename(const char *path)
{

    char *forward_slash;
    char *back_slash;

    if (!path)
        return NULL;

    forward_slash = strrchr(path, '/');
    back_slash = strrchr(path, '\\');

    if (!forward_slash && !back_slash)
        return (char *) path;

    if (forward_slash > back_slash)
        return forward_slash + 1;
    else
        return back_slash + 1;
}

/* Function to count the total concatenated length of the strings in a
 * validation list array, including commas. */
size_t
_validation_list_length(char **list)
{
    uint8_t i = 0;
    size_t length = 0;

    if (!list || !list[0])
        return 0;

    while (list[i] && length <= LXW_VALIDATION_MAX_STRING_LENGTH) {
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
_validation_list_to_csv(char **list)
{
    uint8_t i = 0;
    char *str;

    /* Create a buffer for the concatenated, and quoted, string. */
    /* Add +3 for quotes and EOL. */
    str = calloc(1, LXW_VALIDATION_MAX_STRING_LENGTH + 3);
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

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

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
    if (!self->show_zeros) {
        LXW_PUSH_ATTRIBUTES_STR("showZeros", "0");
    }

    /* Display worksheet right to left for Hebrew, Arabic and others. */
    if (self->right_to_left) {
        LXW_PUSH_ATTRIBUTES_STR("rightToLeft", "1");
    }

    /* Show that the sheet tab is selected. */
    if (self->selected)
        LXW_PUSH_ATTRIBUTES_STR("tabSelected", "1");

    /* Turn outlines off. Also required in the outlinePr element. */
    if (!self->outline_on) {
        LXW_PUSH_ATTRIBUTES_STR("showOutlineSymbols", "0");
    }

    /* Set the page view/layout mode if required. */
    if (self->page_view)
        LXW_PUSH_ATTRIBUTES_STR("view", "pageLayout");

    /* Set the zoom level. */
    if (self->zoom != 100) {
        if (!self->page_view) {
            LXW_PUSH_ATTRIBUTES_INT("zoomScale", self->zoom);

            if (self->zoom_scale_normal)
                LXW_PUSH_ATTRIBUTES_INT("zoomScaleNormal", self->zoom);
        }
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

        /* Flush and rewind the temp file. */
        fflush(self->optimize_tmpfile);
        rewind(self->optimize_tmpfile);

        while (read_size) {
            read_size =
                fread(buffer, 1, LXW_BUFFER_SIZE, self->optimize_tmpfile);
            fwrite(buffer, 1, read_size, self->file);
        }

        fclose(self->optimize_tmpfile);

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
_worksheet_size_col(lxw_worksheet *self, lxw_col_t col_num)
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
        if (width == 0) {
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
_worksheet_size_row(lxw_worksheet *self, lxw_row_t row_num)
{
    lxw_row *row;
    uint32_t pixels;
    double height;

    row = lxw_worksheet_find_row(self, row_num);

    if (row) {
        height = row->height;

        if (height == 0)
            pixels = 0;
        else
            pixels = (uint32_t) (4.0 / 3.0 * height);
    }
    else {
        pixels = (uint32_t) (4.0 / 3.0 * self->default_row_height);
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
                                  lxw_image_options *image,
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

    col_start = image->col;
    row_start = image->row;
    x1 = image->x_offset;
    y1 = image->y_offset;
    width = image->width;
    height = image->height;

    /* Adjust start column for negative offsets. */
    while (x1 < 0 && col_start > 0) {
        x1 += _worksheet_size_col(self, col_start - 1);
        col_start--;
    }

    /* Adjust start row for negative offsets. */
    while (y1 < 0 && row_start > 0) {
        y1 += _worksheet_size_row(self, row_start - 1);
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
            x_abs += _worksheet_size_col(self, i);
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
            y_abs += _worksheet_size_row(self, i);
    }
    else {
        /* Optimization for when the row heights haven"t changed. */
        y_abs += self->default_row_pixels * row_start;
    }

    y_abs += y1;

    /* Adjust start col for offsets that are greater than the col width. */
    while (x1 >= _worksheet_size_col(self, col_start)) {
        x1 -= _worksheet_size_col(self, col_start);
        col_start++;
    }

    /* Adjust start row for offsets that are greater than the row height. */
    while (y1 >= _worksheet_size_row(self, row_start)) {
        y1 -= _worksheet_size_row(self, row_start);
        row_start++;
    }

    /* Initialize end cell to the same as the start cell. */
    col_end = col_start;
    row_end = row_start;

    width = width + x1;
    height = height + y1;

    /* Subtract the underlying cell widths to find the end cell. */
    while (width >= _worksheet_size_col(self, col_end)) {
        width -= _worksheet_size_col(self, col_end);
        col_end++;
    }

    /* Subtract the underlying cell heights to find the end cell. */
    while (height >= _worksheet_size_row(self, row_end)) {
        height -= _worksheet_size_row(self, row_end);
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
                                lxw_image_options *image,
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
 * Set up image/drawings.
 */
void
lxw_worksheet_prepare_image(lxw_worksheet *self,
                            uint16_t image_ref_id, uint16_t drawing_id,
                            lxw_image_options *image_data)
{
    lxw_drawing_object *drawing_object;
    lxw_rel_tuple *relationship;
    double width;
    double height;
    char filename[LXW_FILENAME_LENGTH];

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

    drawing_object->anchor_type = LXW_ANCHOR_TYPE_IMAGE;
    drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ONE_CELL;
    drawing_object->description = lxw_strdup(image_data->short_name);

    /* Scale to user scale. */
    width = image_data->width * image_data->x_scale;
    height = image_data->height * image_data->y_scale;

    /* Scale by non 96dpi resolutions. */
    width *= 96.0 / image_data->x_dpi;
    height *= 96.0 / image_data->y_dpi;

    /* Convert to the nearest pixel. */
    image_data->width = width;
    image_data->height = height;

    _worksheet_position_object_emus(self, image_data, drawing_object);

    /* Convert from pixels to emus. */
    drawing_object->width = (uint32_t) (0.5 + width * 9525);
    drawing_object->height = (uint32_t) (0.5 + height * 9525);

    lxw_add_drawing_object(self->drawing, drawing_object);

    relationship = calloc(1, sizeof(lxw_rel_tuple));
    GOTO_LABEL_ON_MEM_ERROR(relationship, mem_error);

    relationship->type = lxw_strdup("/image");
    GOTO_LABEL_ON_MEM_ERROR(relationship->type, mem_error);

    lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                 image_data->extension);

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
 * Set up chart/drawings.
 */
void
lxw_worksheet_prepare_chart(lxw_worksheet *self,
                            uint16_t chart_ref_id, uint16_t drawing_id,
                            lxw_image_options *image_data)
{
    lxw_drawing_object *drawing_object;
    lxw_rel_tuple *relationship;
    double width;
    double height;
    char filename[LXW_FILENAME_LENGTH];

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

    drawing_object->anchor_type = LXW_ANCHOR_TYPE_CHART;
    drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ONE_CELL;
    drawing_object->description = lxw_strdup("TODO_DESC");

    /* Scale to user scale. */
    width = image_data->width * image_data->x_scale;
    height = image_data->height * image_data->y_scale;

    /* Convert to the nearest pixel. */
    image_data->width = width;
    image_data->height = height;

    _worksheet_position_object_emus(self, image_data, drawing_object);

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
 * Extract width and height information from a PNG file.
 */
STATIC lxw_error
_process_png(lxw_image_options *image_options)
{
    uint32_t length;
    uint32_t offset;
    char type[4];
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_options->stream;

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
    image_options->image_type = LXW_IMAGE_PNG;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi ? x_dpi : 96;
    image_options->y_dpi = y_dpi ? x_dpi : 96;
    image_options->extension = lxw_strdup("png");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                     "no size data found in file: %s.",
                     image_options->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a JPEG file.
 */
STATIC lxw_error
_process_jpeg(lxw_image_options *image_options)
{
    uint16_t length;
    uint16_t marker;
    uint32_t offset;
    uint16_t width = 0;
    uint16_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_options->stream;

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
                goto file_error;
        }
    }

    /* Ensure that we read some valid data from the file. */
    if (width == 0)
        goto file_error;

    /* Set the image metadata. */
    image_options->image_type = LXW_IMAGE_JPEG;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi ? x_dpi : 96;
    image_options->y_dpi = y_dpi ? x_dpi : 96;
    image_options->extension = lxw_strdup("jpeg");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                     "no size data found in file: %s.",
                     image_options->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract width and height information from a BMP file.
 */
STATIC lxw_error
_process_bmp(lxw_image_options *image_options)
{
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;
    int fseek_err;

    FILE *stream = image_options->stream;

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
    image_options->image_type = LXW_IMAGE_BMP;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi;
    image_options->y_dpi = y_dpi;
    image_options->extension = lxw_strdup("bmp");

    return LXW_NO_ERROR;

file_error:
    LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                     "no size data found in file: %s.",
                     image_options->filename);

    return LXW_ERROR_IMAGE_DIMENSIONS;
}

/*
 * Extract information from the image file such as dimension, type, filename,
 * and extension.
 */
STATIC lxw_error
_get_image_properties(lxw_image_options *image_options)
{
    unsigned char signature[4];

    /* Read 4 bytes to look for the file header/signature. */
    if (fread(signature, 1, 4, image_options->stream) < 4) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "couldn't read file type for file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    if (memcmp(&signature[1], "PNG", 3) == 0) {
        if (_process_png(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (signature[0] == 0xFF && signature[1] == 0xD8) {
        if (_process_jpeg(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "BM", 2) == 0) {
        if (_process_bmp(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "unsupported image format for file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    return LXW_NO_ERROR;
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
#ifdef USE_DOUBLE_FUNCTION
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
                "<c r=\"%s\" s=\"%d\"><v>%.16g</v></c>",
                range, style_index, cell->u.number);
    else
        fprintf(self->file,
                "<c r=\"%s\"><v>%.16g</v></c>", range, cell->u.number);

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

    if (cell->u.number)
        data[0] = '1';
    else
        data[0] = '0';

    data[1] = '\0';

    lxw_xml_data_element(self->file, "v", data, NULL);
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

    /* For other cell types use the general functions. */
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("r", range);

    if (style_index)
        LXW_PUSH_ATTRIBUTES_INT("s", style_index);

    if (cell->type == FORMULA_CELL) {
        lxw_xml_start_tag(self->file, "c", &attributes);
        _write_formula_num_cell(self, cell);
        lxw_xml_end_tag(self->file, "c");
    }
    else if (cell->type == BLANK_CELL) {
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

            RB_FOREACH(cell, lxw_table_cells, row->cells) {
                _write_cell(self, cell, row->format);
            }
            lxw_xml_end_tag(self->file, "row");
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

    if (self->header[0] != '\0')
        _worksheet_write_odd_header(self);

    if (self->footer[0] != '\0')
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
        && !self->outline_changed && !self->vba_codename) {
        return;
    }

    LXW_INIT_ATTRIBUTES();

    if (self->vba_codename)
        LXW_PUSH_ATTRIBUTES_INT("codeName", self->vba_codename);

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
 * Write the <autoFilter> element.
 */
STATIC void
_worksheet_write_auto_filter(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char range[LXW_MAX_CELL_RANGE_LENGTH];

    if (!self->autofilter.in_use)
        return;

    lxw_rowcol_to_range(range,
                        self->autofilter.first_row,
                        self->autofilter.first_col,
                        self->autofilter.last_row, self->autofilter.last_col);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", range);

    lxw_xml_empty_tag(self->file, "autoFilter", &attributes);

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
_worksheet_write_sheet_protection(lxw_worksheet *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    struct lxw_protection *protect = &self->protection;

    if (!protect->is_configured)
        return;

    LXW_INIT_ATTRIBUTES();

    if (*protect->hash)
        LXW_PUSH_ATTRIBUTES_STR("password", protect->hash);

    if (!protect->no_sheet)
        LXW_PUSH_ATTRIBUTES_INT("sheet", 1);

    if (protect->content)
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
 * Write the <drawing> element.
 */
STATIC void
_write_drawing(lxw_worksheet *self, uint16_t id)
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
_write_drawings(lxw_worksheet *self)
{
    if (!self->drawing)
        return;

    self->rel_count++;

    _write_drawing(self, self->rel_count);
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
                                 lxw_data_validation *validation)
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
    lxw_data_validation *data_validation;

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
    _worksheet_write_sheet_protection(self);

    /* Write the autoFilter element. */
    _worksheet_write_auto_filter(self);

    /* Write the mergeCells element. */
    _worksheet_write_merge_cells(self);

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

    /* Write the drawing element. */
    _write_drawings(self);

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
        sst_element = lxw_get_sst_index(self->sst, string);

        if (!sst_element)
            return LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id,
                                sst_element->string, format);
    }
    else {
        /* Look for and escape control chars in the string. */
        if (strpbrk(string, "\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C"
                    "\x0D\x0E\x0F\x10\x11\x12\x13\x14\x15\x16"
                    "\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F")) {
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
 * Write a formula with a numerical result to a cell in Excel.
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

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(self, last_row, last_col, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

    /* Define the array range. */
    range = calloc(1, LXW_MAX_CELL_RANGE_LENGTH);
    RETURN_ON_MEM_ERROR(range, LXW_ERROR_MEMORY_MALLOC_FAILED);

    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(range, first_row, last_col);
    else
        lxw_rowcol_to_range(range, first_row, first_col, last_row, last_col);

    /* Copy and trip leading "{=" from formula. */
    if (formula[0] == '{')
        if (formula[1] == '=')
            formula_copy = lxw_strdup(formula + 2);
        else
            formula_copy = lxw_strdup(formula + 1);
    else
        formula_copy = lxw_strdup(formula);

    /* Strip trailing "}" from formula. */
    if (formula_copy[strlen(formula_copy) - 1] == '}')
        formula_copy[strlen(formula_copy) - 1] = '\0';

    /* Create a new array formula cell object. */
    cell = _new_array_formula_cell(first_row, first_col,
                                   formula_copy, range, format);

    cell->formula_result = result;

    _insert_cell(self, first_row, first_col, cell);

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
 * Write an array formula with a default result to a cell in Excel .
 */
lxw_error
worksheet_write_array_formula(lxw_worksheet *self,
                              lxw_row_t first_row,
                              lxw_col_t first_col,
                              lxw_row_t last_row,
                              lxw_col_t last_col,
                              const char *formula, lxw_format *format)
{
    return worksheet_write_array_formula_num(self, first_row, first_col,
                                             last_row, last_col, formula,
                                             format, 0);
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

    excel_date = lxw_datetime_to_excel_date(datetime, LXW_EPOCH_1900);

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
                        lxw_format *format, const char *string,
                        const char *tooltip)
{
    lxw_cell *link;
    char *string_copy = NULL;
    char *url_copy = NULL;
    char *url_external = NULL;
    char *url_string = NULL;
    char *tooltip_copy = NULL;
    char *found_string;
    lxw_error err;
    size_t string_size;
    size_t i;
    enum cell_types link_type = HYPERLINK_URL;

    if (!url || !*url)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* Check the Excel limit of URLS per worksheet. */
    if (self->hlink_count > LXW_MAX_NUMBER_URLS)
        return LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED;

    err = _check_dimensions(self, row_num, col_num, LXW_FALSE, LXW_FALSE);
    if (err)
        return err;

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

    /* Escape the URL. */
    if (link_type == HYPERLINK_URL && strlen(url_copy) >= 3) {
        uint8_t not_escaped = 1;

        /* First check if the URL is already escaped by the user. */
        for (i = 0; i <= strlen(url_copy) - 3; i++) {
            if (url_copy[i] == '%' && isxdigit(url_copy[i + 1])
                && isxdigit(url_copy[i + 2])) {

                not_escaped = 0;
                break;
            }
        }

        if (not_escaped) {
            url_external = calloc(1, strlen(url_copy) * 3 + 1);
            GOTO_LABEL_ON_MEM_ERROR(url_external, mem_error);

            for (i = 0; i <= strlen(url_copy); i++) {
                switch (url_copy[i]) {
                    case (' '):
                    case ('"'):
                    case ('%'):
                    case ('<'):
                    case ('>'):
                    case ('['):
                    case (']'):
                    case ('`'):
                    case ('^'):
                    case ('{'):
                    case ('}'):
                        lxw_snprintf(url_external + strlen(url_external),
                                     sizeof("%xx"), "%%%2x", url_copy[i]);
                        break;
                    default:
                        url_external[strlen(url_external)] = url_copy[i];
                }

            }

            free(url_copy);
            url_copy = lxw_strdup(url_external);
            GOTO_LABEL_ON_MEM_ERROR(url_copy, mem_error);

            free(url_external);
            url_external = NULL;
        }
    }

    if (link_type == HYPERLINK_EXTERNAL) {
        /* External Workbook links need to be modified into the right format.
         * The URL will look something like "c:\temp\file.xlsx#Sheet!A1".
         * We need the part to the left of the # as the URL and the part to
         * the right as the "location" string (if it exists).
         */

        /* For external links change the dir separator from Unix to DOS. */
        for (i = 0; i <= strlen(url_copy); i++)
            if (url_copy[i] == '/')
                url_copy[i] = '\\';

        for (i = 0; i <= strlen(string_copy); i++)
            if (string_copy[i] == '/')
                string_copy[i] = '\\';

        found_string = strchr(url_copy, '#');

        if (found_string) {
            url_string = lxw_strdup(found_string + 1);
            GOTO_LABEL_ON_MEM_ERROR(url_string, mem_error);

            *found_string = '\0';
        }

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

    /* Excel limits escaped URL to 255 characters. */
    if (lxw_utf8_strlen(url_copy) > 255)
        goto mem_error;

    err = worksheet_write_string(self, row_num, col_num, string_copy, format);
    if (err)
        goto mem_error;

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
    return LXW_ERROR_MEMORY_MALLOC_FAILED;
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
        lxw_col_t col;
        lxw_col_t old_size = self->col_options_max;
        lxw_col_t new_size = _next_power_of_two(firstcol + 1);
        lxw_col_options **new_ptr = realloc(self->col_options,
                                            new_size *
                                            sizeof(lxw_col_options *));

        if (new_ptr) {
            for (col = old_size; col < new_size; col++)
                new_ptr[col] = NULL;

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

    self->autofilter.in_use = LXW_TRUE;
    self->autofilter.first_row = first_row;
    self->autofilter.first_col = first_col;
    self->autofilter.last_row = last_row;
    self->autofilter.last_col = last_col;

    return LXW_NO_ERROR;
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
void
worksheet_set_selection(lxw_worksheet *self,
                        lxw_row_t first_row, lxw_col_t first_col,
                        lxw_row_t last_row, lxw_col_t last_col)
{
    lxw_selection *selection;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    char active_cell[LXW_MAX_CELL_RANGE_LENGTH];
    char sqref[LXW_MAX_CELL_RANGE_LENGTH];

    /* Only allow selection to be set once to avoid freeing/re-creating it. */
    if (!STAILQ_EMPTY(self->selections))
        return;

    /* Excel doesn't set a selection for cell A1 since it is the default. */
    if (first_row == 0 && first_col == 0 && last_row == 0 && last_col == 0)
        return;

    selection = calloc(1, sizeof(lxw_selection));
    RETURN_VOID_ON_MEM_ERROR(selection);

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
    if (options) {
        if (options->margin >= 0.0)
            self->margin_header = options->margin;
    }

    if (!string)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_utf8_strlen(string) >= LXW_HEADER_FOOTER_MAX)
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;

    lxw_strcpy(self->header, string);
    self->header_footer_changed = 1;

    return LXW_NO_ERROR;
}

/*
 * Set the page footer caption and options.
 */
lxw_error
worksheet_set_footer_opt(lxw_worksheet *self, const char *string,
                         lxw_header_footer_options *options)
{
    if (options) {
        if (options->margin >= 0.0)
            self->margin_footer = options->margin;
    }

    if (!string)
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (lxw_utf8_strlen(string) >= LXW_HEADER_FOOTER_MAX)
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;

    lxw_strcpy(self->footer, string);
    self->header_footer_changed = 1;

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
    struct lxw_protection *protect = &self->protection;

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
        uint16_t hash = _hash_password(password);
        lxw_snprintf(protect->hash, 5, "%X", hash);
    }

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
 * Insert an image into the worksheet.
 */
lxw_error
worksheet_insert_image_opt(lxw_worksheet *self,
                           lxw_row_t row_num, lxw_col_t col_num,
                           const char *filename,
                           lxw_image_options *user_options)
{
    FILE *image_stream;
    char *short_name;
    lxw_image_options *options;

    if (!filename) {
        LXW_WARN("worksheet_insert_image()/_opt(): "
                 "filename must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = fopen(filename, "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Get the filename from the full path to add to the Drawing object. */
    short_name = lxw_basename(filename);
    if (!short_name) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "couldn't get basename for file: %s.", filename);
        fclose(image_stream);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image options. */
    options = calloc(1, sizeof(lxw_image_options));
    if (!options) {
        fclose(image_stream);
        return LXW_ERROR_MEMORY_MALLOC_FAILED;
    }

    if (user_options) {
        options->x_offset = user_options->x_offset;
        options->y_offset = user_options->y_offset;
        options->x_scale = user_options->x_scale;
        options->y_scale = user_options->y_scale;
    }

    /* Copy other options or set defaults. */
    options->filename = lxw_strdup(filename);
    options->short_name = lxw_strdup(short_name);
    options->stream = image_stream;
    options->row = row_num;
    options->col = col_num;

    if (!options->x_scale)
        options->x_scale = 1;

    if (!options->y_scale)
        options->y_scale = 1;

    if (_get_image_properties(options) == LXW_NO_ERROR) {
        STAILQ_INSERT_TAIL(self->image_data, options, list_pointers);
        fclose(image_stream);
        return LXW_NO_ERROR;
    }
    else {
        free(options);
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
 * Insert an chart into the worksheet.
 */
lxw_error
worksheet_insert_chart_opt(lxw_worksheet *self,
                           lxw_row_t row_num, lxw_col_t col_num,
                           lxw_chart *chart, lxw_image_options *user_options)
{
    lxw_image_options *options;
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

    /* Create a new object to hold the chart image options. */
    options = calloc(1, sizeof(lxw_image_options));
    RETURN_ON_MEM_ERROR(options, LXW_ERROR_MEMORY_MALLOC_FAILED);

    if (user_options) {
        options->x_offset = user_options->x_offset;
        options->y_offset = user_options->y_offset;
        options->x_scale = user_options->x_scale;
        options->y_scale = user_options->y_scale;
    }

    /* Copy other options or set defaults. */
    options->row = row_num;
    options->col = col_num;

    /* TODO. Read defaults from chart. */
    options->width = 480;
    options->height = 288;

    if (!options->x_scale)
        options->x_scale = 1;

    if (!options->y_scale)
        options->y_scale = 1;

    /* Store chart references so they can be ordered in the workbook. */
    options->chart = chart;

    STAILQ_INSERT_TAIL(self->chart_data, options, list_pointers);

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
    lxw_data_validation *copy;
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
    copy = calloc(1, sizeof(lxw_data_validation));
    GOTO_LABEL_ON_MEM_ERROR(copy, mem_error);

    /* Create the data validation range. */
    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(copy->sqref, first_row, last_col);
    else
        lxw_rowcol_to_range(copy->sqref, first_row, first_col, last_row,
                            last_col);

    /* Copy the parameters from the user data validation. */
    copy->validate = validation->validate;
    copy->value_number = validation->value_number;
    copy->error_type = validation->error_type;
    copy->dropdown = validation->dropdown;
    copy->is_between = is_between;

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

    if (validation->validate == LXW_VALIDATION_TYPE_LIST_FORMULA) {
        copy->value_formula = lxw_strdup_formula(validation->value_formula);
        GOTO_LABEL_ON_MEM_ERROR(copy->value_formula, mem_error);
    }

    if (validation->validate == LXW_VALIDATION_TYPE_DATE
        || validation->validate == LXW_VALIDATION_TYPE_TIME) {
        if (is_between) {
            copy->value_number =
                lxw_datetime_to_excel_date(&validation->minimum_datetime,
                                           LXW_EPOCH_1900);
            copy->maximum_number =
                lxw_datetime_to_excel_date(&validation->maximum_datetime,
                                           LXW_EPOCH_1900);
        }
        else {
            copy->value_number =
                lxw_datetime_to_excel_date(&validation->value_datetime,
                                           LXW_EPOCH_1900);
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
