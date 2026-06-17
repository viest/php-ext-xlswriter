/*****************************************************************************
 * table - A library for creating Excel XLSX table files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/worksheet.h"
#include "lxlsx/table.h"
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
 * Create a new table object.
 */
lxw_table *
lxw_table_new(void)
{
    lxw_table *table = calloc(1, sizeof(lxw_table));
    GOTO_LABEL_ON_MEM_ERROR(table, mem_error);

    return table;

mem_error:
    lxw_table_free(table);
    return NULL;
}

/*
 * Free a table object.
 */
void
lxw_table_free(lxw_table *table)
{
    if (!table)
        return;

    free(table);
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
_table_xml_declaration(lxw_table *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <table> element.
 */
STATIC void
_table_write_table(lxw_table *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    lxw_table_obj *table_obj = self->table_obj;

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_INT("id", table_obj->id);

    if (table_obj->name)
        LXW_PUSH_ATTRIBUTES_STR("name", table_obj->name);
    else
        LXW_PUSH_ATTRIBUTES_STR("name", "Table1");

    if (table_obj->name)
        LXW_PUSH_ATTRIBUTES_STR("displayName", table_obj->name);
    else
        LXW_PUSH_ATTRIBUTES_STR("displayName", "Table1");

    LXW_PUSH_ATTRIBUTES_STR("ref", table_obj->sqref);

    if (table_obj->no_header_row)
        LXW_PUSH_ATTRIBUTES_STR("headerRowCount", "0");

    if (table_obj->total_row)
        LXW_PUSH_ATTRIBUTES_STR("totalsRowCount", "1");
    else
        LXW_PUSH_ATTRIBUTES_STR("totalsRowShown", "0");

    lxw_xml_start_tag(self->file, "table", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <autoFilter> element.
 */
STATIC void
_table_write_auto_filter(lxw_table *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (self->table_obj->no_autofilter)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ref", self->table_obj->filter_sqref);

    lxw_xml_empty_tag(self->file, "autoFilter", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <tableColumn> element.
 */
STATIC void
_table_write_table_column(lxw_table *self, uint16_t id,
                          lxw_table_column *column)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    int32_t dfx_id;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("id", id);

    LXW_PUSH_ATTRIBUTES_STR("name", column->header);

    if (column->total_string) {
        LXW_PUSH_ATTRIBUTES_STR("totalsRowLabel", column->total_string);
    }
    else if (column->total_function) {
        if (column->total_function == LXW_TABLE_FUNCTION_AVERAGE)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "average");
        if (column->total_function == LXW_TABLE_FUNCTION_COUNT_NUMS)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "countNums");
        if (column->total_function == LXW_TABLE_FUNCTION_COUNT)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "count");
        if (column->total_function == LXW_TABLE_FUNCTION_MAX)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "max");
        if (column->total_function == LXW_TABLE_FUNCTION_MIN)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "min");
        if (column->total_function == LXW_TABLE_FUNCTION_STD_DEV)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "stdDev");
        if (column->total_function == LXW_TABLE_FUNCTION_SUM)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "sum");
        if (column->total_function == LXW_TABLE_FUNCTION_VAR)
            LXW_PUSH_ATTRIBUTES_STR("totalsRowFunction", "var");
    }

    if (column->format) {
        dfx_id = lxw_format_get_dxf_index(column->format);
        LXW_PUSH_ATTRIBUTES_INT("dataDxfId", dfx_id);
    }

    if (column->formula) {
        lxw_xml_start_tag(self->file, "tableColumn", &attributes);
        lxw_xml_data_element(self->file, "calculatedColumnFormula",
                             column->formula, NULL);
        lxw_xml_end_tag(self->file, "tableColumn");
    }
    else {
        lxw_xml_empty_tag(self->file, "tableColumn", &attributes);
    }

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <tableColumns> element.
 */
STATIC void
_table_write_table_columns(lxw_table *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    uint16_t i;
    uint16_t num_cols = self->table_obj->num_cols;
    lxw_table_column **columns = self->table_obj->columns;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("count", num_cols);

    lxw_xml_start_tag(self->file, "tableColumns", &attributes);

    for (i = 0; i < num_cols; i++)
        _table_write_table_column(self, i + 1, columns[i]);

    lxw_xml_end_tag(self->file, "tableColumns");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <tableStyleInfo> element.
 */
STATIC void
_table_write_table_style_info(lxw_table *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char name[LXW_ATTR_32];
    lxw_table_obj *table_obj = self->table_obj;

    LXW_INIT_ATTRIBUTES();

    if (table_obj->style_type == LXW_TABLE_STYLE_TYPE_LIGHT) {
        if (table_obj->style_type_number != 0) {
            lxw_snprintf(name, LXW_ATTR_32, "TableStyleLight%d",
                         table_obj->style_type_number);
            LXW_PUSH_ATTRIBUTES_STR("name", name);
        }
    }
    else if (table_obj->style_type == LXW_TABLE_STYLE_TYPE_MEDIUM) {
        lxw_snprintf(name, LXW_ATTR_32, "TableStyleMedium%d",
                     table_obj->style_type_number);
        LXW_PUSH_ATTRIBUTES_STR("name", name);
    }
    else if (table_obj->style_type == LXW_TABLE_STYLE_TYPE_DARK) {
        lxw_snprintf(name, LXW_ATTR_32, "TableStyleDark%d",
                     table_obj->style_type_number);
        LXW_PUSH_ATTRIBUTES_STR("name", name);
    }
    else {
        LXW_PUSH_ATTRIBUTES_STR("name", "TableStyleMedium9");
    }

    if (table_obj->first_column)
        LXW_PUSH_ATTRIBUTES_STR("showFirstColumn", "1");
    else
        LXW_PUSH_ATTRIBUTES_STR("showFirstColumn", "0");

    if (table_obj->last_column)
        LXW_PUSH_ATTRIBUTES_STR("showLastColumn", "1");
    else
        LXW_PUSH_ATTRIBUTES_STR("showLastColumn", "0");

    if (table_obj->no_banded_rows)
        LXW_PUSH_ATTRIBUTES_STR("showRowStripes", "0");
    else
        LXW_PUSH_ATTRIBUTES_STR("showRowStripes", "1");

    if (table_obj->banded_columns)
        LXW_PUSH_ATTRIBUTES_STR("showColumnStripes", "1");
    else
        LXW_PUSH_ATTRIBUTES_STR("showColumnStripes", "0");

    lxw_xml_empty_tag(self->file, "tableStyleInfo", &attributes);

    LXW_FREE_ATTRIBUTES();
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
lxw_table_assemble_xml_file(lxw_table *self)
{
    /* Write the XML declaration. */
    _table_xml_declaration(self);

    /* Write the table element. */
    _table_write_table(self);

    /* Write the autoFilter element. */
    _table_write_auto_filter(self);

    /* Write the tableColumns element. */
    _table_write_table_columns(self);

    /* Write the tableStyleInfo element. */
    _table_write_table_style_info(self);

    lxw_xml_end_tag(self->file, "table");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
