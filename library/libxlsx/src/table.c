/*****************************************************************************
 * table - A library for creating Excel XLSX table files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/table.h"
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
 * Create a new table object.
 */
lxlsx_table *
lxlsx_table_new(void)
{
    lxlsx_table *table = calloc(1, sizeof(lxlsx_table));
    GOTO_LABEL_ON_MEM_ERROR(table, mem_error);

    return table;

mem_error:
    lxlsx_table_free(table);
    return NULL;
}

/*
 * Free a table object.
 */
void
lxlsx_table_free(lxlsx_table *table)
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
_table_xml_declaration(lxlsx_table *self)
{
    lxlsx_xml_declaration(self->file);
}

/*
 * Write the <table> element.
 */
STATIC void
_table_write_table(lxlsx_table *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    lxlsx_table_obj *lxlsx_table_obj = self->lxlsx_table_obj;

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXLSX_PUSH_ATTRIBUTES_INT("id", lxlsx_table_obj->id);

    if (lxlsx_table_obj->name)
        LXLSX_PUSH_ATTRIBUTES_STR("name", lxlsx_table_obj->name);
    else
        LXLSX_PUSH_ATTRIBUTES_STR("name", "Table1");

    if (lxlsx_table_obj->name)
        LXLSX_PUSH_ATTRIBUTES_STR("displayName", lxlsx_table_obj->name);
    else
        LXLSX_PUSH_ATTRIBUTES_STR("displayName", "Table1");

    LXLSX_PUSH_ATTRIBUTES_STR("ref", lxlsx_table_obj->sqref);

    if (lxlsx_table_obj->no_header_row)
        LXLSX_PUSH_ATTRIBUTES_STR("headerRowCount", "0");

    if (lxlsx_table_obj->total_row)
        LXLSX_PUSH_ATTRIBUTES_STR("totalsRowCount", "1");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("totalsRowShown", "0");

    lxlsx_xml_start_tag(self->file, "table", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <autoFilter> element.
 */
STATIC void
_table_write_auto_filter(lxlsx_table *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    if (self->lxlsx_table_obj->no_autofilter)
        return;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ref", self->lxlsx_table_obj->filter_sqref);

    lxlsx_xml_empty_tag(self->file, "autoFilter", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <tableColumn> element.
 */
STATIC void
_table_write_table_column(lxlsx_table *self, uint16_t id,
                          lxlsx_table_column *column)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    int32_t dfx_id;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("id", id);

    LXLSX_PUSH_ATTRIBUTES_STR("name", column->header);

    if (column->total_string) {
        LXLSX_PUSH_ATTRIBUTES_STR("totalsRowLabel", column->total_string);
    }
    else if (column->total_function) {
        if (column->total_function == LXLSX_TABLE_FUNCTION_AVERAGE)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "average");
        if (column->total_function == LXLSX_TABLE_FUNCTION_COUNT_NUMS)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "countNums");
        if (column->total_function == LXLSX_TABLE_FUNCTION_COUNT)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "count");
        if (column->total_function == LXLSX_TABLE_FUNCTION_MAX)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "max");
        if (column->total_function == LXLSX_TABLE_FUNCTION_MIN)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "min");
        if (column->total_function == LXLSX_TABLE_FUNCTION_STD_DEV)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "stdDev");
        if (column->total_function == LXLSX_TABLE_FUNCTION_SUM)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "sum");
        if (column->total_function == LXLSX_TABLE_FUNCTION_VAR)
            LXLSX_PUSH_ATTRIBUTES_STR("totalsRowFunction", "var");
    }

    if (column->format) {
        dfx_id = lxlsx_format_get_dxf_index(column->format);
        LXLSX_PUSH_ATTRIBUTES_INT("dataDxfId", dfx_id);
    }

    if (column->formula) {
        lxlsx_xml_start_tag(self->file, "tableColumn", &attributes);
        lxlsx_xml_data_element(self->file, "calculatedColumnFormula",
                             column->formula, NULL);
        lxlsx_xml_end_tag(self->file, "tableColumn");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "tableColumn", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <tableColumns> element.
 */
STATIC void
_table_write_table_columns(lxlsx_table *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    uint16_t i;
    uint16_t num_cols = self->lxlsx_table_obj->num_cols;
    lxlsx_table_column **columns = self->lxlsx_table_obj->columns;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("count", num_cols);

    lxlsx_xml_start_tag(self->file, "tableColumns", &attributes);

    for (i = 0; i < num_cols; i++)
        _table_write_table_column(self, i + 1, columns[i]);

    lxlsx_xml_end_tag(self->file, "tableColumns");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <tableStyleInfo> element.
 */
STATIC void
_table_write_table_style_info(lxlsx_table *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char name[LXLSX_ATTR_32];
    lxlsx_table_obj *lxlsx_table_obj = self->lxlsx_table_obj;

    LXLSX_INIT_ATTRIBUTES();

    if (lxlsx_table_obj->style_type == LXLSX_TABLE_STYLE_TYPE_LIGHT) {
        if (lxlsx_table_obj->style_type_number != 0) {
            lxlsx_snprintf(name, LXLSX_ATTR_32, "TableStyleLight%d",
                         lxlsx_table_obj->style_type_number);
            LXLSX_PUSH_ATTRIBUTES_STR("name", name);
        }
    }
    else if (lxlsx_table_obj->style_type == LXLSX_TABLE_STYLE_TYPE_MEDIUM) {
        lxlsx_snprintf(name, LXLSX_ATTR_32, "TableStyleMedium%d",
                     lxlsx_table_obj->style_type_number);
        LXLSX_PUSH_ATTRIBUTES_STR("name", name);
    }
    else if (lxlsx_table_obj->style_type == LXLSX_TABLE_STYLE_TYPE_DARK) {
        lxlsx_snprintf(name, LXLSX_ATTR_32, "TableStyleDark%d",
                     lxlsx_table_obj->style_type_number);
        LXLSX_PUSH_ATTRIBUTES_STR("name", name);
    }
    else {
        LXLSX_PUSH_ATTRIBUTES_STR("name", "TableStyleMedium9");
    }

    if (lxlsx_table_obj->first_column)
        LXLSX_PUSH_ATTRIBUTES_STR("showFirstColumn", "1");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("showFirstColumn", "0");

    if (lxlsx_table_obj->last_column)
        LXLSX_PUSH_ATTRIBUTES_STR("showLastColumn", "1");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("showLastColumn", "0");

    if (lxlsx_table_obj->no_banded_rows)
        LXLSX_PUSH_ATTRIBUTES_STR("showRowStripes", "0");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("showRowStripes", "1");

    if (lxlsx_table_obj->banded_columns)
        LXLSX_PUSH_ATTRIBUTES_STR("showColumnStripes", "1");
    else
        LXLSX_PUSH_ATTRIBUTES_STR("showColumnStripes", "0");

    lxlsx_xml_empty_tag(self->file, "tableStyleInfo", &attributes);

    LXLSX_FREE_ATTRIBUTES();
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
lxlsx_table_assemble_xml_file(lxlsx_table *self)
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

    lxlsx_xml_end_tag(self->file, "table");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
