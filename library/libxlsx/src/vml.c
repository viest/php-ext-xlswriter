/*****************************************************************************
 * vml - A library for creating Excel XLSX vml files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/vml.h"
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
 * Create a new vml object.
 */
lxlsx_vml *
lxlsx_vml_new(void)
{
    lxlsx_vml *vml = calloc(1, sizeof(lxlsx_vml));
    GOTO_LABEL_ON_MEM_ERROR(vml, mem_error);

    return vml;

mem_error:
    lxlsx_vml_free(vml);
    return NULL;
}

/*
 * Free a vml object.
 */
void
lxlsx_vml_free(lxlsx_vml *vml)
{
    if (!vml)
        return;

    free(vml);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/
/*
 * Write the <x:Visible> element.
 */
STATIC void
_vml_write_visible(lxlsx_vml *self)
{
    lxlsx_xml_empty_tag(self->file, "x:Visible", NULL);
}

/*
 * Write the <v:f> element.
 */
STATIC void
_vml_write_formula(lxlsx_vml *self, char *equation)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("eqn", equation);

    lxlsx_xml_empty_tag(self->file, "v:f", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:formulas> element.
 */
STATIC void
_vml_write_formulas(lxlsx_vml *self)
{
    lxlsx_xml_start_tag(self->file, "v:formulas", NULL);

    _vml_write_formula(self, "if lineDrawn pixelLineWidth 0");
    _vml_write_formula(self, "sum @0 1 0");
    _vml_write_formula(self, "sum 0 0 @1");
    _vml_write_formula(self, "prod @2 1 2");
    _vml_write_formula(self, "prod @3 21600 pixelWidth");
    _vml_write_formula(self, "prod @3 21600 pixelHeight");
    _vml_write_formula(self, "sum @0 0 1");
    _vml_write_formula(self, "prod @6 1 2");
    _vml_write_formula(self, "prod @7 21600 pixelWidth");
    _vml_write_formula(self, "sum @8 21600 0");
    _vml_write_formula(self, "prod @7 21600 pixelHeight");
    _vml_write_formula(self, "sum @10 21600 0");

    lxlsx_xml_end_tag(self->file, "v:formulas");
}

/*
 * Write the <x:TextHAlign> element.
 */
STATIC void
_vml_write_text_halign(lxlsx_vml *self)
{

    lxlsx_xml_data_element(self->file, "x:TextHAlign", "Center", NULL);
}

/*
 * Write the <x:TextVAlign> element.
 */
STATIC void
_vml_write_text_valign(lxlsx_vml *self)
{
    lxlsx_xml_data_element(self->file, "x:TextVAlign", "Center", NULL);
}

/*
 * Write the <x:FmlaMacro> element.
 */
STATIC void
_vml_write_fmla_macro(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    lxlsx_xml_data_element(self->file, "x:FmlaMacro", lxlsx_vml_obj->macro, NULL);
}

/*
 * Write the <x:PrintObject> element.
 */
STATIC void
_vml_write_print_object(lxlsx_vml *self)
{
    lxlsx_xml_data_element(self->file, "x:PrintObject", "False", NULL);
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_aspect_ratio_lock(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXLSX_PUSH_ATTRIBUTES_STR("aspectratio", "t");

    lxlsx_xml_empty_tag(self->file, "o:lock", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_rotation_lock(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXLSX_PUSH_ATTRIBUTES_STR("rotation", "t");

    lxlsx_xml_empty_tag(self->file, "o:lock", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x:Column> element.
 */
STATIC void
_vml_write_column(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    char data[LXLSX_ATTR_32];

    lxlsx_snprintf(data, LXLSX_ATTR_32, "%d", lxlsx_vml_obj->col);

    lxlsx_xml_data_element(self->file, "x:Column", data, NULL);
}

/*
 * Write the <x:Row> element.
 */
STATIC void
_vml_write_row(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    char data[LXLSX_ATTR_32];

    lxlsx_snprintf(data, LXLSX_ATTR_32, "%d", lxlsx_vml_obj->row);

    lxlsx_xml_data_element(self->file, "x:Row", data, NULL);
}

/*
 * Write the <x:AutoFill> element.
 */
STATIC void
_vml_write_auto_fill(lxlsx_vml *self)
{
    lxlsx_xml_data_element(self->file, "x:AutoFill", "False", NULL);
}

/*
 * Write the <x:Anchor> element.
 */
STATIC void
_vml_write_anchor(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    char anchor_data[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(anchor_data,
                 LXLSX_MAX_ATTRIBUTE_LENGTH,
                 "%d, %d, %d, %d, %d, %d, %d, %d",
                 lxlsx_vml_obj->from.col,
                 (uint32_t) lxlsx_vml_obj->from.col_offset,
                 lxlsx_vml_obj->from.row,
                 (uint32_t) lxlsx_vml_obj->from.row_offset,
                 lxlsx_vml_obj->to.col,
                 (uint32_t) lxlsx_vml_obj->to.col_offset,
                 lxlsx_vml_obj->to.row, (uint32_t) lxlsx_vml_obj->to.row_offset);

    lxlsx_xml_data_element(self->file, "x:Anchor", anchor_data, NULL);
}

/*
 * Write the <x:SizeWithCells> element.
 */
STATIC void
_vml_write_size_with_cells(lxlsx_vml *self)
{
    lxlsx_xml_empty_tag(self->file, "x:SizeWithCells", NULL);
}

/*
 * Write the <x:MoveWithCells> element.
 */
STATIC void
_vml_write_move_with_cells(lxlsx_vml *self)
{
    lxlsx_xml_empty_tag(self->file, "x:MoveWithCells", NULL);
}

/*
 * Write the <v:shadow> element.
 */
STATIC void
_vml_write_shadow(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("on", "t");
    LXLSX_PUSH_ATTRIBUTES_STR("color", "black");
    LXLSX_PUSH_ATTRIBUTES_STR("obscured", "t");

    lxlsx_xml_empty_tag(self->file, "v:shadow", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:stroke> element.
 */
STATIC void
_vml_write_stroke(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("joinstyle", "miter");

    lxlsx_xml_empty_tag(self->file, "v:stroke", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_shapetype_lock(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXLSX_PUSH_ATTRIBUTES_STR("shapetype", "t");

    lxlsx_xml_empty_tag(self->file, "o:lock", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <font> element.
 */
STATIC void
_vml_write_font(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("face", "Calibri");
    LXLSX_PUSH_ATTRIBUTES_STR("size", "220");
    LXLSX_PUSH_ATTRIBUTES_STR("color", "#000000");

    lxlsx_xml_data_element(self->file, "font", lxlsx_vml_obj->name, &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:imagedata> element.
 */
STATIC void
_vml_write_imagedata(lxlsx_vml *self, uint32_t rel_index, char *name)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char rel_id[LXLSX_ATTR_32];

    lxlsx_snprintf(rel_id, LXLSX_ATTR_32, "rId%d", rel_index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("o:relid", rel_id);
    LXLSX_PUSH_ATTRIBUTES_STR("o:title", name);

    lxlsx_xml_empty_tag(self->file, "v:imagedata", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element.
 */
STATIC void
_vml_write_image_path(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("o:extrusionok", "f");
    LXLSX_PUSH_ATTRIBUTES_STR("gradientshapeok", "t");
    LXLSX_PUSH_ATTRIBUTES_STR("o:connecttype", "rect");

    lxlsx_xml_empty_tag(self->file, "v:path", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element.
 */
STATIC void
_vml_write_image_shape(lxlsx_vml *self, uint32_t lxlsx_vml_shape_id, uint32_t z_index,
                       lxlsx_vml_obj *image_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char width_str[LXLSX_ATTR_32];
    char height_str[LXLSX_ATTR_32];
    char style[LXLSX_MAX_ATTRIBUTE_LENGTH];
    char o_spid[LXLSX_ATTR_32];
    char type[] = "#_x0000_t75";
    double width;
    double height;

    /* Scale the height/width by the resolution, relative to 72dpi. */
    width = image_obj->width * (72.0 / image_obj->x_dpi);
    height = image_obj->height * (72.0 / image_obj->y_dpi);

    /* Excel uses a rounding based around 72 and 96 dpi. */
    width = 72.0 / 96.0 * (uint32_t) (width * 96.0 / 72 + 0.25);
    height = 72.0 / 96.0 * (uint32_t) (height * 96.0 / 72 + 0.25);

    lxlsx_sprintf_dbl(width_str, width);
    lxlsx_sprintf_dbl(height_str, height);

    lxlsx_snprintf(o_spid, LXLSX_ATTR_32, "_x0000_s%d", lxlsx_vml_shape_id);

    lxlsx_snprintf(style,
                 LXLSX_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:0;"
                 "margin-top:0;"
                 "width:%spt;"
                 "height:%spt;" "z-index:%d", width_str, height_str, z_index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", image_obj->image_position);
    LXLSX_PUSH_ATTRIBUTES_STR("o:spid", o_spid);
    LXLSX_PUSH_ATTRIBUTES_STR("type", type);
    LXLSX_PUSH_ATTRIBUTES_STR("style", style);

    lxlsx_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:imagedata element. */
    _vml_write_imagedata(self, image_obj->rel_index, image_obj->name);

    /* Write the o:lock element. */
    _vml_write_rotation_lock(self);

    lxlsx_xml_end_tag(self->file, "v:shape");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for images.
 */
STATIC void
_vml_write_image_shapetype(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char id[] = "_x0000_t75";
    char coordsize[] = "21600,21600";
    char o_spt[] = "75";
    char o_preferrelative[] = "t";
    char path[] = "m@4@5l@4@11@9@11@9@5xe";
    char filled[] = "f";
    char stroked[] = "f";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", id);
    LXLSX_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXLSX_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXLSX_PUSH_ATTRIBUTES_STR("o:preferrelative", o_preferrelative);
    LXLSX_PUSH_ATTRIBUTES_STR("path", path);
    LXLSX_PUSH_ATTRIBUTES_STR("filled", filled);
    LXLSX_PUSH_ATTRIBUTES_STR("stroked", stroked);

    lxlsx_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:formulas element. */
    _vml_write_formulas(self);

    /* Write the v:path element. */
    _vml_write_image_path(self);

    /* Write the o:lock element. */
    _vml_write_aspect_ratio_lock(self);

    lxlsx_xml_end_tag(self->file, "v:shapetype");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x:ClientData> element.
 */
STATIC void
_vml_write_button_client_data(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ObjectType", "Button");

    lxlsx_xml_start_tag(self->file, "x:ClientData", &attributes);

    /* Write the <x:Anchor> element. */
    _vml_write_anchor(self, lxlsx_vml_obj);

    /* Write the x:PrintObject element. */
    _vml_write_print_object(self);

    /* Write the x:AutoFill element. */
    _vml_write_auto_fill(self);

    /* Write the x:FmlaMacro element. */
    _vml_write_fmla_macro(self, lxlsx_vml_obj);

    /* Write the x:TextHAlign element. */
    _vml_write_text_halign(self);

    /* Write the x:TextVAlign element. */
    _vml_write_text_valign(self);

    lxlsx_xml_end_tag(self->file, "x:ClientData");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <div> element.
 */
STATIC void
_vml_write_button_div(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("style", "text-align:center");

    lxlsx_xml_start_tag(self->file, "div", &attributes);

    /* Write the font element. */
    _vml_write_font(self, lxlsx_vml_obj);

    lxlsx_xml_end_tag(self->file, "div");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:textbox> element.
 */
STATIC void
_vml_write_button_textbox(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("style", "mso-direction-alt:auto");
    LXLSX_PUSH_ATTRIBUTES_STR("o:singleclick", "f");

    lxlsx_xml_start_tag(self->file, "v:textbox", &attributes);

    /* Write the div element. */
    _vml_write_button_div(self, lxlsx_vml_obj);

    lxlsx_xml_end_tag(self->file, "v:textbox");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:fill> element.
 */
STATIC void
_vml_write_button_fill(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("color2", "buttonFace [67]");
    LXLSX_PUSH_ATTRIBUTES_STR("o:detectmouseclick", "t");

    lxlsx_xml_empty_tag(self->file, "v:fill", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element for buttons.
 */
STATIC void
_vml_write_button_path(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("shadowok", "f");
    LXLSX_PUSH_ATTRIBUTES_STR("o:extrusionok", "f");
    LXLSX_PUSH_ATTRIBUTES_STR("strokeok", "f");
    LXLSX_PUSH_ATTRIBUTES_STR("fillok", "f");
    LXLSX_PUSH_ATTRIBUTES_STR("o:connecttype", "rect");

    lxlsx_xml_empty_tag(self->file, "v:path", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element for buttons.
 */
STATIC void
_vml_write_button_shape(lxlsx_vml *self, uint32_t lxlsx_vml_shape_id,
                        uint32_t z_index, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char type[] = "#_x0000_t201";
    char o_button[] = "t";
    char fillcolor[] = "buttonFace [67]";
    char strokecolor[] = "windowText [64]";
    char o_insetmode[] = "auto";

    char id[LXLSX_ATTR_32];
    char margin_left[LXLSX_ATTR_32];
    char margin_top[LXLSX_ATTR_32];
    char width[LXLSX_ATTR_32];
    char height[LXLSX_ATTR_32];
    char style[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_sprintf_dbl(margin_left, lxlsx_vml_obj->col_absolute * 0.75);
    lxlsx_sprintf_dbl(margin_top, lxlsx_vml_obj->row_absolute * 0.75);
    lxlsx_sprintf_dbl(width, lxlsx_vml_obj->width * 0.75);
    lxlsx_sprintf_dbl(height, lxlsx_vml_obj->height * 0.75);

    lxlsx_snprintf(id, LXLSX_ATTR_32, "_x0000_s%d", lxlsx_vml_shape_id);

    lxlsx_snprintf(style,
                 LXLSX_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:%spt;"
                 "margin-top:%spt;"
                 "width:%spt;"
                 "height:%spt;"
                 "z-index:%d;"
                 "mso-wrap-style:tight",
                 margin_left, margin_top, width, height, z_index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", id);
    LXLSX_PUSH_ATTRIBUTES_STR("type", type);

    if (lxlsx_vml_obj->text)
        LXLSX_PUSH_ATTRIBUTES_STR("alt", lxlsx_vml_obj->text);

    LXLSX_PUSH_ATTRIBUTES_STR("style", style);
    LXLSX_PUSH_ATTRIBUTES_STR("o:button", o_button);
    LXLSX_PUSH_ATTRIBUTES_STR("fillcolor", fillcolor);
    LXLSX_PUSH_ATTRIBUTES_STR("strokecolor", strokecolor);
    LXLSX_PUSH_ATTRIBUTES_STR("o:insetmode", o_insetmode);

    lxlsx_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:fill element. */
    _vml_write_button_fill(self);

    /* Write the o:lock element. */
    _vml_write_rotation_lock(self);

    /* Write the v:textbox element. */
    _vml_write_button_textbox(self, lxlsx_vml_obj);

    /* Write the x:ClientData element. */
    _vml_write_button_client_data(self, lxlsx_vml_obj);

    lxlsx_xml_end_tag(self->file, "v:shape");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for buttons.
 */
STATIC void
_vml_write_button_shapetype(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char id[] = "_x0000_t201";
    char coordsize[] = "21600,21600";
    char o_spt[] = "201";
    char path[] = "m,l,21600r21600,l21600,xe";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", id);
    LXLSX_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXLSX_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXLSX_PUSH_ATTRIBUTES_STR("path", path);

    lxlsx_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:path element. */
    _vml_write_button_path(self);

    /* Write the o:lock element. */
    _vml_write_shapetype_lock(self);

    lxlsx_xml_end_tag(self->file, "v:shapetype");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <x:ClientData> element.
 */
STATIC void
_vml_write_comment_client_data(lxlsx_vml *self, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("ObjectType", "Note");

    lxlsx_xml_start_tag(self->file, "x:ClientData", &attributes);

    /* Write the <x:MoveWithCells> element. */
    _vml_write_move_with_cells(self);

    /* Write the <x:SizeWithCells> element. */
    _vml_write_size_with_cells(self);

    /* Write the <x:Anchor> element. */
    _vml_write_anchor(self, lxlsx_vml_obj);

    /* Write the <x:AutoFill> element. */
    _vml_write_auto_fill(self);

    /* Write the <x:Row> element. */
    _vml_write_row(self, lxlsx_vml_obj);

    /* Write the <x:Column> element. */
    _vml_write_column(self, lxlsx_vml_obj);

    /* Write the x:Visible element. */
    if (lxlsx_vml_obj->visible == LXLSX_COMMENT_DISPLAY_VISIBLE)
        _vml_write_visible(self);

    lxlsx_xml_end_tag(self->file, "x:ClientData");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <div> element.
 */
STATIC void
_vml_write_comment_div(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("style", "text-align:left");

    lxlsx_xml_start_tag(self->file, "div", &attributes);

    lxlsx_xml_end_tag(self->file, "div");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:textbox> element.
 */
STATIC void
_vml_write_comment_textbox(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("style", "mso-direction-alt:auto");

    lxlsx_xml_start_tag(self->file, "v:textbox", &attributes);

    /* Write the div element. */
    _vml_write_comment_div(self);

    lxlsx_xml_end_tag(self->file, "v:textbox");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:fill> element.
 */
STATIC void
_vml_write_comment_fill(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("color2", "#ffffe1");

    lxlsx_xml_empty_tag(self->file, "v:fill", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element.
 */
STATIC void
_vml_write_comment_path(lxlsx_vml *self, uint8_t has_gradient, char *type)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (has_gradient)
        LXLSX_PUSH_ATTRIBUTES_STR("gradientshapeok", "t");

    LXLSX_PUSH_ATTRIBUTES_STR("o:connecttype", type);

    lxlsx_xml_empty_tag(self->file, "v:path", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element for comments.
 */
STATIC void
_vml_write_comment_shape(lxlsx_vml *self, uint32_t lxlsx_vml_shape_id,
                         uint32_t z_index, lxlsx_vml_obj *lxlsx_vml_obj)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char id[LXLSX_ATTR_32];
    char margin_left[LXLSX_ATTR_32];
    char margin_top[LXLSX_ATTR_32];
    char width[LXLSX_ATTR_32];
    char height[LXLSX_ATTR_32];
    char visible[LXLSX_ATTR_32];
    char fillcolor[LXLSX_ATTR_32];
    char style[LXLSX_MAX_ATTRIBUTE_LENGTH];
    char type[] = "#_x0000_t202";
    char o_insetmode[] = "auto";

    lxlsx_sprintf_dbl(margin_left, lxlsx_vml_obj->col_absolute * 0.75);
    lxlsx_sprintf_dbl(margin_top, lxlsx_vml_obj->row_absolute * 0.75);
    lxlsx_sprintf_dbl(width, lxlsx_vml_obj->width * 0.75);
    lxlsx_sprintf_dbl(height, lxlsx_vml_obj->height * 0.75);

    lxlsx_snprintf(id, LXLSX_ATTR_32, "_x0000_s%d", lxlsx_vml_shape_id);

    if (lxlsx_vml_obj->visible == LXLSX_COMMENT_DISPLAY_DEFAULT)
        lxlsx_vml_obj->visible = self->comment_display_default;

    if (lxlsx_vml_obj->visible == LXLSX_COMMENT_DISPLAY_VISIBLE)
        lxlsx_snprintf(visible, LXLSX_ATTR_32, "visible");
    else
        lxlsx_snprintf(visible, LXLSX_ATTR_32, "hidden");

    if (lxlsx_vml_obj->color)
        lxlsx_snprintf(fillcolor, LXLSX_ATTR_32, "#%06x",
                     lxlsx_vml_obj->color & LXLSX_COLOR_MASK);
    else
        lxlsx_snprintf(fillcolor, LXLSX_ATTR_32, "#%06x", 0xffffe1);

    lxlsx_snprintf(style,
                 LXLSX_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:%spt;"
                 "margin-top:%spt;"
                 "width:%spt;"
                 "height:%spt;"
                 "z-index:%d;"
                 "visibility:%s",
                 margin_left, margin_top, width, height, z_index, visible);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", id);
    LXLSX_PUSH_ATTRIBUTES_STR("type", type);
    LXLSX_PUSH_ATTRIBUTES_STR("style", style);
    LXLSX_PUSH_ATTRIBUTES_STR("fillcolor", fillcolor);
    LXLSX_PUSH_ATTRIBUTES_STR("o:insetmode", o_insetmode);

    lxlsx_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:fill element. */
    _vml_write_comment_fill(self);

    /* Write the v:shadow element. */
    _vml_write_shadow(self);

    /* Write the v:path element. */
    _vml_write_comment_path(self, LXLSX_FALSE, "none");

    /* Write the v:textbox element. */
    _vml_write_comment_textbox(self);

    /* Write the x:ClientData element. */
    _vml_write_comment_client_data(self, lxlsx_vml_obj);

    lxlsx_xml_end_tag(self->file, "v:shape");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for comments.
 */
STATIC void
_vml_write_comment_shapetype(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char id[] = "_x0000_t202";
    char coordsize[] = "21600,21600";
    char o_spt[] = "202";
    char path[] = "m,l,21600r21600,l21600,xe";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("id", id);
    LXLSX_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXLSX_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXLSX_PUSH_ATTRIBUTES_STR("path", path);

    lxlsx_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:path element. */
    _vml_write_comment_path(self, LXLSX_TRUE, "rect");

    lxlsx_xml_end_tag(self->file, "v:shapetype");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <o:idmap> element.
 */
STATIC void
_vml_write_idmap(lxlsx_vml *self)
{
    /* Since the lxlsx_vml_data_id_str may exceed the LXLSX_MAX_ATTRIBUTE_LENGTH we
     * write it directly without the xml helper functions. */
    fprintf(self->file, "<o:idmap v:ext=\"edit\" data=\"%s\"/>",
            self->lxlsx_vml_data_id_str);
}

/*
 * Write the <o:shapelayout> element.
 */
STATIC void
_vml_write_shapelayout(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("v:ext", "edit");

    lxlsx_xml_start_tag(self->file, "o:shapelayout", &attributes);

    /* Write the o:idmap element. */
    _vml_write_idmap(self);

    lxlsx_xml_end_tag(self->file, "o:shapelayout");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xml> element.
 */
STATIC void
_vml_write_xml_namespace(lxlsx_vml *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_v[] = "urn:schemas-microsoft-com:vml";
    char xmlns_o[] = "urn:schemas-microsoft-com:office:office";
    char xmlns_x[] = "urn:schemas-microsoft-com:office:excel";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:v", xmlns_v);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:o", xmlns_o);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:x", xmlns_x);

    lxlsx_xml_start_tag(self->file, "xml", &attributes);

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
lxlsx_vml_assemble_xml_file(lxlsx_vml *self)
{
    lxlsx_vml_obj *comment_obj;
    lxlsx_vml_obj *button_obj;
    lxlsx_vml_obj *image_obj;
    uint32_t z_index = 1;

    /* Write the xml namespace element. Note, the VML files have no
     * XML declaration.*/
    _vml_write_xml_namespace(self);

    /* Write the o:shapelayout element. */
    _vml_write_shapelayout(self);

    if (self->button_objs && !STAILQ_EMPTY(self->button_objs)) {
        /* Write the <v:shapetype> element. */
        _vml_write_button_shapetype(self);

        STAILQ_FOREACH(button_obj, self->button_objs, list_pointers) {
            self->lxlsx_vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_button_shape(self, self->lxlsx_vml_shape_id, z_index,
                                    button_obj);

            z_index++;
        }
    }

    if (self->comment_objs && !STAILQ_EMPTY(self->comment_objs)) {
        /* Write the <v:shapetype> element. */
        _vml_write_comment_shapetype(self);

        STAILQ_FOREACH(comment_obj, self->comment_objs, list_pointers) {
            self->lxlsx_vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_comment_shape(self, self->lxlsx_vml_shape_id, z_index,
                                     comment_obj);

            z_index++;
        }
    }

    if (self->image_objs && !STAILQ_EMPTY(self->image_objs)) {
        /* Write the <v:shapetype> element. */
        _vml_write_image_shapetype(self);

        STAILQ_FOREACH(image_obj, self->image_objs, list_pointers) {
            self->lxlsx_vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_image_shape(self, self->lxlsx_vml_shape_id, z_index,
                                   image_obj);

            z_index++;
        }
    }

    lxlsx_xml_end_tag(self->file, "xml");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
