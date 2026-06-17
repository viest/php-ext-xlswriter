/*****************************************************************************
 * vml - A library for creating Excel XLSX vml files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "lxlsx/xmlwriter.h"
#include "lxlsx/vml.h"
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
 * Create a new vml object.
 */
lxw_vml *
lxw_vml_new(void)
{
    lxw_vml *vml = calloc(1, sizeof(lxw_vml));
    GOTO_LABEL_ON_MEM_ERROR(vml, mem_error);

    return vml;

mem_error:
    lxw_vml_free(vml);
    return NULL;
}

/*
 * Free a vml object.
 */
void
lxw_vml_free(lxw_vml *vml)
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
_vml_write_visible(lxw_vml *self)
{
    lxw_xml_empty_tag(self->file, "x:Visible", NULL);
}

/*
 * Write the <v:f> element.
 */
STATIC void
_vml_write_formula(lxw_vml *self, char *equation)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("eqn", equation);

    lxw_xml_empty_tag(self->file, "v:f", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:formulas> element.
 */
STATIC void
_vml_write_formulas(lxw_vml *self)
{
    lxw_xml_start_tag(self->file, "v:formulas", NULL);

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

    lxw_xml_end_tag(self->file, "v:formulas");
}

/*
 * Write the <x:TextHAlign> element.
 */
STATIC void
_vml_write_text_halign(lxw_vml *self)
{

    lxw_xml_data_element(self->file, "x:TextHAlign", "Center", NULL);
}

/*
 * Write the <x:TextVAlign> element.
 */
STATIC void
_vml_write_text_valign(lxw_vml *self)
{
    lxw_xml_data_element(self->file, "x:TextVAlign", "Center", NULL);
}

/*
 * Write the <x:FmlaMacro> element.
 */
STATIC void
_vml_write_fmla_macro(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    lxw_xml_data_element(self->file, "x:FmlaMacro", vml_obj->macro, NULL);
}

/*
 * Write the <x:PrintObject> element.
 */
STATIC void
_vml_write_print_object(lxw_vml *self)
{
    lxw_xml_data_element(self->file, "x:PrintObject", "False", NULL);
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_aspect_ratio_lock(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXW_PUSH_ATTRIBUTES_STR("aspectratio", "t");

    lxw_xml_empty_tag(self->file, "o:lock", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_rotation_lock(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXW_PUSH_ATTRIBUTES_STR("rotation", "t");

    lxw_xml_empty_tag(self->file, "o:lock", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x:Column> element.
 */
STATIC void
_vml_write_column(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%d", vml_obj->col);

    lxw_xml_data_element(self->file, "x:Column", data, NULL);
}

/*
 * Write the <x:Row> element.
 */
STATIC void
_vml_write_row(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%d", vml_obj->row);

    lxw_xml_data_element(self->file, "x:Row", data, NULL);
}

/*
 * Write the <x:AutoFill> element.
 */
STATIC void
_vml_write_auto_fill(lxw_vml *self)
{
    lxw_xml_data_element(self->file, "x:AutoFill", "False", NULL);
}

/*
 * Write the <x:Anchor> element.
 */
STATIC void
_vml_write_anchor(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    char anchor_data[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(anchor_data,
                 LXW_MAX_ATTRIBUTE_LENGTH,
                 "%d, %d, %d, %d, %d, %d, %d, %d",
                 vml_obj->from.col,
                 (uint32_t) vml_obj->from.col_offset,
                 vml_obj->from.row,
                 (uint32_t) vml_obj->from.row_offset,
                 vml_obj->to.col,
                 (uint32_t) vml_obj->to.col_offset,
                 vml_obj->to.row, (uint32_t) vml_obj->to.row_offset);

    lxw_xml_data_element(self->file, "x:Anchor", anchor_data, NULL);
}

/*
 * Write the <x:SizeWithCells> element.
 */
STATIC void
_vml_write_size_with_cells(lxw_vml *self)
{
    lxw_xml_empty_tag(self->file, "x:SizeWithCells", NULL);
}

/*
 * Write the <x:MoveWithCells> element.
 */
STATIC void
_vml_write_move_with_cells(lxw_vml *self)
{
    lxw_xml_empty_tag(self->file, "x:MoveWithCells", NULL);
}

/*
 * Write the <v:shadow> element.
 */
STATIC void
_vml_write_shadow(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("on", "t");
    LXW_PUSH_ATTRIBUTES_STR("color", "black");
    LXW_PUSH_ATTRIBUTES_STR("obscured", "t");

    lxw_xml_empty_tag(self->file, "v:shadow", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:stroke> element.
 */
STATIC void
_vml_write_stroke(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("joinstyle", "miter");

    lxw_xml_empty_tag(self->file, "v:stroke", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <o:lock> element.
 */
STATIC void
_vml_write_shapetype_lock(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("v:ext", "edit");
    LXW_PUSH_ATTRIBUTES_STR("shapetype", "t");

    lxw_xml_empty_tag(self->file, "o:lock", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <font> element.
 */
STATIC void
_vml_write_font(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("face", "Calibri");
    LXW_PUSH_ATTRIBUTES_STR("size", "220");
    LXW_PUSH_ATTRIBUTES_STR("color", "#000000");

    lxw_xml_data_element(self->file, "font", vml_obj->name, &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:imagedata> element.
 */
STATIC void
_vml_write_imagedata(lxw_vml *self, uint32_t rel_index, char *name)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rel_id[LXW_ATTR_32];

    lxw_snprintf(rel_id, LXW_ATTR_32, "rId%d", rel_index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("o:relid", rel_id);
    LXW_PUSH_ATTRIBUTES_STR("o:title", name);

    lxw_xml_empty_tag(self->file, "v:imagedata", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element.
 */
STATIC void
_vml_write_image_path(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("o:extrusionok", "f");
    LXW_PUSH_ATTRIBUTES_STR("gradientshapeok", "t");
    LXW_PUSH_ATTRIBUTES_STR("o:connecttype", "rect");

    lxw_xml_empty_tag(self->file, "v:path", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element.
 */
STATIC void
_vml_write_image_shape(lxw_vml *self, uint32_t vml_shape_id, uint32_t z_index,
                       lxw_vml_obj *image_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char width_str[LXW_ATTR_32];
    char height_str[LXW_ATTR_32];
    char style[LXW_MAX_ATTRIBUTE_LENGTH];
    char o_spid[LXW_ATTR_32];
    char type[] = "#_x0000_t75";
    double width;
    double height;

    /* Scale the height/width by the resolution, relative to 72dpi. */
    width = image_obj->width * (72.0 / image_obj->x_dpi);
    height = image_obj->height * (72.0 / image_obj->y_dpi);

    /* Excel uses a rounding based around 72 and 96 dpi. */
    width = 72.0 / 96.0 * (uint32_t) (width * 96.0 / 72 + 0.25);
    height = 72.0 / 96.0 * (uint32_t) (height * 96.0 / 72 + 0.25);

    lxw_sprintf_dbl(width_str, width);
    lxw_sprintf_dbl(height_str, height);

    lxw_snprintf(o_spid, LXW_ATTR_32, "_x0000_s%d", vml_shape_id);

    lxw_snprintf(style,
                 LXW_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:0;"
                 "margin-top:0;"
                 "width:%spt;"
                 "height:%spt;" "z-index:%d", width_str, height_str, z_index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", image_obj->image_position);
    LXW_PUSH_ATTRIBUTES_STR("o:spid", o_spid);
    LXW_PUSH_ATTRIBUTES_STR("type", type);
    LXW_PUSH_ATTRIBUTES_STR("style", style);

    lxw_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:imagedata element. */
    _vml_write_imagedata(self, image_obj->rel_index, image_obj->name);

    /* Write the o:lock element. */
    _vml_write_rotation_lock(self);

    lxw_xml_end_tag(self->file, "v:shape");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for images.
 */
STATIC void
_vml_write_image_shapetype(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char id[] = "_x0000_t75";
    char coordsize[] = "21600,21600";
    char o_spt[] = "75";
    char o_preferrelative[] = "t";
    char path[] = "m@4@5l@4@11@9@11@9@5xe";
    char filled[] = "f";
    char stroked[] = "f";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", id);
    LXW_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXW_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXW_PUSH_ATTRIBUTES_STR("o:preferrelative", o_preferrelative);
    LXW_PUSH_ATTRIBUTES_STR("path", path);
    LXW_PUSH_ATTRIBUTES_STR("filled", filled);
    LXW_PUSH_ATTRIBUTES_STR("stroked", stroked);

    lxw_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:formulas element. */
    _vml_write_formulas(self);

    /* Write the v:path element. */
    _vml_write_image_path(self);

    /* Write the o:lock element. */
    _vml_write_aspect_ratio_lock(self);

    lxw_xml_end_tag(self->file, "v:shapetype");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x:ClientData> element.
 */
STATIC void
_vml_write_button_client_data(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ObjectType", "Button");

    lxw_xml_start_tag(self->file, "x:ClientData", &attributes);

    /* Write the <x:Anchor> element. */
    _vml_write_anchor(self, vml_obj);

    /* Write the x:PrintObject element. */
    _vml_write_print_object(self);

    /* Write the x:AutoFill element. */
    _vml_write_auto_fill(self);

    /* Write the x:FmlaMacro element. */
    _vml_write_fmla_macro(self, vml_obj);

    /* Write the x:TextHAlign element. */
    _vml_write_text_halign(self);

    /* Write the x:TextVAlign element. */
    _vml_write_text_valign(self);

    lxw_xml_end_tag(self->file, "x:ClientData");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <div> element.
 */
STATIC void
_vml_write_button_div(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("style", "text-align:center");

    lxw_xml_start_tag(self->file, "div", &attributes);

    /* Write the font element. */
    _vml_write_font(self, vml_obj);

    lxw_xml_end_tag(self->file, "div");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:textbox> element.
 */
STATIC void
_vml_write_button_textbox(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("style", "mso-direction-alt:auto");
    LXW_PUSH_ATTRIBUTES_STR("o:singleclick", "f");

    lxw_xml_start_tag(self->file, "v:textbox", &attributes);

    /* Write the div element. */
    _vml_write_button_div(self, vml_obj);

    lxw_xml_end_tag(self->file, "v:textbox");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:fill> element.
 */
STATIC void
_vml_write_button_fill(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("color2", "buttonFace [67]");
    LXW_PUSH_ATTRIBUTES_STR("o:detectmouseclick", "t");

    lxw_xml_empty_tag(self->file, "v:fill", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element for buttons.
 */
STATIC void
_vml_write_button_path(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("shadowok", "f");
    LXW_PUSH_ATTRIBUTES_STR("o:extrusionok", "f");
    LXW_PUSH_ATTRIBUTES_STR("strokeok", "f");
    LXW_PUSH_ATTRIBUTES_STR("fillok", "f");
    LXW_PUSH_ATTRIBUTES_STR("o:connecttype", "rect");

    lxw_xml_empty_tag(self->file, "v:path", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element for buttons.
 */
STATIC void
_vml_write_button_shape(lxw_vml *self, uint32_t vml_shape_id,
                        uint32_t z_index, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char type[] = "#_x0000_t201";
    char o_button[] = "t";
    char fillcolor[] = "buttonFace [67]";
    char strokecolor[] = "windowText [64]";
    char o_insetmode[] = "auto";

    char id[LXW_ATTR_32];
    char margin_left[LXW_ATTR_32];
    char margin_top[LXW_ATTR_32];
    char width[LXW_ATTR_32];
    char height[LXW_ATTR_32];
    char style[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_sprintf_dbl(margin_left, vml_obj->col_absolute * 0.75);
    lxw_sprintf_dbl(margin_top, vml_obj->row_absolute * 0.75);
    lxw_sprintf_dbl(width, vml_obj->width * 0.75);
    lxw_sprintf_dbl(height, vml_obj->height * 0.75);

    lxw_snprintf(id, LXW_ATTR_32, "_x0000_s%d", vml_shape_id);

    lxw_snprintf(style,
                 LXW_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:%spt;"
                 "margin-top:%spt;"
                 "width:%spt;"
                 "height:%spt;"
                 "z-index:%d;"
                 "mso-wrap-style:tight",
                 margin_left, margin_top, width, height, z_index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", id);
    LXW_PUSH_ATTRIBUTES_STR("type", type);

    if (vml_obj->text)
        LXW_PUSH_ATTRIBUTES_STR("alt", vml_obj->text);

    LXW_PUSH_ATTRIBUTES_STR("style", style);
    LXW_PUSH_ATTRIBUTES_STR("o:button", o_button);
    LXW_PUSH_ATTRIBUTES_STR("fillcolor", fillcolor);
    LXW_PUSH_ATTRIBUTES_STR("strokecolor", strokecolor);
    LXW_PUSH_ATTRIBUTES_STR("o:insetmode", o_insetmode);

    lxw_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:fill element. */
    _vml_write_button_fill(self);

    /* Write the o:lock element. */
    _vml_write_rotation_lock(self);

    /* Write the v:textbox element. */
    _vml_write_button_textbox(self, vml_obj);

    /* Write the x:ClientData element. */
    _vml_write_button_client_data(self, vml_obj);

    lxw_xml_end_tag(self->file, "v:shape");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for buttons.
 */
STATIC void
_vml_write_button_shapetype(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char id[] = "_x0000_t201";
    char coordsize[] = "21600,21600";
    char o_spt[] = "201";
    char path[] = "m,l,21600r21600,l21600,xe";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", id);
    LXW_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXW_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXW_PUSH_ATTRIBUTES_STR("path", path);

    lxw_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:path element. */
    _vml_write_button_path(self);

    /* Write the o:lock element. */
    _vml_write_shapetype_lock(self);

    lxw_xml_end_tag(self->file, "v:shapetype");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <x:ClientData> element.
 */
STATIC void
_vml_write_comment_client_data(lxw_vml *self, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("ObjectType", "Note");

    lxw_xml_start_tag(self->file, "x:ClientData", &attributes);

    /* Write the <x:MoveWithCells> element. */
    _vml_write_move_with_cells(self);

    /* Write the <x:SizeWithCells> element. */
    _vml_write_size_with_cells(self);

    /* Write the <x:Anchor> element. */
    _vml_write_anchor(self, vml_obj);

    /* Write the <x:AutoFill> element. */
    _vml_write_auto_fill(self);

    /* Write the <x:Row> element. */
    _vml_write_row(self, vml_obj);

    /* Write the <x:Column> element. */
    _vml_write_column(self, vml_obj);

    /* Write the x:Visible element. */
    if (vml_obj->visible == LXW_COMMENT_DISPLAY_VISIBLE)
        _vml_write_visible(self);

    lxw_xml_end_tag(self->file, "x:ClientData");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <div> element.
 */
STATIC void
_vml_write_comment_div(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("style", "text-align:left");

    lxw_xml_start_tag(self->file, "div", &attributes);

    lxw_xml_end_tag(self->file, "div");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:textbox> element.
 */
STATIC void
_vml_write_comment_textbox(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("style", "mso-direction-alt:auto");

    lxw_xml_start_tag(self->file, "v:textbox", &attributes);

    /* Write the div element. */
    _vml_write_comment_div(self);

    lxw_xml_end_tag(self->file, "v:textbox");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:fill> element.
 */
STATIC void
_vml_write_comment_fill(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("color2", "#ffffe1");

    lxw_xml_empty_tag(self->file, "v:fill", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:path> element.
 */
STATIC void
_vml_write_comment_path(lxw_vml *self, uint8_t has_gradient, char *type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (has_gradient)
        LXW_PUSH_ATTRIBUTES_STR("gradientshapeok", "t");

    LXW_PUSH_ATTRIBUTES_STR("o:connecttype", type);

    lxw_xml_empty_tag(self->file, "v:path", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shape> element for comments.
 */
STATIC void
_vml_write_comment_shape(lxw_vml *self, uint32_t vml_shape_id,
                         uint32_t z_index, lxw_vml_obj *vml_obj)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char id[LXW_ATTR_32];
    char margin_left[LXW_ATTR_32];
    char margin_top[LXW_ATTR_32];
    char width[LXW_ATTR_32];
    char height[LXW_ATTR_32];
    char visible[LXW_ATTR_32];
    char fillcolor[LXW_ATTR_32];
    char style[LXW_MAX_ATTRIBUTE_LENGTH];
    char type[] = "#_x0000_t202";
    char o_insetmode[] = "auto";

    lxw_sprintf_dbl(margin_left, vml_obj->col_absolute * 0.75);
    lxw_sprintf_dbl(margin_top, vml_obj->row_absolute * 0.75);
    lxw_sprintf_dbl(width, vml_obj->width * 0.75);
    lxw_sprintf_dbl(height, vml_obj->height * 0.75);

    lxw_snprintf(id, LXW_ATTR_32, "_x0000_s%d", vml_shape_id);

    if (vml_obj->visible == LXW_COMMENT_DISPLAY_DEFAULT)
        vml_obj->visible = self->comment_display_default;

    if (vml_obj->visible == LXW_COMMENT_DISPLAY_VISIBLE)
        lxw_snprintf(visible, LXW_ATTR_32, "visible");
    else
        lxw_snprintf(visible, LXW_ATTR_32, "hidden");

    if (vml_obj->color)
        lxw_snprintf(fillcolor, LXW_ATTR_32, "#%06x",
                     vml_obj->color & LXW_COLOR_MASK);
    else
        lxw_snprintf(fillcolor, LXW_ATTR_32, "#%06x", 0xffffe1);

    lxw_snprintf(style,
                 LXW_MAX_ATTRIBUTE_LENGTH,
                 "position:absolute;"
                 "margin-left:%spt;"
                 "margin-top:%spt;"
                 "width:%spt;"
                 "height:%spt;"
                 "z-index:%d;"
                 "visibility:%s",
                 margin_left, margin_top, width, height, z_index, visible);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", id);
    LXW_PUSH_ATTRIBUTES_STR("type", type);
    LXW_PUSH_ATTRIBUTES_STR("style", style);
    LXW_PUSH_ATTRIBUTES_STR("fillcolor", fillcolor);
    LXW_PUSH_ATTRIBUTES_STR("o:insetmode", o_insetmode);

    lxw_xml_start_tag(self->file, "v:shape", &attributes);

    /* Write the v:fill element. */
    _vml_write_comment_fill(self);

    /* Write the v:shadow element. */
    _vml_write_shadow(self);

    /* Write the v:path element. */
    _vml_write_comment_path(self, LXW_FALSE, "none");

    /* Write the v:textbox element. */
    _vml_write_comment_textbox(self);

    /* Write the x:ClientData element. */
    _vml_write_comment_client_data(self, vml_obj);

    lxw_xml_end_tag(self->file, "v:shape");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <v:shapetype> element for comments.
 */
STATIC void
_vml_write_comment_shapetype(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char id[] = "_x0000_t202";
    char coordsize[] = "21600,21600";
    char o_spt[] = "202";
    char path[] = "m,l,21600r21600,l21600,xe";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("id", id);
    LXW_PUSH_ATTRIBUTES_STR("coordsize", coordsize);
    LXW_PUSH_ATTRIBUTES_STR("o:spt", o_spt);
    LXW_PUSH_ATTRIBUTES_STR("path", path);

    lxw_xml_start_tag(self->file, "v:shapetype", &attributes);

    /* Write the v:stroke element. */
    _vml_write_stroke(self);

    /* Write the v:path element. */
    _vml_write_comment_path(self, LXW_TRUE, "rect");

    lxw_xml_end_tag(self->file, "v:shapetype");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <o:idmap> element.
 */
STATIC void
_vml_write_idmap(lxw_vml *self)
{
    /* Since the vml_data_id_str may exceed the LXW_MAX_ATTRIBUTE_LENGTH we
     * write it directly without the xml helper functions. */
    fprintf(self->file, "<o:idmap v:ext=\"edit\" data=\"%s\"/>",
            self->vml_data_id_str);
}

/*
 * Write the <o:shapelayout> element.
 */
STATIC void
_vml_write_shapelayout(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("v:ext", "edit");

    lxw_xml_start_tag(self->file, "o:shapelayout", &attributes);

    /* Write the o:idmap element. */
    _vml_write_idmap(self);

    lxw_xml_end_tag(self->file, "o:shapelayout");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xml> element.
 */
STATIC void
_vml_write_xml_namespace(lxw_vml *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_v[] = "urn:schemas-microsoft-com:vml";
    char xmlns_o[] = "urn:schemas-microsoft-com:office:office";
    char xmlns_x[] = "urn:schemas-microsoft-com:office:excel";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:v", xmlns_v);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:o", xmlns_o);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:x", xmlns_x);

    lxw_xml_start_tag(self->file, "xml", &attributes);

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
lxw_vml_assemble_xml_file(lxw_vml *self)
{
    lxw_vml_obj *comment_obj;
    lxw_vml_obj *button_obj;
    lxw_vml_obj *image_obj;
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
            self->vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_button_shape(self, self->vml_shape_id, z_index,
                                    button_obj);

            z_index++;
        }
    }

    if (self->comment_objs && !STAILQ_EMPTY(self->comment_objs)) {
        /* Write the <v:shapetype> element. */
        _vml_write_comment_shapetype(self);

        STAILQ_FOREACH(comment_obj, self->comment_objs, list_pointers) {
            self->vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_comment_shape(self, self->vml_shape_id, z_index,
                                     comment_obj);

            z_index++;
        }
    }

    if (self->image_objs && !STAILQ_EMPTY(self->image_objs)) {
        /* Write the <v:shapetype> element. */
        _vml_write_image_shapetype(self);

        STAILQ_FOREACH(image_obj, self->image_objs, list_pointers) {
            self->vml_shape_id++;

            /* Write the <v:shape> element. */
            _vml_write_image_shape(self, self->vml_shape_id, z_index,
                                   image_obj);

            z_index++;
        }
    }

    lxw_xml_end_tag(self->file, "xml");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
