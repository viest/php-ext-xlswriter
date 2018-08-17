/*****************************************************************************
 * drawing - A library for creating Excel XLSX drawing files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/common.h"
#include "xlsxwriter/drawing.h"
#include "xlsxwriter/utility.h"

#define LXW_OBJ_NAME_LENGTH 14  /* "Picture 65536", or "Chart 65536" */
/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new drawing collection.
 */
lxw_drawing *
lxw_drawing_new()
{
    lxw_drawing *drawing = calloc(1, sizeof(lxw_drawing));
    GOTO_LABEL_ON_MEM_ERROR(drawing, mem_error);

    drawing->drawing_objects = calloc(1, sizeof(struct lxw_drawing_objects));
    GOTO_LABEL_ON_MEM_ERROR(drawing->drawing_objects, mem_error);

    STAILQ_INIT(drawing->drawing_objects);

    return drawing;

mem_error:
    lxw_drawing_free(drawing);
    return NULL;
}

/*
 * Free a drawing object.
 */
void
lxw_free_drawing_object(lxw_drawing_object *drawing_object)
{
    if (!drawing_object)
        return;

    free(drawing_object->description);
    free(drawing_object->url);
    free(drawing_object->tip);

    free(drawing_object);
}

/*
 * Free a drawing collection.
 */
void
lxw_drawing_free(lxw_drawing *drawing)
{
    lxw_drawing_object *drawing_object;

    if (!drawing)
        return;

    if (drawing->drawing_objects) {
        while (!STAILQ_EMPTY(drawing->drawing_objects)) {
            drawing_object = STAILQ_FIRST(drawing->drawing_objects);
            STAILQ_REMOVE_HEAD(drawing->drawing_objects, list_pointers);
            lxw_free_drawing_object(drawing_object);
        }

        free(drawing->drawing_objects);
    }

    free(drawing);
}

/*
 * Add a drawing object to the drawing collection.
 */
void
lxw_add_drawing_object(lxw_drawing *drawing,
                       lxw_drawing_object *drawing_object)
{
    STAILQ_INSERT_TAIL(drawing->drawing_objects, drawing_object,
                       list_pointers);
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
_drawing_xml_declaration(lxw_drawing *self)
{
    lxw_xml_declaration(self->file);
}

/*
 * Write the <xdr:wsDr> element.
 */
STATIC void
_write_drawing_workspace(lxw_drawing *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_xdr[] = LXW_SCHEMA_DRAWING "/spreadsheetDrawing";
    char xmlns_a[] = LXW_SCHEMA_DRAWING "/main";

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("xmlns:xdr", xmlns_xdr);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:a", xmlns_a);

    lxw_xml_start_tag(self->file, "xdr:wsDr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:col> element.
 */
STATIC void
_drawing_write_col(lxw_drawing *self, char *data)
{
    lxw_xml_data_element(self->file, "xdr:col", data, NULL);
}

/*
 * Write the <xdr:colOff> element.
 */
STATIC void
_drawing_write_col_off(lxw_drawing *self, char *data)
{
    lxw_xml_data_element(self->file, "xdr:colOff", data, NULL);
}

/*
 * Write the <xdr:row> element.
 */
STATIC void
_drawing_write_row(lxw_drawing *self, char *data)
{
    lxw_xml_data_element(self->file, "xdr:row", data, NULL);
}

/*
 * Write the <xdr:rowOff> element.
 */
STATIC void
_drawing_write_row_off(lxw_drawing *self, char *data)
{
    lxw_xml_data_element(self->file, "xdr:rowOff", data, NULL);
}

/*
 * Write the <xdr:from> element.
 */
STATIC void
_drawing_write_from(lxw_drawing *self, lxw_drawing_coords *coords)
{
    char data[LXW_UINT32_T_LENGTH];

    lxw_xml_start_tag(self->file, "xdr:from", NULL);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u", coords->col);
    _drawing_write_col(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->col_offset);
    _drawing_write_col_off(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u", coords->row);
    _drawing_write_row(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->row_offset);
    _drawing_write_row_off(self, data);

    lxw_xml_end_tag(self->file, "xdr:from");
}

/*
 * Write the <xdr:to> element.
 */
STATIC void
_drawing_write_to(lxw_drawing *self, lxw_drawing_coords *coords)
{
    char data[LXW_UINT32_T_LENGTH];

    lxw_xml_start_tag(self->file, "xdr:to", NULL);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u", coords->col);
    _drawing_write_col(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->col_offset);
    _drawing_write_col_off(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u", coords->row);
    _drawing_write_row(self, data);

    lxw_snprintf(data, LXW_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->row_offset);
    _drawing_write_row_off(self, data);

    lxw_xml_end_tag(self->file, "xdr:to");
}

/*
 * Write the <xdr:cNvPr> element.
 */
STATIC void
_drawing_write_c_nv_pr(lxw_drawing *self, char *object_name, uint16_t index,
                       lxw_drawing_object *drawing_object)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    char name[LXW_OBJ_NAME_LENGTH];
    lxw_snprintf(name, LXW_OBJ_NAME_LENGTH, "%s %d", object_name, index);

    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_INT("id", index + 1);
    LXW_PUSH_ATTRIBUTES_STR("name", name);

    if (drawing_object)
        LXW_PUSH_ATTRIBUTES_STR("descr", drawing_object->description);

    lxw_xml_empty_tag(self->file, "xdr:cNvPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:picLocks> element.
 */
STATIC void
_drawing_write_a_pic_locks(lxw_drawing *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("noChangeAspect", "1");

    lxw_xml_empty_tag(self->file, "a:picLocks", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:cNvPicPr> element.
 */
STATIC void
_drawing_write_c_nv_pic_pr(lxw_drawing *self)
{
    lxw_xml_start_tag(self->file, "xdr:cNvPicPr", NULL);

    /* Write the a:picLocks element. */
    _drawing_write_a_pic_locks(self);

    lxw_xml_end_tag(self->file, "xdr:cNvPicPr");
}

/*
 * Write the <xdr:nvPicPr> element.
 */
STATIC void
_drawing_write_nv_pic_pr(lxw_drawing *self, uint16_t index,
                         lxw_drawing_object *drawing_object)
{
    lxw_xml_start_tag(self->file, "xdr:nvPicPr", NULL);

    /* Write the xdr:cNvPr element. */
    _drawing_write_c_nv_pr(self, "Picture", index, drawing_object);

    /* Write the xdr:cNvPicPr element. */
    _drawing_write_c_nv_pic_pr(self);

    lxw_xml_end_tag(self->file, "xdr:nvPicPr");
}

/*
 * Write the <a:blip> element.
 */
STATIC void
_drawing_write_a_blip(lxw_drawing *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);
    LXW_PUSH_ATTRIBUTES_STR("r:embed", r_id);

    lxw_xml_empty_tag(self->file, "a:blip", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:fillRect> element.
 */
STATIC void
_drawing_write_a_fill_rect(lxw_drawing *self)
{
    lxw_xml_empty_tag(self->file, "a:fillRect", NULL);
}

/*
 * Write the <a:stretch> element.
 */
STATIC void
_drawing_write_a_stretch(lxw_drawing *self)
{
    lxw_xml_start_tag(self->file, "a:stretch", NULL);

    /* Write the a:fillRect element. */
    _drawing_write_a_fill_rect(self);

    lxw_xml_end_tag(self->file, "a:stretch");
}

/*
 * Write the <xdr:blipFill> element.
 */
STATIC void
_drawing_write_blip_fill(lxw_drawing *self, uint16_t index)
{
    lxw_xml_start_tag(self->file, "xdr:blipFill", NULL);

    /* Write the a:blip element. */
    _drawing_write_a_blip(self, index);

    /* Write the a:stretch element. */
    _drawing_write_a_stretch(self);

    lxw_xml_end_tag(self->file, "xdr:blipFill");
}

/*
 * Write the <a:ext> element.
 */
STATIC void
_drawing_write_a_ext(lxw_drawing *self, lxw_drawing_object *drawing_object)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("cx", drawing_object->width);
    LXW_PUSH_ATTRIBUTES_INT("cy", drawing_object->height);

    lxw_xml_empty_tag(self->file, "a:ext", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:off> element.
 */
STATIC void
_drawing_write_a_off(lxw_drawing *self, lxw_drawing_object *drawing_object)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("x", drawing_object->col_absolute);
    LXW_PUSH_ATTRIBUTES_INT("y", drawing_object->row_absolute);

    lxw_xml_empty_tag(self->file, "a:off", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:xfrm> element.
 */
STATIC void
_drawing_write_a_xfrm(lxw_drawing *self, lxw_drawing_object *drawing_object)
{
    lxw_xml_start_tag(self->file, "a:xfrm", NULL);

    /* Write the a:off element. */
    _drawing_write_a_off(self, drawing_object);

    /* Write the a:ext element. */
    _drawing_write_a_ext(self, drawing_object);

    lxw_xml_end_tag(self->file, "a:xfrm");
}

/*
 * Write the <a:avLst> element.
 */
STATIC void
_drawing_write_a_av_lst(lxw_drawing *self)
{
    lxw_xml_empty_tag(self->file, "a:avLst", NULL);
}

/*
 * Write the <a:prstGeom> element.
 */
STATIC void
_drawing_write_a_prst_geom(lxw_drawing *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("prst", "rect");

    lxw_xml_start_tag(self->file, "a:prstGeom", &attributes);

    /* Write the a:avLst element. */
    _drawing_write_a_av_lst(self);

    lxw_xml_end_tag(self->file, "a:prstGeom");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:spPr> element.
 */
STATIC void
_drawing_write_sp_pr(lxw_drawing *self, lxw_drawing_object *drawing_object)
{
    lxw_xml_start_tag(self->file, "xdr:spPr", NULL);

    /* Write the a:xfrm element. */
    _drawing_write_a_xfrm(self, drawing_object);

    /* Write the a:prstGeom element. */
    _drawing_write_a_prst_geom(self);

    lxw_xml_end_tag(self->file, "xdr:spPr");
}

/*
 * Write the <xdr:pic> element.
 */
STATIC void
_drawing_write_pic(lxw_drawing *self, uint16_t index,
                   lxw_drawing_object *drawing_object)
{
    lxw_xml_start_tag(self->file, "xdr:pic", NULL);

    /* Write the xdr:nvPicPr element. */
    _drawing_write_nv_pic_pr(self, index, drawing_object);

    /* Write the xdr:blipFill element. */
    _drawing_write_blip_fill(self, index);

    /* Write the xdr:spPr element. */
    _drawing_write_sp_pr(self, drawing_object);

    lxw_xml_end_tag(self->file, "xdr:pic");
}

/*
 * Write the <xdr:clientData> element.
 */
STATIC void
_drawing_write_client_data(lxw_drawing *self)
{
    lxw_xml_empty_tag(self->file, "xdr:clientData", NULL);
}

/*
 * Write the <xdr:cNvGraphicFramePr> element.
 */
STATIC void
_drawing_write_c_nv_graphic_frame_pr(lxw_drawing *self)
{
    lxw_xml_empty_tag(self->file, "xdr:cNvGraphicFramePr", NULL);
}

/*
 * Write the <xdr:nvGraphicFramePr> element.
 */
STATIC void
_drawing_write_nv_graphic_frame_pr(lxw_drawing *self, uint16_t index)
{
    lxw_xml_start_tag(self->file, "xdr:nvGraphicFramePr", NULL);

    /* Write the xdr:cNvPr element. */
    _drawing_write_c_nv_pr(self, "Chart", index, NULL);

    /* Write the xdr:cNvGraphicFramePr element. */
    _drawing_write_c_nv_graphic_frame_pr(self);

    lxw_xml_end_tag(self->file, "xdr:nvGraphicFramePr");
}

/*
 * Write the <a:off> element.
 */
STATIC void
_drawing_write_xfrm_offset(lxw_drawing *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("x", "0");
    LXW_PUSH_ATTRIBUTES_STR("y", "0");

    lxw_xml_empty_tag(self->file, "a:off", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:ext> element.
 */
STATIC void
_drawing_write_xfrm_extension(lxw_drawing *self)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("cx", "0");
    LXW_PUSH_ATTRIBUTES_STR("cy", "0");

    lxw_xml_empty_tag(self->file, "a:ext", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:xfrm> element.
 */
STATIC void
_drawing_write_xfrm(lxw_drawing *self)
{
    lxw_xml_start_tag(self->file, "xdr:xfrm", NULL);

    /* Write the a:off element. */
    _drawing_write_xfrm_offset(self);

    /* Write the a:ext element. */
    _drawing_write_xfrm_extension(self);

    lxw_xml_end_tag(self->file, "xdr:xfrm");
}

/*
 * Write the <c:chart> element.
 */
STATIC void
_drawing_write_chart(lxw_drawing *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_c[] = LXW_SCHEMA_DRAWING "/chart";
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", index);

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:c", xmlns_c);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);
    LXW_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxw_xml_empty_tag(self->file, "c:chart", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:graphicData> element.
 */
STATIC void
_drawing_write_a_graphic_data(lxw_drawing *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char uri[] = LXW_SCHEMA_DRAWING "/chart";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("uri", uri);

    lxw_xml_start_tag(self->file, "a:graphicData", &attributes);

    /* Write the c:chart element. */
    _drawing_write_chart(self, index);

    lxw_xml_end_tag(self->file, "a:graphicData");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:graphic> element.
 */
STATIC void
_drawing_write_a_graphic(lxw_drawing *self, uint16_t index)
{

    lxw_xml_start_tag(self->file, "a:graphic", NULL);

    /* Write the a:graphicData element. */
    _drawing_write_a_graphic_data(self, index);

    lxw_xml_end_tag(self->file, "a:graphic");
}

/*
 * Write the <xdr:graphicFrame> element.
 */
STATIC void
_drawing_write_graphic_frame(lxw_drawing *self, uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("macro", "");

    lxw_xml_start_tag(self->file, "xdr:graphicFrame", &attributes);

    /* Write the xdr:nvGraphicFramePr element. */
    _drawing_write_nv_graphic_frame_pr(self, index);

    /* Write the xdr:xfrm element. */
    _drawing_write_xfrm(self);

    /* Write the a:graphic element. */
    _drawing_write_a_graphic(self, index);

    lxw_xml_end_tag(self->file, "xdr:graphicFrame");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:twoCellAnchor> element.
 */
STATIC void
_drawing_write_two_cell_anchor(lxw_drawing *self, uint16_t index,
                               lxw_drawing_object *drawing_object)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_IMAGE) {

        if (drawing_object->edit_as == LXW_ANCHOR_EDIT_AS_ABSOLUTE)
            LXW_PUSH_ATTRIBUTES_STR("editAs", "absolute");
        else if (drawing_object->edit_as != LXW_ANCHOR_EDIT_AS_RELATIVE)
            LXW_PUSH_ATTRIBUTES_STR("editAs", "oneCell");
    }

    lxw_xml_start_tag(self->file, "xdr:twoCellAnchor", &attributes);

    _drawing_write_from(self, &drawing_object->from);
    _drawing_write_to(self, &drawing_object->to);

    if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_CHART) {
        /* Write the xdr:graphicFrame element for charts. */
        _drawing_write_graphic_frame(self, index);
    }
    else if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_IMAGE) {
        /* Write the xdr:pic element. */
        _drawing_write_pic(self, index, drawing_object);
    }
    else {
        /* Write the xdr:sp element for shapes. */
        /* _drawing_write_sp(self, index, col_absolute, row_absolute, width,
           height,  shape); */
    }

    /* Write the xdr:clientData element. */
    _drawing_write_client_data(self);

    lxw_xml_end_tag(self->file, "xdr:twoCellAnchor");

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
lxw_drawing_assemble_xml_file(lxw_drawing *self)
{
    uint16_t index;
    lxw_drawing_object *drawing_object;

    /* Write the XML declaration. */
    _drawing_xml_declaration(self);

    /* Write the xdr:wsDr element. */
    _write_drawing_workspace(self);

    if (self->embedded) {
        index = 1;

        STAILQ_FOREACH(drawing_object, self->drawing_objects, list_pointers) {
            _drawing_write_two_cell_anchor(self, index, drawing_object);
            index++;
        }

    }

    lxw_xml_end_tag(self->file, "xdr:wsDr");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
