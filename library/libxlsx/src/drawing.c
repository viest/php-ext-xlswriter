/*****************************************************************************
 * drawing - A library for creating Excel XLSX drawing files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "libxlsx/xmlwriter.h"
#include "libxlsx/common.h"
#include "libxlsx/drawing.h"
#include "libxlsx/worksheet.h"
#include "libxlsx/utility.h"

#define LXLSX_OBJ_NAME_LENGTH 14  /* "Picture 65536", or "Chart 65536" */
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
lxlsx_drawing *
lxlsx_drawing_new(void)
{
    lxlsx_drawing *drawing = calloc(1, sizeof(lxlsx_drawing));
    GOTO_LABEL_ON_MEM_ERROR(drawing, mem_error);

    drawing->lxlsx_drawing_objects = calloc(1, sizeof(struct lxlsx_drawing_objects));
    GOTO_LABEL_ON_MEM_ERROR(drawing->lxlsx_drawing_objects, mem_error);

    STAILQ_INIT(drawing->lxlsx_drawing_objects);

    return drawing;

mem_error:
    lxlsx_drawing_free(drawing);
    return NULL;
}

/*
 * Free a drawing object.
 */
void
lxlsx_free_drawing_object(lxlsx_drawing_object *lxlsx_drawing_object)
{
    if (!lxlsx_drawing_object)
        return;

    free(lxlsx_drawing_object->description);
    free(lxlsx_drawing_object->tip);

    free(lxlsx_drawing_object);
}

/*
 * Free a drawing collection.
 */
void
lxlsx_drawing_free(lxlsx_drawing *drawing)
{
    lxlsx_drawing_object *lxlsx_drawing_object;

    if (!drawing)
        return;

    if (drawing->lxlsx_drawing_objects) {
        while (!STAILQ_EMPTY(drawing->lxlsx_drawing_objects)) {
            lxlsx_drawing_object = STAILQ_FIRST(drawing->lxlsx_drawing_objects);
            STAILQ_REMOVE_HEAD(drawing->lxlsx_drawing_objects, list_pointers);
            lxlsx_free_drawing_object(lxlsx_drawing_object);
        }

        free(drawing->lxlsx_drawing_objects);
    }

    free(drawing);
}

/*
 * Add a drawing object to the drawing collection.
 */
void
lxlsx_add_drawing_object(lxlsx_drawing *drawing,
                       lxlsx_drawing_object *lxlsx_drawing_object)
{
    STAILQ_INSERT_TAIL(drawing->lxlsx_drawing_objects, lxlsx_drawing_object,
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
LXLSX_DEFINE_XML_DECLARATION(_drawing_xml_declaration, lxlsx_drawing)

/*
 * Write the <xdr:wsDr> element.
 */
STATIC void
_write_drawing_workspace(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_xdr[] = LXLSX_SCHEMA_DRAWING "/spreadsheetDrawing";
    char xmlns_a[] = LXLSX_SCHEMA_DRAWING "/main";

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:xdr", xmlns_xdr);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:a", xmlns_a);

    lxlsx_xml_start_tag(self->file, "xdr:wsDr", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:col> element.
 */
STATIC void
_drawing_write_col(lxlsx_drawing *self, char *data)
{
    lxlsx_xml_data_element(self->file, "xdr:col", data, NULL);
}

/*
 * Write the <xdr:colOff> element.
 */
STATIC void
_drawing_write_col_off(lxlsx_drawing *self, char *data)
{
    lxlsx_xml_data_element(self->file, "xdr:colOff", data, NULL);
}

/*
 * Write the <xdr:row> element.
 */
STATIC void
_drawing_write_row(lxlsx_drawing *self, char *data)
{
    lxlsx_xml_data_element(self->file, "xdr:row", data, NULL);
}

/*
 * Write the <xdr:rowOff> element.
 */
STATIC void
_drawing_write_row_off(lxlsx_drawing *self, char *data)
{
    lxlsx_xml_data_element(self->file, "xdr:rowOff", data, NULL);
}

/*
 * Write the main part of the <xdr:from> and <xdr:to> elements.
 */
STATIC void
_drawing_write_coords(lxlsx_drawing *self, lxlsx_drawing_coords *coords)
{
    char data[LXLSX_UINT32_T_LENGTH];

    lxlsx_snprintf(data, LXLSX_UINT32_T_LENGTH, "%u", coords->col);
    _drawing_write_col(self, data);

    lxlsx_snprintf(data, LXLSX_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->col_offset);
    _drawing_write_col_off(self, data);

    lxlsx_snprintf(data, LXLSX_UINT32_T_LENGTH, "%u", coords->row);
    _drawing_write_row(self, data);

    lxlsx_snprintf(data, LXLSX_UINT32_T_LENGTH, "%u",
                 (uint32_t) coords->row_offset);
    _drawing_write_row_off(self, data);
}

/*
 * Write the <xdr:from> element.
 */
STATIC void
_drawing_write_from(lxlsx_drawing *self, lxlsx_drawing_coords *coords)
{
    lxlsx_xml_start_tag(self->file, "xdr:from", NULL);

    _drawing_write_coords(self, coords);

    lxlsx_xml_end_tag(self->file, "xdr:from");
}

/*
 * Write the <xdr:to> element.
 */
STATIC void
_drawing_write_to(lxlsx_drawing *self, lxlsx_drawing_coords *coords)
{
    lxlsx_xml_start_tag(self->file, "xdr:to", NULL);

    _drawing_write_coords(self, coords);

    lxlsx_xml_end_tag(self->file, "xdr:to");
}

/*
 * Write the <a:hlinkClick> element.
 */
STATIC void
_drawing_write_a_hlink_click(lxlsx_drawing *self, uint32_t rel_index, char *tip)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_r[] = "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships";
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", rel_index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    if (tip)
        LXLSX_PUSH_ATTRIBUTES_STR("tooltip", tip);

    lxlsx_xml_empty_tag(self->file, "a:hlinkClick", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a16:creationId> element.
 */
STATIC void
_drawing_write_a16_creation_id(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] = "http://schemas.microsoft.com/office/drawing/2014/main";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:a16", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("id", "{00000000-0008-0000-0000-000002000000}");

    lxlsx_xml_empty_tag(self->file, "a16:creationId", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <adec:decorative> element.
 */
STATIC void
_workbook_write_adec_decorative(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns[] =
        "http://schemas.microsoft.com/office/drawing/2017/decorative";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:adec", xmlns);
    LXLSX_PUSH_ATTRIBUTES_STR("val", "1");

    lxlsx_xml_empty_tag(self->file, "adec:decorative", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:ext> element.
 */
STATIC void
_drawing_write_uri_ext(lxlsx_drawing *self, char *uri)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("uri", uri);

    lxlsx_xml_start_tag(self->file, "a:ext", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the decorative elements.
 */
STATIC void
_workbook_write_decorative(lxlsx_drawing *self)
{
    lxlsx_xml_start_tag(self->file, "a:extLst", NULL);

    _drawing_write_uri_ext(self, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}");
    _drawing_write_a16_creation_id(self);
    lxlsx_xml_end_tag(self->file, "a:ext");

    _drawing_write_uri_ext(self, "{C183D7F6-B498-43B3-948B-1728B52AA6E4}");
    _workbook_write_adec_decorative(self);
    lxlsx_xml_end_tag(self->file, "a:ext");

    lxlsx_xml_end_tag(self->file, "a:extLst");
}

/*
 * Write the <xdr:cNvPr> element.
 */
STATIC void
_drawing_write_c_nv_pr(lxlsx_drawing *self, char *object_name, uint32_t index,
                       lxlsx_drawing_object *lxlsx_drawing_object)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    char name[LXLSX_OBJ_NAME_LENGTH];
    lxlsx_snprintf(name, LXLSX_OBJ_NAME_LENGTH, "%s %d", object_name, index);

    LXLSX_INIT_ATTRIBUTES();

    LXLSX_PUSH_ATTRIBUTES_INT("id", index + 1);
    LXLSX_PUSH_ATTRIBUTES_STR("name", name);

    if (lxlsx_drawing_object && lxlsx_drawing_object->description
        && strlen(lxlsx_drawing_object->description)
        && !lxlsx_drawing_object->decorative) {

        LXLSX_PUSH_ATTRIBUTES_STR("descr", lxlsx_drawing_object->description);
    }

    if (lxlsx_drawing_object
        && (lxlsx_drawing_object->url_rel_index || lxlsx_drawing_object->decorative)) {
        lxlsx_xml_start_tag(self->file, "xdr:cNvPr", &attributes);

        if (lxlsx_drawing_object->url_rel_index) {
            /* Write the a:hlinkClick element. */
            _drawing_write_a_hlink_click(self,
                                         lxlsx_drawing_object->url_rel_index,
                                         lxlsx_drawing_object->tip);
        }

        if (lxlsx_drawing_object->decorative) {
            _workbook_write_decorative(self);
        }

        lxlsx_xml_end_tag(self->file, "xdr:cNvPr");
    }
    else {
        lxlsx_xml_empty_tag(self->file, "xdr:cNvPr", &attributes);
    }

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:picLocks> element.
 */
STATIC void
_drawing_write_a_pic_locks(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("noChangeAspect", "1");

    lxlsx_xml_empty_tag(self->file, "a:picLocks", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:cNvPicPr> element.
 */
STATIC void
_drawing_write_c_nv_pic_pr(lxlsx_drawing *self)
{
    lxlsx_xml_start_tag(self->file, "xdr:cNvPicPr", NULL);

    /* Write the a:picLocks element. */
    _drawing_write_a_pic_locks(self);

    lxlsx_xml_end_tag(self->file, "xdr:cNvPicPr");
}

/*
 * Write the <xdr:nvPicPr> element.
 */
STATIC void
_drawing_write_nv_pic_pr(lxlsx_drawing *self, uint32_t index,
                         lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_xml_start_tag(self->file, "xdr:nvPicPr", NULL);

    /* Write the xdr:cNvPr element. */
    _drawing_write_c_nv_pr(self, "Picture", index, lxlsx_drawing_object);

    /* Write the xdr:cNvPicPr element. */
    _drawing_write_c_nv_pic_pr(self);

    lxlsx_xml_end_tag(self->file, "xdr:nvPicPr");
}

/*
 * Write the <a:blip> element.
 */
STATIC void
_drawing_write_a_blip(lxlsx_drawing *self, uint32_t index)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_r[] = LXLSX_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);
    LXLSX_PUSH_ATTRIBUTES_STR("r:embed", r_id);

    lxlsx_xml_empty_tag(self->file, "a:blip", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:fillRect> element.
 */
STATIC void
_drawing_write_a_fill_rect(lxlsx_drawing *self)
{
    lxlsx_xml_empty_tag(self->file, "a:fillRect", NULL);
}

/*
 * Write the <a:stretch> element.
 */
STATIC void
_drawing_write_a_stretch(lxlsx_drawing *self)
{
    lxlsx_xml_start_tag(self->file, "a:stretch", NULL);

    /* Write the a:fillRect element. */
    _drawing_write_a_fill_rect(self);

    lxlsx_xml_end_tag(self->file, "a:stretch");
}

/*
 * Write the <xdr:blipFill> element.
 */
STATIC void
_drawing_write_blip_fill(lxlsx_drawing *self, uint32_t index)
{
    lxlsx_xml_start_tag(self->file, "xdr:blipFill", NULL);

    /* Write the a:blip element. */
    _drawing_write_a_blip(self, index);

    /* Write the a:stretch element. */
    _drawing_write_a_stretch(self);

    lxlsx_xml_end_tag(self->file, "xdr:blipFill");
}

/*
 * Write the <a:ext> element.
 */
STATIC void
_drawing_write_a_ext(lxlsx_drawing *self, lxlsx_drawing_object *lxlsx_drawing_object)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("cx", lxlsx_drawing_object->width);
    LXLSX_PUSH_ATTRIBUTES_INT("cy", lxlsx_drawing_object->height);

    lxlsx_xml_empty_tag(self->file, "a:ext", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:off> element.
 */
STATIC void
_drawing_write_a_off(lxlsx_drawing *self, lxlsx_drawing_object *lxlsx_drawing_object)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    /* Use %d to allow for writing uint64_t on ansi 32bit systems. The largest
       Excel value will fit without losing precision. */
    LXLSX_PUSH_ATTRIBUTES_DBL("x", lxlsx_drawing_object->col_absolute);
    LXLSX_PUSH_ATTRIBUTES_DBL("y", lxlsx_drawing_object->row_absolute);

    lxlsx_xml_empty_tag(self->file, "a:off", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:xfrm> element.
 */
STATIC void
_drawing_write_a_xfrm(lxlsx_drawing *self, lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_xml_start_tag(self->file, "a:xfrm", NULL);

    /* Write the a:off element. */
    _drawing_write_a_off(self, lxlsx_drawing_object);

    /* Write the a:ext element. */
    _drawing_write_a_ext(self, lxlsx_drawing_object);

    lxlsx_xml_end_tag(self->file, "a:xfrm");
}

/*
 * Write the <a:avLst> element.
 */
STATIC void
_drawing_write_a_av_lst(lxlsx_drawing *self)
{
    lxlsx_xml_empty_tag(self->file, "a:avLst", NULL);
}

/*
 * Write the <a:prstGeom> element.
 */
STATIC void
_drawing_write_a_prst_geom(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("prst", "rect");

    lxlsx_xml_start_tag(self->file, "a:prstGeom", &attributes);

    /* Write the a:avLst element. */
    _drawing_write_a_av_lst(self);

    lxlsx_xml_end_tag(self->file, "a:prstGeom");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:spPr> element.
 */
STATIC void
_drawing_write_sp_pr(lxlsx_drawing *self, lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_xml_start_tag(self->file, "xdr:spPr", NULL);

    /* Write the a:xfrm element. */
    _drawing_write_a_xfrm(self, lxlsx_drawing_object);

    /* Write the a:prstGeom element. */
    _drawing_write_a_prst_geom(self);

    lxlsx_xml_end_tag(self->file, "xdr:spPr");
}

/*
 * Write the <xdr:pic> element.
 */
STATIC void
_drawing_write_pic(lxlsx_drawing *self, uint32_t index,
                   lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_xml_start_tag(self->file, "xdr:pic", NULL);

    /* Write the xdr:nvPicPr element. */
    _drawing_write_nv_pic_pr(self, index, lxlsx_drawing_object);

    /* Write the xdr:blipFill element. */
    _drawing_write_blip_fill(self, lxlsx_drawing_object->rel_index);

    /* Write the xdr:spPr element. */
    _drawing_write_sp_pr(self, lxlsx_drawing_object);

    lxlsx_xml_end_tag(self->file, "xdr:pic");
}

/*
 * Write the <xdr:clientData> element.
 */
STATIC void
_drawing_write_client_data(lxlsx_drawing *self)
{
    lxlsx_xml_empty_tag(self->file, "xdr:clientData", NULL);
}

/*
 * Write the <a:graphicFrameLocks> element.
 */
STATIC void
_drawing_write_a_graphic_frame_locks(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("noGrp", 1);

    lxlsx_xml_empty_tag(self->file, "a:graphicFrameLocks", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:cNvGraphicFramePr> element.
 */
STATIC void
_drawing_write_c_nv_graphic_frame_pr(lxlsx_drawing *self)
{
    if (self->embedded) {
        lxlsx_xml_empty_tag(self->file, "xdr:cNvGraphicFramePr", NULL);
    }
    else {
        lxlsx_xml_start_tag(self->file, "xdr:cNvGraphicFramePr", NULL);

        /* Write the a:graphicFrameLocks element. */
        _drawing_write_a_graphic_frame_locks(self);

        lxlsx_xml_end_tag(self->file, "xdr:cNvGraphicFramePr");
    }
}

/*
 * Write the <xdr:nvGraphicFramePr> element.
 */
STATIC void
_drawing_write_nv_graphic_frame_pr(lxlsx_drawing *self, uint32_t index,
                                   lxlsx_drawing_object *lxlsx_drawing_object)
{
    lxlsx_xml_start_tag(self->file, "xdr:nvGraphicFramePr", NULL);

    /* Write the xdr:cNvPr element. */
    _drawing_write_c_nv_pr(self, "Chart", index, lxlsx_drawing_object);

    /* Write the xdr:cNvGraphicFramePr element. */
    _drawing_write_c_nv_graphic_frame_pr(self);

    lxlsx_xml_end_tag(self->file, "xdr:nvGraphicFramePr");
}

/*
 * Write the <a:off> element.
 */
STATIC void
_drawing_write_xfrm_offset(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("x", "0");
    LXLSX_PUSH_ATTRIBUTES_STR("y", "0");

    lxlsx_xml_empty_tag(self->file, "a:off", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:ext> element.
 */
STATIC void
_drawing_write_xfrm_extension(lxlsx_drawing *self)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("cx", "0");
    LXLSX_PUSH_ATTRIBUTES_STR("cy", "0");

    lxlsx_xml_empty_tag(self->file, "a:ext", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:xfrm> element.
 */
STATIC void
_drawing_write_xfrm(lxlsx_drawing *self)
{
    lxlsx_xml_start_tag(self->file, "xdr:xfrm", NULL);

    /* Write the a:off element. */
    _drawing_write_xfrm_offset(self);

    /* Write the a:ext element. */
    _drawing_write_xfrm_extension(self);

    lxlsx_xml_end_tag(self->file, "xdr:xfrm");
}

/*
 * Write the <c:chart> element.
 */
STATIC void
_drawing_write_chart(lxlsx_drawing *self, uint32_t index)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char xmlns_c[] = LXLSX_SCHEMA_DRAWING "/chart";
    char xmlns_r[] = LXLSX_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXLSX_MAX_ATTRIBUTE_LENGTH];

    lxlsx_snprintf(r_id, LXLSX_ATTR_32, "rId%d", index);

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:c", xmlns_c);
    LXLSX_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);
    LXLSX_PUSH_ATTRIBUTES_STR("r:id", r_id);

    lxlsx_xml_empty_tag(self->file, "c:chart", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:graphicData> element.
 */
STATIC void
_drawing_write_a_graphic_data(lxlsx_drawing *self, uint32_t index)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;
    char uri[] = LXLSX_SCHEMA_DRAWING "/chart";

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("uri", uri);

    lxlsx_xml_start_tag(self->file, "a:graphicData", &attributes);

    /* Write the c:chart element. */
    _drawing_write_chart(self, index);

    lxlsx_xml_end_tag(self->file, "a:graphicData");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <a:graphic> element.
 */
STATIC void
_drawing_write_a_graphic(lxlsx_drawing *self, uint32_t index)
{

    lxlsx_xml_start_tag(self->file, "a:graphic", NULL);

    /* Write the a:graphicData element. */
    _drawing_write_a_graphic_data(self, index);

    lxlsx_xml_end_tag(self->file, "a:graphic");
}

/*
 * Write the <xdr:graphicFrame> element.
 */
STATIC void
_drawing_write_graphic_frame(lxlsx_drawing *self, uint32_t index,
                             uint32_t rel_index,
                             lxlsx_drawing_object *lxlsx_drawing_object)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_STR("macro", "");

    lxlsx_xml_start_tag(self->file, "xdr:graphicFrame", &attributes);

    /* Write the xdr:nvGraphicFramePr element. */
    _drawing_write_nv_graphic_frame_pr(self, index, lxlsx_drawing_object);

    /* Write the xdr:xfrm element. */
    _drawing_write_xfrm(self);

    /* Write the a:graphic element. */
    _drawing_write_a_graphic(self, rel_index);

    lxlsx_xml_end_tag(self->file, "xdr:graphicFrame");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:twoCellAnchor> element.
 */
STATIC void
_drawing_write_two_cell_anchor(lxlsx_drawing *self, uint32_t index,
                               lxlsx_drawing_object *lxlsx_drawing_object)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();

    if (lxlsx_drawing_object->anchor == LXLSX_OBJECT_MOVE_DONT_SIZE)
        LXLSX_PUSH_ATTRIBUTES_STR("editAs", "oneCell");
    else if (lxlsx_drawing_object->anchor == LXLSX_OBJECT_DONT_MOVE_DONT_SIZE)
        LXLSX_PUSH_ATTRIBUTES_STR("editAs", "absolute");

    lxlsx_xml_start_tag(self->file, "xdr:twoCellAnchor", &attributes);

    _drawing_write_from(self, &lxlsx_drawing_object->from);
    _drawing_write_to(self, &lxlsx_drawing_object->to);

    if (lxlsx_drawing_object->type == LXLSX_DRAWING_CHART) {
        /* Write the xdr:graphicFrame element for charts. */
        _drawing_write_graphic_frame(self, index, lxlsx_drawing_object->rel_index,
                                     lxlsx_drawing_object);
    }
    else if (lxlsx_drawing_object->type == LXLSX_DRAWING_IMAGE) {
        /* Write the xdr:pic element. */
        _drawing_write_pic(self, index, lxlsx_drawing_object);
    }
    else {
        /* Write the xdr:sp element for shapes. */
        /* _drawing_write_sp(self, index, col_absolute, row_absolute, width,
           height,  shape); */
    }

    /* Write the xdr:clientData element. */
    _drawing_write_client_data(self);

    lxlsx_xml_end_tag(self->file, "xdr:twoCellAnchor");

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:ext> element.
 */
STATIC void
_drawing_write_ext(lxlsx_drawing *self, uint32_t cx, uint32_t cy)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("cx", cx);
    LXLSX_PUSH_ATTRIBUTES_INT("cy", cy);

    lxlsx_xml_empty_tag(self->file, "xdr:ext", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:pos> element.
 */
STATIC void
_drawing_write_pos(lxlsx_drawing *self, int32_t x, int32_t y)
{
    struct lxlsx_xml_attribute_list attributes;
    struct lxlsx_xml_attribute *attribute;

    LXLSX_INIT_ATTRIBUTES();
    LXLSX_PUSH_ATTRIBUTES_INT("x", x);
    LXLSX_PUSH_ATTRIBUTES_INT("y", y);

    lxlsx_xml_empty_tag(self->file, "xdr:pos", &attributes);

    LXLSX_FREE_ATTRIBUTES();
}

/*
 * Write the <xdr:absoluteAnchor> element.
 */
STATIC void
_drawing_write_absolute_anchor(lxlsx_drawing *self, uint32_t frame_index)
{
    lxlsx_xml_start_tag(self->file, "xdr:absoluteAnchor", NULL);

    if (self->orientation == LXLSX_LANDSCAPE) {
        /* Write the xdr:pos element. */
        _drawing_write_pos(self, 0, 0);

        /* Write the xdr:ext element. */
        _drawing_write_ext(self, 9308969, 6078325);
    }
    else {
        /* Write the xdr:pos element. */
        _drawing_write_pos(self, 0, -47625);

        /* Write the xdr:ext element. */
        _drawing_write_ext(self, 6162675, 6124575);
    }

    _drawing_write_graphic_frame(self, frame_index, frame_index, NULL);

    /* Write the xdr:clientData element. */
    _drawing_write_client_data(self);

    lxlsx_xml_end_tag(self->file, "xdr:absoluteAnchor");
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
lxlsx_drawing_assemble_xml_file(lxlsx_drawing *self)
{
    uint32_t index;
    lxlsx_drawing_object *lxlsx_drawing_object;

    /* Write the XML declaration. */
    _drawing_xml_declaration(self);

    /* Write the xdr:wsDr element. */
    _write_drawing_workspace(self);

    if (self->embedded) {
        index = 1;

        STAILQ_FOREACH(lxlsx_drawing_object, self->lxlsx_drawing_objects, list_pointers) {
            _drawing_write_two_cell_anchor(self, index, lxlsx_drawing_object);
            index++;
        }
    }
    else {
        /* Write the xdr:absoluteAnchor element. Mainly for chartsheets. */
        _drawing_write_absolute_anchor(self, 1);
    }

    lxlsx_xml_end_tag(self->file, "xdr:wsDr");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/


/****************************************************************************
 *
 * XLSX drawing/image read support.
 *
 ****************************************************************************/

#include <stdlib.h>
#include <string.h>

#include "xlsx_private.h"
#include "xlsx_util.h"
#include "libxlsx/drawing.h"

/* ------------------------------------------------------------------------- */
/* MIME helpers                                                              */
/* ------------------------------------------------------------------------- */

static const char *guess_mime(const char *path)
{
    const char *dot = path ? strrchr(path, '.') : NULL;
    if (!dot) return "application/octet-stream";
    if (lxlsx_reader_ascii_case_eq(dot, ".png"))  return "image/png";
    if (lxlsx_reader_ascii_case_eq(dot, ".jpg") ||
        lxlsx_reader_ascii_case_eq(dot, ".jpeg")) return "image/jpeg";
    if (lxlsx_reader_ascii_case_eq(dot, ".gif"))  return "image/gif";
    if (lxlsx_reader_ascii_case_eq(dot, ".bmp"))  return "image/bmp";
    if (lxlsx_reader_ascii_case_eq(dot, ".tif") ||
        lxlsx_reader_ascii_case_eq(dot, ".tiff")) return "image/tiff";
    if (lxlsx_reader_ascii_case_eq(dot, ".svg"))  return "image/svg+xml";
    if (lxlsx_reader_ascii_case_eq(dot, ".webp")) return "image/webp";
    return "application/octet-stream";
}

/* ------------------------------------------------------------------------- */
/* Drawing.xml parser                                                        */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxlsx_reader_workbook    *wb;
    char            *drawing_base;
    const lxlsx_reader_rel_map *rids;
    lxlsx_reader_image_cb     cb;
    void            *userdata;
    int              stop;

    int    in_from, in_to, in_col, in_row;
    size_t from_row, from_col, to_row, to_col;
    int    have_from, have_to, have_rid;
    char   current_rid[32];
    char   text_buf[64];
    size_t text_len;
} drawing_ctx;

static void anchor_reset(drawing_ctx *c)
{
    c->in_from = c->in_to = c->in_col = c->in_row = 0;
    c->have_from = c->have_to = c->have_rid = 0;
    c->from_row = c->from_col = c->to_row = c->to_col = 0;
    c->current_rid[0] = 0;
    c->text_len = 0;
}

static void drawing_start(void *ud, const char *name, const char **attrs)
{
    drawing_ctx *c = (drawing_ctx *)ud;
    if (c->stop) return;

    if (lxlsx_reader_xml_name_eq(name, "twoCellAnchor") ||
        lxlsx_reader_xml_name_eq(name, "oneCellAnchor")) {
        anchor_reset(c);
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "from")) { c->in_from = 1; return; }
    if (lxlsx_reader_xml_name_eq(name, "to"))   { c->in_to   = 1; return; }
    if ((c->in_from || c->in_to) && lxlsx_reader_xml_name_eq(name, "col")) {
        c->in_col = 1; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if ((c->in_from || c->in_to) && lxlsx_reader_xml_name_eq(name, "row")) {
        c->in_row = 1; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "blip") && attrs) {
        const char **a;
        for (a = attrs; *a; a += 2) {
            if (lxlsx_reader_xml_name_eq(*a, "embed")) {
                strncpy(c->current_rid, *(a + 1), sizeof(c->current_rid) - 1);
                c->current_rid[sizeof(c->current_rid) - 1] = 0;
                c->have_rid = 1;
                break;
            }
        }
    }
}

static void drawing_text(void *ud, const char *text, int len)
{
    drawing_ctx *c = (drawing_ctx *)ud;
    if ((!c->in_col && !c->in_row) || len <= 0) return;
    if (c->text_len + (size_t)len < sizeof(c->text_buf) - 1) {
        memcpy(c->text_buf + c->text_len, text, (size_t)len);
        c->text_len += (size_t)len;
        c->text_buf[c->text_len] = 0;
    }
}

static void drawing_end(void *ud, const char *name)
{
    drawing_ctx *c = (drawing_ctx *)ud;

    if (c->in_col && lxlsx_reader_xml_name_eq(name, "col")) {
        size_t v = (size_t)strtoul(c->text_buf, NULL, 10);
        if (c->in_from)    { c->from_col = v; c->have_from = 1; }
        else if (c->in_to) { c->to_col   = v; c->have_to   = 1; }
        c->in_col = 0; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (c->in_row && lxlsx_reader_xml_name_eq(name, "row")) {
        size_t v = (size_t)strtoul(c->text_buf, NULL, 10);
        if (c->in_from)    c->from_row = v;
        else if (c->in_to) c->to_row   = v;
        c->in_row = 0; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "from")) { c->in_from = 0; return; }
    if (lxlsx_reader_xml_name_eq(name, "to"))   { c->in_to   = 0; return; }

    if ((lxlsx_reader_xml_name_eq(name, "twoCellAnchor") ||
         lxlsx_reader_xml_name_eq(name, "oneCellAnchor"))
        && c->have_from && c->have_rid) {
        const char *target =
            lxlsx_reader_rel_map_target_for_id(c->rids, c->current_rid);
        if (target) {
            char *full = lxlsx_reader_zip_join_path(c->drawing_base, target);
            if (full) {
                void  *data;
                size_t data_len;
                if (lxlsx_reader_zip_read_entry_all(c->wb->zip, full,
                                                    &data, &data_len) == 0) {
                    lxlsx_reader_image img;
                    const char *slash;
                    memset(&img, 0, sizeof(img));
                    img.from_row  = c->from_row;
                    img.from_col  = c->from_col;
                    img.to_row    = c->have_to ? c->to_row : c->from_row;
                    img.to_col    = c->have_to ? c->to_col : c->from_col;
                    img.mime_type = guess_mime(full);
                    img.data      = data;
                    img.data_len  = data_len;
                    slash         = strrchr(full, '/');
                    img.name      = slash ? slash + 1 : full;

                    if (c->cb(&img, c->userdata) != 0) c->stop = 1;
                    free(data);
                }
                free(full);
            }
        }
        anchor_reset(c);
    }
}

/* ------------------------------------------------------------------------- */
/* Public entry                                                              */
/* ------------------------------------------------------------------------- */

lxlsx_reader_error lxlsx_reader_worksheet_iterate_images(lxlsx_reader_worksheet *ws, lxlsx_reader_image_cb cb, void *ud)
{
    char          *drawing_path = NULL;
    char          *drawing_base = NULL;
    const char    *drawing_target;
    lxlsx_reader_rel_map ws_rels;
    lxlsx_reader_rel_map drawing_rels;
    lxlsx_reader_error rc;

    if (!ws || !ws->wb || !cb) return LXLSX_READER_ERROR_NULL_PARAMETER;
    if (!ws->target_path)   return LXLSX_READER_NO_ERROR;

    rc = lxlsx_reader_load_rels(ws->wb->zip, ws->target_path, &ws_rels, 1);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    drawing_target =
        lxlsx_reader_rel_map_first_target_by_type_suffix(&ws_rels, "/drawing");
    if (!drawing_target) {
        lxlsx_reader_rel_map_free(&ws_rels);
        return LXLSX_READER_NO_ERROR;
    }

    drawing_path = lxlsx_reader_zip_resolve_path(ws->target_path, drawing_target);
    lxlsx_reader_rel_map_free(&ws_rels);
    if (!drawing_path) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    rc = lxlsx_reader_load_rels(ws->wb->zip, drawing_path, &drawing_rels, 1);
    if (rc != LXLSX_READER_NO_ERROR) {
        free(drawing_path);
        return rc;
    }

    /* 4) Parse drawing.xml; emit images via callback.
     *
     * The drawing.xml must be slurped fully into memory before parsing rather
     * than streamed from the zip. minizip's unzFile permits only one open
     * entry per handle, and the drawing_end callback opens each media entry
     * on that same handle — which clobbers the drawing.xml
     * read cursor. A streamed parse therefore dies at the first chunk refill
     * (drawing.xml > LXLSX_READER_PARSE_BUFFER_SIZE), silently truncating the image
     * list to whatever fit in the first buffer. Reading drawing.xml into a
     * private buffer first decouples the two. */
    drawing_base = lxlsx_reader_zip_base_path(drawing_path);
    if (!drawing_base) {
        lxlsx_reader_rel_map_free(&drawing_rels);
        free(drawing_path);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }
    {
        drawing_ctx  dc;
        void        *xml_data = NULL;
        size_t       xml_len  = 0;

        memset(&dc, 0, sizeof(dc));
        dc.wb           = ws->wb;
        dc.drawing_base = drawing_base;
        dc.rids         = &drawing_rels;
        dc.cb           = cb;
        dc.userdata     = ud;

        if (lxlsx_reader_zip_read_entry_all(ws->wb->zip, drawing_path,
                                            &xml_data, &xml_len) == 0) {
            lxlsx_reader_xml_pump *p = lxlsx_reader_xml_pump_create_buffer((const char *)xml_data, xml_len);
            if (p) {
                lxlsx_reader_xml_pump_set_handlers(p, drawing_start, drawing_end, drawing_text, &dc);
                lxlsx_reader_xml_pump_run(p);
                lxlsx_reader_xml_pump_destroy(p);
            }
            free(xml_data);
        }
    }

    free(drawing_path);
    free(drawing_base);
    lxlsx_reader_rel_map_free(&drawing_rels);
    return LXLSX_READER_NO_ERROR;
}
