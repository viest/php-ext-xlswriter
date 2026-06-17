#include <stdlib.h>
#include <string.h>

#include "lxlsx_reader_internal.h"
#include "lxlsx_reader_util.h"
#include "lxlsx/reader/drawing.h"

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
