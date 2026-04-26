#include <ctype.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"
#include "xlsxreader/drawing.h"

/* ------------------------------------------------------------------------- */
/* Path utilities (small, self-contained — duplicates workbook.c on purpose) */
/* ------------------------------------------------------------------------- */

static char *xstrdup(const char *s) { return s ? strdup(s) : NULL; }

static int ci_eq(const char *a, const char *b)
{
    while (*a && *b) {
        char x = tolower((unsigned char)*a++);
        char y = tolower((unsigned char)*b++);
        if (x != y) return 0;
    }
    return *a == 0 && *b == 0;
}

static char *base_dir(const char *path)
{
    const char *slash = path ? strrchr(path, '/') : NULL;
    char *out;
    size_t n;
    if (!slash) return xstrdup("");
    n = (size_t)(slash - path) + 1;
    out = (char *)malloc(n + 1);
    if (!out) return NULL;
    memcpy(out, path, n);
    out[n] = 0;
    return out;
}

static char *rels_path_for(const char *path)
{
    /* "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels" */
    const char *slash;
    size_t      prefix_len, flen;
    const char *file;
    char       *out;

    if (!path) return NULL;
    slash      = strrchr(path, '/');
    prefix_len = slash ? (size_t)(slash - path) + 1 : 0;
    file       = slash ? slash + 1 : path;
    flen       = strlen(file);

    out = (char *)malloc(prefix_len + 6 + flen + 5 + 1);
    if (!out) return NULL;
    memcpy(out,                           path,    prefix_len);
    memcpy(out + prefix_len,              "_rels/", 6);
    memcpy(out + prefix_len + 6,          file,    flen);
    memcpy(out + prefix_len + 6 + flen,   ".rels", 5);
    out[prefix_len + 6 + flen + 5] = 0;
    return out;
}

/* Resolve a Target uri like "../media/image1.png" relative to a base dir
 * (which always ends in '/' if non-empty). Caller frees. */
static char *resolve_relative(const char *base, const char *target)
{
    char  *bcopy;
    size_t blen, tlen;
    char  *out;

    if (!target) return NULL;
    if (target[0] == '/') return xstrdup(target + 1);

    bcopy = xstrdup(base ? base : "");
    if (!bcopy) return NULL;

    while (target[0] == '.' && target[1] == '.' && target[2] == '/') {
        size_t len = strlen(bcopy);
        if (len > 0 && bcopy[len - 1] == '/') len--;
        while (len > 0 && bcopy[len - 1] != '/') len--;
        bcopy[len] = 0;
        target += 3;
    }
    blen = strlen(bcopy);
    tlen = strlen(target);
    out  = (char *)malloc(blen + tlen + 1);
    if (!out) { free(bcopy); return NULL; }
    if (blen) memcpy(out, bcopy, blen);
    memcpy(out + blen, target, tlen);
    out[blen + tlen] = 0;
    free(bcopy);
    return out;
}

static const char *guess_mime(const char *path)
{
    const char *dot = path ? strrchr(path, '.') : NULL;
    if (!dot) return "application/octet-stream";
    if (ci_eq(dot, ".png"))  return "image/png";
    if (ci_eq(dot, ".jpg")  || ci_eq(dot, ".jpeg")) return "image/jpeg";
    if (ci_eq(dot, ".gif"))  return "image/gif";
    if (ci_eq(dot, ".bmp"))  return "image/bmp";
    if (ci_eq(dot, ".tif")  || ci_eq(dot, ".tiff")) return "image/tiff";
    if (ci_eq(dot, ".svg"))  return "image/svg+xml";
    if (ci_eq(dot, ".webp")) return "image/webp";
    return "application/octet-stream";
}

/* ------------------------------------------------------------------------- */
/* Read a full zip entry into a fresh buffer                                 */
/* ------------------------------------------------------------------------- */

static int read_entry_full(lxr_zip *zip, const char *path,
                           void **out_data, size_t *out_len)
{
    lxr_zip_file *zf;
    char         *buf;
    size_t        cap = 8192, len = 0;

    *out_data = NULL;
    *out_len  = 0;
    zf = lxr_zip_open_entry(zip, path);
    if (!zf) return -1;

    buf = (char *)malloc(cap);
    if (!buf) { lxr_zip_close_entry(zf); return -1; }

    for (;;) {
        ssize_t n;
        if (cap - len < 4096) {
            char *nb;
            cap *= 2;
            nb = (char *)realloc(buf, cap);
            if (!nb) { free(buf); lxr_zip_close_entry(zf); return -1; }
            buf = nb;
        }
        n = lxr_zip_read(zf, buf + len, cap - len);
        if (n < 0) { free(buf); lxr_zip_close_entry(zf); return -1; }
        if (n == 0) break;
        len += (size_t)n;
    }
    lxr_zip_close_entry(zf);

    *out_data = buf;
    *out_len  = len;
    return 0;
}

/* ------------------------------------------------------------------------- */
/* Worksheet rels — find drawing target                                      */
/* ------------------------------------------------------------------------- */

typedef struct {
    char *drawing_target;
} ws_rels_state;

static void ws_rels_start(void *ud, const char *name, const char **attrs)
{
    ws_rels_state *s = (ws_rels_state *)ud;
    const char    *type, *target, *p;

    if (!lxr_xml_name_eq(name, "Relationship")) return;
    type   = lxr_xml_attr(attrs, "Type");
    target = lxr_xml_attr(attrs, "Target");
    if (!type || !target) return;

    p = strstr(type, "/relationships/drawing");
    if (p && strcmp(p, "/relationships/drawing") == 0) {
        free(s->drawing_target);
        s->drawing_target = xstrdup(target);
    }
}

/* ------------------------------------------------------------------------- */
/* Drawing rels — rId -> media path                                          */
/* ------------------------------------------------------------------------- */

typedef struct {
    char  rid[32];
    char *target;
} rid_pair;

typedef struct {
    rid_pair *items;
    size_t    count;
    size_t    cap;
} rid_table;

static void rid_table_add(rid_table *t, const char *rid, const char *target)
{
    if (t->count >= t->cap) {
        size_t cap = t->cap ? t->cap * 2 : 8;
        rid_pair *nb = (rid_pair *)realloc(t->items, cap * sizeof(*nb));
        if (!nb) return;
        t->items = nb;
        t->cap   = cap;
    }
    strncpy(t->items[t->count].rid, rid, sizeof(t->items[t->count].rid) - 1);
    t->items[t->count].rid[sizeof(t->items[t->count].rid) - 1] = 0;
    t->items[t->count].target = xstrdup(target);
    t->count++;
}

static const char *rid_table_get(const rid_table *t, const char *rid)
{
    size_t i;
    for (i = 0; i < t->count; i++) {
        if (strcmp(t->items[i].rid, rid) == 0) return t->items[i].target;
    }
    return NULL;
}

static void rid_table_free(rid_table *t)
{
    size_t i;
    for (i = 0; i < t->count; i++) free(t->items[i].target);
    free(t->items);
    t->items = NULL;
    t->count = t->cap = 0;
}

static void drawing_rels_start(void *ud, const char *name, const char **attrs)
{
    rid_table  *t = (rid_table *)ud;
    const char *id, *target;

    if (!lxr_xml_name_eq(name, "Relationship")) return;
    id     = lxr_xml_attr(attrs, "Id");
    target = lxr_xml_attr(attrs, "Target");
    if (id && target) rid_table_add(t, id, target);
}

/* ------------------------------------------------------------------------- */
/* Drawing.xml parser                                                        */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxr_workbook    *wb;
    char            *drawing_base;
    rid_table       *rids;
    lxr_image_cb     cb;
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

    if (lxr_xml_name_eq(name, "twoCellAnchor") ||
        lxr_xml_name_eq(name, "oneCellAnchor")) {
        anchor_reset(c);
        return;
    }
    if (lxr_xml_name_eq(name, "from")) { c->in_from = 1; return; }
    if (lxr_xml_name_eq(name, "to"))   { c->in_to   = 1; return; }
    if ((c->in_from || c->in_to) && lxr_xml_name_eq(name, "col")) {
        c->in_col = 1; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if ((c->in_from || c->in_to) && lxr_xml_name_eq(name, "row")) {
        c->in_row = 1; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (lxr_xml_name_eq(name, "blip") && attrs) {
        const char **a;
        for (a = attrs; *a; a += 2) {
            if (lxr_xml_name_eq(*a, "embed")) {
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

    if (c->in_col && lxr_xml_name_eq(name, "col")) {
        size_t v = (size_t)strtoul(c->text_buf, NULL, 10);
        if (c->in_from)    { c->from_col = v; c->have_from = 1; }
        else if (c->in_to) { c->to_col   = v; c->have_to   = 1; }
        c->in_col = 0; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (c->in_row && lxr_xml_name_eq(name, "row")) {
        size_t v = (size_t)strtoul(c->text_buf, NULL, 10);
        if (c->in_from)    c->from_row = v;
        else if (c->in_to) c->to_row   = v;
        c->in_row = 0; c->text_len = 0; c->text_buf[0] = 0;
        return;
    }
    if (lxr_xml_name_eq(name, "from")) { c->in_from = 0; return; }
    if (lxr_xml_name_eq(name, "to"))   { c->in_to   = 0; return; }

    if ((lxr_xml_name_eq(name, "twoCellAnchor") ||
         lxr_xml_name_eq(name, "oneCellAnchor"))
        && c->have_from && c->have_rid) {
        const char *target = rid_table_get(c->rids, c->current_rid);
        if (target) {
            char *full = resolve_relative(c->drawing_base, target);
            if (full) {
                void  *data;
                size_t data_len;
                if (read_entry_full(c->wb->zip, full, &data, &data_len) == 0) {
                    lxr_image img;
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

lxr_error lxr_worksheet_iterate_images(lxr_worksheet *ws, lxr_image_cb cb, void *ud)
{
    char          *ws_rels = NULL;
    char          *drawing_path = NULL;
    char          *drawing_rels = NULL;
    char          *drawing_base = NULL;
    char          *ws_base      = NULL;
    ws_rels_state  wrs          = { NULL };
    rid_table      rids         = { NULL, 0, 0 };

    if (!ws || !cb)         return LXR_ERROR_NULL_PARAMETER;
    if (!ws->target_path)   return LXR_NO_ERROR;

    /* 1) Find drawing target via worksheet rels (if any) */
    ws_rels = rels_path_for(ws->target_path);
    if (!ws_rels) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    if (!lxr_zip_entry_exists(ws->wb->zip, ws_rels)) {
        free(ws_rels);
        return LXR_NO_ERROR;
    }
    {
        lxr_zip_file *zf = lxr_zip_open_entry(ws->wb->zip, ws_rels);
        free(ws_rels);
        if (!zf) return LXR_NO_ERROR;
        {
            lxr_xml_pump *p = lxr_xml_pump_create_zip_file(zf);
            if (p) {
                lxr_xml_pump_set_handlers(p, ws_rels_start, NULL, NULL, &wrs);
                lxr_xml_pump_run(p);
                lxr_xml_pump_destroy(p);
            }
        }
        lxr_zip_close_entry(zf);
    }

    if (!wrs.drawing_target) return LXR_NO_ERROR;

    /* 2) Resolve drawing path relative to worksheet base */
    ws_base      = base_dir(ws->target_path);
    drawing_path = resolve_relative(ws_base, wrs.drawing_target);
    free(wrs.drawing_target);
    free(ws_base);
    if (!drawing_path) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    /* 3) Drawing rels: rId -> image path */
    drawing_rels = rels_path_for(drawing_path);
    if (drawing_rels && lxr_zip_entry_exists(ws->wb->zip, drawing_rels)) {
        lxr_zip_file *zf = lxr_zip_open_entry(ws->wb->zip, drawing_rels);
        if (zf) {
            lxr_xml_pump *p = lxr_xml_pump_create_zip_file(zf);
            if (p) {
                lxr_xml_pump_set_handlers(p, drawing_rels_start, NULL, NULL, &rids);
                lxr_xml_pump_run(p);
                lxr_xml_pump_destroy(p);
            }
            lxr_zip_close_entry(zf);
        }
    }
    free(drawing_rels);

    /* 4) Parse drawing.xml; emit images via callback */
    drawing_base = base_dir(drawing_path);
    {
        drawing_ctx   dc;
        lxr_zip_file *zf;

        memset(&dc, 0, sizeof(dc));
        dc.wb           = ws->wb;
        dc.drawing_base = drawing_base;
        dc.rids         = &rids;
        dc.cb           = cb;
        dc.userdata     = ud;

        zf = lxr_zip_open_entry(ws->wb->zip, drawing_path);
        if (zf) {
            lxr_xml_pump *p = lxr_xml_pump_create_zip_file(zf);
            if (p) {
                lxr_xml_pump_set_handlers(p, drawing_start, drawing_end, drawing_text, &dc);
                lxr_xml_pump_run(p);
                lxr_xml_pump_destroy(p);
            }
            lxr_zip_close_entry(zf);
        }
    }

    free(drawing_path);
    free(drawing_base);
    rid_table_free(&rids);
    return LXR_NO_ERROR;
}
