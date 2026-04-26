/*
 * Comments reader.
 *
 * Iterates legacy comments (xl/comments*.xml + xl/drawings/vmlDrawing*.vml
 * for visibility) and threaded comments (xl/threadedComments/threadedComment*.xml)
 * for a given worksheet. The sheet's _rels file is the entry point — we
 * resolve the relationship targets there.
 */

#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"

/* --- tiny string-list helper -------------------------------------------- */

typedef struct {
    char  *id;
    char  *target;
} crel_entry;

typedef struct {
    crel_entry *items;
    size_t      count;
    size_t      cap;
} crel_map;

static void crel_init(crel_map *m) { m->items = NULL; m->count = 0; m->cap = 0; }
static void crel_free(crel_map *m) {
    size_t i;
    if (!m) return;
    for (i = 0; i < m->count; i++) { free(m->items[i].id); free(m->items[i].target); }
    free(m->items);
    m->items = NULL; m->count = m->cap = 0;
}
static void crel_push(crel_map *m, const char *id, const char *target) {
    if (m->count >= m->cap) {
        size_t nc = m->cap ? m->cap * 2 : 4;
        crel_entry *nb = (crel_entry *)realloc(m->items, nc * sizeof(*nb));
        if (!nb) return;
        m->items = nb; m->cap = nc;
    }
    m->items[m->count].id     = id     ? strdup(id)     : NULL;
    m->items[m->count].target = target ? strdup(target) : NULL;
    m->count++;
}

/* --- comments XML parse ------------------------------------------------- */

typedef struct {
    char *guid;        /* author GUID for threaded; placeholder index for legacy */
    char *display;
} c_author;

typedef struct {
    size_t row;
    size_t col;
    char  *text;        /* concatenated */
    char  *author;
    char  *guid;        /* for threaded */
    char  *parent_guid;
    int    visible;
    int    threaded;
} c_record;

typedef struct {
    /* Authors table (legacy & threaded share via parallel arrays). */
    c_author *authors;
    size_t    authors_count;
    size_t    authors_cap;
    int       legacy_author_idx;    /* used when collecting authorId-indexed legacy authors */

    /* Records pool (both legacy and threaded mixed in). */
    c_record *records;
    size_t    records_count;
    size_t    records_cap;

    /* SAX state. */
    int   in_authors;
    int   in_author;
    int   in_commentList;
    int   in_comment;
    int   in_text;          /* text/<r>/<t> chain */
    int   in_run;
    int   in_run_t;
    char *cur_text;
    size_t cur_text_len;
    size_t cur_text_cap;
    char *author_text;
    size_t author_text_len;
    size_t author_text_cap;
    /* threaded specifics */
    int   in_threaded_root;
    int   in_threaded_text;
    char *th_pending_id;
    char *th_pending_parent;
    char *th_pending_ref;
    char *th_pending_author_id;
} c_ctx;

/* helper: append text to a heap buffer. */
static void buf_append(char **buf, size_t *len, size_t *cap, const char *s, size_t n)
{
    size_t need = *len + n + 1;
    if (need > *cap) {
        size_t nc = *cap ? *cap : 64;
        char *nb;
        while (nc < need) nc *= 2;
        nb = (char *)realloc(*buf, nc);
        if (!nb) return;
        *buf = nb;
        *cap = nc;
    }
    memcpy(*buf + *len, s, n);
    *len += n;
    (*buf)[*len] = 0;
}

static void buf_reset(char **buf, size_t *len)
{
    *len = 0;
    if (*buf) (*buf)[0] = 0;
}

/* Parse a cell ref like "A1" into 1-based (row, col). */
static void parse_a1(const char *ref, size_t *row, size_t *col)
{
    size_t r = 0, c = 0;
    *row = *col = 0;
    if (!ref) return;
    while (*ref == '$') ref++;
    while ((*ref >= 'A' && *ref <= 'Z') || (*ref >= 'a' && *ref <= 'z')) {
        char x = *ref;
        if (x >= 'a' && x <= 'z') x = (char)(x - 32);
        c = c * 26 + (size_t)(x - 'A' + 1);
        ref++;
    }
    while (*ref == '$') ref++;
    while (*ref >= '0' && *ref <= '9') {
        r = r * 10 + (size_t)(*ref - '0');
        ref++;
    }
    *row = r;
    *col = c;
}

/* --- legacy comments SAX (xl/comments*.xml) ----------------------------- */

static c_record *records_push(c_ctx *c)
{
    if (c->records_count >= c->records_cap) {
        size_t nc = c->records_cap ? c->records_cap * 2 : 8;
        c_record *nb = (c_record *)realloc(c->records, nc * sizeof(*nb));
        if (!nb) return NULL;
        c->records = nb; c->records_cap = nc;
    }
    {
        c_record *r = &c->records[c->records_count++];
        memset(r, 0, sizeof(*r));
        return r;
    }
}

static void on_legacy_start(void *ud, const char *name, const char **attrs)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxr_xml_name_eq(name, "authors")) { c->in_authors = 1; c->legacy_author_idx = 0; return; }
    if (lxr_xml_name_eq(name, "commentList")) { c->in_commentList = 1; return; }
    if (c->in_authors && lxr_xml_name_eq(name, "author")) {
        c->in_author = 1;
        if (c->authors_count >= c->authors_cap) {
            size_t nc = c->authors_cap ? c->authors_cap * 2 : 4;
            c_author *nb = (c_author *)realloc(c->authors, nc * sizeof(*nb));
            if (nb) { c->authors = nb; c->authors_cap = nc; }
        }
        if (c->authors_count < c->authors_cap) {
            char buf[24];
            snprintf(buf, sizeof(buf), "%d", c->legacy_author_idx);
            c->authors[c->authors_count].guid    = strdup(buf);
            c->authors[c->authors_count].display = NULL;
            c->authors_count++;
            c->legacy_author_idx++;
        }
        buf_reset(&c->author_text, &c->author_text_len);
        return;
    }
    if (c->in_commentList && lxr_xml_name_eq(name, "comment")) {
        const char *ref      = lxr_xml_attr(attrs, "ref");
        const char *authorId = lxr_xml_attr(attrs, "authorId");
        c_record   *r        = records_push(c);
        if (r) {
            parse_a1(ref, &r->row, &r->col);
            if (authorId) {
                int idx = (int)strtol(authorId, NULL, 10);
                if (idx >= 0 && (size_t)idx < c->authors_count
                              && c->authors[idx].display) {
                    r->author = strdup(c->authors[idx].display);
                }
            }
            c->in_comment = 1;
            buf_reset(&c->cur_text, &c->cur_text_len);
        }
        return;
    }
    if (c->in_comment) {
        if (lxr_xml_name_eq(name, "text"))    c->in_text = 1;
        else if (c->in_text && lxr_xml_name_eq(name, "r"))   c->in_run = 1;
        else if (c->in_text && lxr_xml_name_eq(name, "t"))   c->in_run_t = 1;
        else if (c->in_run && lxr_xml_name_eq(name, "t"))    c->in_run_t = 1;
    }
}

static void on_legacy_end(void *ud, const char *name)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxr_xml_name_eq(name, "authors"))     c->in_authors = 0;
    if (lxr_xml_name_eq(name, "commentList")) c->in_commentList = 0;
    if (c->in_author && lxr_xml_name_eq(name, "author")) {
        if (c->authors_count > 0) {
            free(c->authors[c->authors_count - 1].display);
            c->authors[c->authors_count - 1].display =
                c->author_text_len > 0 ? strdup(c->author_text) : NULL;
        }
        c->in_author = 0;
    }
    if (c->in_run_t && lxr_xml_name_eq(name, "t")) c->in_run_t = 0;
    if (c->in_run   && lxr_xml_name_eq(name, "r")) c->in_run = 0;
    if (c->in_text  && lxr_xml_name_eq(name, "text")) c->in_text = 0;
    if (c->in_comment && lxr_xml_name_eq(name, "comment")) {
        if (c->records_count > 0 && c->cur_text_len > 0) {
            c_record *r = &c->records[c->records_count - 1];
            r->text = strdup(c->cur_text);
        }
        c->in_comment = 0;
    }
}

static void on_legacy_text(void *ud, const char *text, int len)
{
    c_ctx *c = (c_ctx *)ud;
    if (len <= 0) return;
    if (c->in_author && !c->in_run_t)
        buf_append(&c->author_text, &c->author_text_len, &c->author_text_cap,
                   text, (size_t)len);
    if (c->in_comment && c->in_run_t)
        buf_append(&c->cur_text, &c->cur_text_len, &c->cur_text_cap, text, (size_t)len);
}

/* --- threadedComments parse --------------------------------------------- */

static void on_thread_start(void *ud, const char *name, const char **attrs)
{
    c_ctx *c = (c_ctx *)ud;
    /* Microsoft's threadedComments root is conventionally <ThreadedComments>
     * (capital T), but accept the lowercase form too. */
    if (lxr_xml_name_eq(name, "ThreadedComments") ||
        lxr_xml_name_eq(name, "threadedComments")) { c->in_threaded_root = 1; return; }
    if (!c->in_threaded_root) return;
    if (lxr_xml_name_eq(name, "threadedComment")) {
        const char *ref      = lxr_xml_attr(attrs, "ref");
        const char *id       = lxr_xml_attr(attrs, "id");
        const char *parent   = lxr_xml_attr(attrs, "parentId");
        const char *authorId = lxr_xml_attr(attrs, "personId");
        free(c->th_pending_id);        c->th_pending_id        = id       ? strdup(id) : NULL;
        free(c->th_pending_parent);    c->th_pending_parent    = parent   ? strdup(parent) : NULL;
        free(c->th_pending_ref);       c->th_pending_ref       = ref      ? strdup(ref) : NULL;
        free(c->th_pending_author_id); c->th_pending_author_id = authorId ? strdup(authorId) : NULL;
        buf_reset(&c->cur_text, &c->cur_text_len);
        c->in_threaded_text = 0;
        return;
    }
    if (lxr_xml_name_eq(name, "text")) c->in_threaded_text = 1;
}

static void on_thread_end(void *ud, const char *name)
{
    c_ctx *c = (c_ctx *)ud;
    if (lxr_xml_name_eq(name, "text")) c->in_threaded_text = 0;
    if (lxr_xml_name_eq(name, "ThreadedComments") ||
        lxr_xml_name_eq(name, "threadedComments")) c->in_threaded_root = 0;
    if (lxr_xml_name_eq(name, "threadedComment")) {
        c_record *r = records_push(c);
        if (r) {
            r->threaded = 1;
            r->visible  = 0;
            parse_a1(c->th_pending_ref, &r->row, &r->col);
            r->guid          = c->th_pending_id;        c->th_pending_id = NULL;
            r->parent_guid   = c->th_pending_parent;    c->th_pending_parent = NULL;
            r->author        = c->th_pending_author_id; c->th_pending_author_id = NULL;
            r->text          = c->cur_text_len > 0 ? strdup(c->cur_text) : NULL;
        }
    }
}

static void on_thread_text(void *ud, const char *text, int len)
{
    c_ctx *c = (c_ctx *)ud;
    if (c->in_threaded_text && len > 0)
        buf_append(&c->cur_text, &c->cur_text_len, &c->cur_text_cap, text, (size_t)len);
}

/* --- _rels resolution --------------------------------------------------- */

static const char *find_rel_target(crel_map *m, const char *id)
{
    size_t i;
    if (!m || !id) return NULL;
    for (i = 0; i < m->count; i++) {
        if (m->items[i].id && strcmp(m->items[i].id, id) == 0)
            return m->items[i].target;
    }
    return NULL;
}

typedef struct {
    crel_map *m;
    const char *want_type;
    char       *target;
} rel_collect_ctx;

static void rel_collect_on_start(void *ud, const char *name, const char **attrs)
{
    rel_collect_ctx *c = (rel_collect_ctx *)ud;
    const char *id, *type, *target;
    if (!lxr_xml_name_eq(name, "Relationship")) return;
    id     = lxr_xml_attr(attrs, "Id");
    type   = lxr_xml_attr(attrs, "Type");
    target = lxr_xml_attr(attrs, "Target");
    if (!id || !type || !target) return;
    /* Match by suffix — e.g. ".../comments" or ".../vmlDrawing" or
     * ".../threadedComment". The full URI has well-known prefixes. */
    if (c->want_type) {
        size_t n = strlen(c->want_type);
        size_t tn = strlen(type);
        if (tn < n) return;
        if (strcmp(type + tn - n, c->want_type) != 0) return;
    }
    crel_push(c->m, id, target);
    if (!c->target) c->target = strdup(target);
}

/* Open xl/worksheets/_rels/<sheet>.xml.rels (or any file's rels). */
static lxr_error load_rels_filter(lxr_zip *zip, const char *target_path,
                                  const char *want_type_suffix,
                                  crel_map *out, char **first_target)
{
    char *rels_path;
    const char *slash = target_path ? strrchr(target_path, '/') : NULL;
    size_t prefix_len = slash ? (size_t)(slash - target_path) + 1 : 0;
    const char *file_part = slash ? slash + 1 : target_path;
    size_t flen = strlen(file_part);
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    rel_collect_ctx rc = { out, want_type_suffix, NULL };
    lxr_error rc_err;

    rels_path = (char *)malloc(prefix_len + 6 + flen + 5 + 1);
    if (!rels_path) return LXR_ERROR_MEMORY_MALLOC_FAILED;
    memcpy(rels_path, target_path, prefix_len);
    memcpy(rels_path + prefix_len, "_rels/", 6);
    memcpy(rels_path + prefix_len + 6, file_part, flen);
    memcpy(rels_path + prefix_len + 6 + flen, ".rels", 5);
    rels_path[prefix_len + 6 + flen + 5] = 0;

    zf = lxr_zip_open_entry(zip, rels_path);
    free(rels_path);
    if (!zf) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }

    lxr_xml_pump_set_handlers(pump, rel_collect_on_start, NULL, NULL, &rc);
    rc_err = lxr_xml_pump_run(pump);
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);

    if (first_target) *first_target = rc.target;
    else free(rc.target);
    return rc_err;
}

/* Resolve a relative target against a base path: e.g.
 * base="xl/worksheets/sheet1.xml", target="../comments1.xml" -> "xl/comments1.xml" */
static char *normalise_rel_target(const char *base, const char *target)
{
    size_t blen, tlen;
    char  *full;
    const char *slash;
    if (!base || !target) return NULL;
    if (target[0] == '/') {
        full = strdup(target + 1);
        return full;
    }
    slash = strrchr(base, '/');
    blen = slash ? (size_t)(slash - base) + 1 : 0;
    tlen = strlen(target);
    full = (char *)malloc(blen + tlen + 1);
    if (!full) return NULL;
    if (blen) memcpy(full, base, blen);
    memcpy(full + blen, target, tlen);
    full[blen + tlen] = 0;
    /* Resolve "../" segments. */
    {
        char *out = full;
        char *p   = full;
        while (*p) {
            if (p[0] == '.' && p[1] == '.' && p[2] == '/') {
                if (out > full) {
                    out--;
                    while (out > full && *(out - 1) != '/') out--;
                }
                p += 3;
            } else {
                *out++ = *p++;
            }
        }
        *out = 0;
    }
    return full;
}

/* --- vmlDrawing visibility extraction (legacy) -------------------------- */

typedef struct {
    /* row/col -> visibility map. We just collect (row,col,visible). */
    struct { size_t row; size_t col; int visible; } *items;
    size_t count;
    size_t cap;
} vml_visibility;

static void vml_push(vml_visibility *v, size_t row, size_t col, int visible)
{
    if (v->count >= v->cap) {
        size_t nc = v->cap ? v->cap * 2 : 8;
        void *nb = realloc(v->items, nc * sizeof(*v->items));
        if (!nb) return;
        v->items = nb; v->cap = nc;
    }
    v->items[v->count].row     = row;
    v->items[v->count].col     = col;
    v->items[v->count].visible = visible;
    v->count++;
}

static int vml_lookup(const vml_visibility *v, size_t row, size_t col)
{
    size_t i;
    if (!v) return 0;
    for (i = 0; i < v->count; i++) {
        if (v->items[i].row == row && v->items[i].col == col)
            return v->items[i].visible;
    }
    return 0;
}

typedef struct {
    vml_visibility *v;
    int             in_shape;
    int             in_clientdata;
    int             in_anchor;
    int             in_visible;
    int             in_row;
    int             in_column;
    int             pending_visible;
    long            pending_col;       /* 0-based VML uses col, row */
    long            pending_row;
    char           *anchor_text;
    size_t          anchor_text_len;
    size_t          anchor_text_cap;
    char           *small_buf;
    size_t          small_buf_len;
    size_t          small_buf_cap;
} vml_ctx;

/* The VML format (legacy Excel comments) is a subset of legacy MS Office
 * shapes: each <v:shape> with an <x:ClientData ObjectType="Note"> child
 * that contains <x:Visible/>, <x:Row>R</x:Row>, <x:Column>C</x:Column>.
 * Tag prefixes vary; lxr_xml_name_eq strips namespaces. */
static void vml_on_start(void *ud, const char *name, const char **attrs)
{
    vml_ctx *c = (vml_ctx *)ud;
    (void)attrs;
    if (lxr_xml_name_eq(name, "shape")) {
        c->in_shape = 1; c->pending_visible = 0;
        c->pending_row = c->pending_col = -1;
    }
    if (c->in_shape && lxr_xml_name_eq(name, "ClientData")) c->in_clientdata = 1;
    if (c->in_clientdata && lxr_xml_name_eq(name, "Visible")) c->pending_visible = 1;
    if (c->in_clientdata && lxr_xml_name_eq(name, "Row"))     { c->in_row = 1; c->small_buf_len = 0; if (c->small_buf) c->small_buf[0] = 0; }
    if (c->in_clientdata && lxr_xml_name_eq(name, "Column"))  { c->in_column = 1; c->small_buf_len = 0; if (c->small_buf) c->small_buf[0] = 0; }
}

static void vml_on_text(void *ud, const char *text, int len)
{
    vml_ctx *c = (vml_ctx *)ud;
    if (len <= 0) return;
    if (c->in_row || c->in_column) {
        buf_append(&c->small_buf, &c->small_buf_len, &c->small_buf_cap, text, (size_t)len);
    }
}

static void vml_on_end(void *ud, const char *name)
{
    vml_ctx *c = (vml_ctx *)ud;
    if (c->in_row && lxr_xml_name_eq(name, "Row")) {
        c->pending_row = strtol(c->small_buf ? c->small_buf : "-1", NULL, 10);
        c->in_row = 0;
    }
    if (c->in_column && lxr_xml_name_eq(name, "Column")) {
        c->pending_col = strtol(c->small_buf ? c->small_buf : "-1", NULL, 10);
        c->in_column = 0;
    }
    if (c->in_clientdata && lxr_xml_name_eq(name, "ClientData")) {
        if (c->pending_row >= 0 && c->pending_col >= 0) {
            vml_push(c->v, (size_t)(c->pending_row + 1),
                     (size_t)(c->pending_col + 1), c->pending_visible);
        }
        c->in_clientdata = 0;
    }
    if (c->in_shape && lxr_xml_name_eq(name, "shape")) {
        c->in_shape = 0;
    }
}

static lxr_error parse_vml(lxr_zip *zip, const char *path, vml_visibility *v)
{
    lxr_zip_file *zf = lxr_zip_open_entry(zip, path);
    lxr_xml_pump *pump;
    vml_ctx       ctx;
    lxr_error     rc;
    if (!zf) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }
    memset(&ctx, 0, sizeof(ctx));
    ctx.v = v;
    lxr_xml_pump_set_handlers(pump, vml_on_start, vml_on_end, vml_on_text, &ctx);
    rc = lxr_xml_pump_run(pump);
    free(ctx.small_buf);
    free(ctx.anchor_text);
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* --- public entry ------------------------------------------------------- */

static void free_records(c_record *records, size_t n)
{
    size_t i;
    for (i = 0; i < n; i++) {
        free(records[i].text);
        free(records[i].author);
        free(records[i].guid);
        free(records[i].parent_guid);
    }
    free(records);
}

static void free_authors(c_author *authors, size_t n)
{
    size_t i;
    for (i = 0; i < n; i++) {
        free(authors[i].guid);
        free(authors[i].display);
    }
    free(authors);
}

static lxr_error scan_one_xml(lxr_zip *zip, const char *path,
                              c_ctx *ctx, int threaded)
{
    lxr_zip_file *zf = lxr_zip_open_entry(zip, path);
    lxr_xml_pump *pump;
    lxr_error     rc;
    if (!zf) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }
    if (threaded) {
        lxr_xml_pump_set_handlers(pump, on_thread_start, on_thread_end, on_thread_text, ctx);
    } else {
        lxr_xml_pump_set_handlers(pump, on_legacy_start, on_legacy_end, on_legacy_text, ctx);
    }
    rc = lxr_xml_pump_run(pump);
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

lxr_error lxr_worksheet_iterate_comments(lxr_worksheet *ws,
                                         lxr_comment_cb cb, void *userdata)
{
    crel_map        rels;
    char           *comments_target = NULL;
    char           *thread_target   = NULL;
    char           *vml_target      = NULL;
    char           *comments_path   = NULL;
    char           *thread_path     = NULL;
    char           *vml_path        = NULL;
    c_ctx           ctx;
    vml_visibility  vml;
    lxr_error       rc;
    size_t          i;

    if (!ws || !ws->wb || !cb) return LXR_ERROR_NULL_PARAMETER;

    crel_init(&rels);
    /* Collect ALL relationships once. */
    rc = load_rels_filter(ws->wb->zip, ws->target_path, NULL,
                          &rels, NULL);
    if (rc != LXR_NO_ERROR) {
        crel_free(&rels);
        return rc;
    }

    /* Walk the rels map and bucket. */
    for (i = 0; i < rels.count; i++) {
        const char *t = rels.items[i].target;
        if (!t) continue;
        if (strstr(t, "comments") && !strstr(t, "threaded") && !comments_target) {
            comments_target = strdup(t);
        }
        if (strstr(t, "threadedComment") && !thread_target) {
            thread_target = strdup(t);
        }
        if (strstr(t, "vmlDrawing") && !vml_target) {
            vml_target = strdup(t);
        }
    }

    crel_free(&rels);

    if (!comments_target && !thread_target) {
        free(vml_target);
        return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    }

    comments_path = comments_target ? normalise_rel_target(ws->target_path, comments_target) : NULL;
    thread_path   = thread_target   ? normalise_rel_target(ws->target_path, thread_target)   : NULL;
    vml_path      = vml_target      ? normalise_rel_target(ws->target_path, vml_target)      : NULL;
    free(comments_target);
    free(thread_target);
    free(vml_target);

    memset(&ctx, 0, sizeof(ctx));
    memset(&vml, 0, sizeof(vml));

    if (comments_path) (void)scan_one_xml(ws->wb->zip, comments_path, &ctx, 0);
    if (thread_path)   (void)scan_one_xml(ws->wb->zip, thread_path,   &ctx, 1);
    if (vml_path)      (void)parse_vml(ws->wb->zip, vml_path, &vml);

    free(comments_path); free(thread_path); free(vml_path);

    /* Deliver. */
    {
        int stop = 0;
        for (i = 0; i < ctx.records_count && !stop; i++) {
            lxr_comment_info info;
            memset(&info, 0, sizeof(info));
            info.row       = ctx.records[i].row;
            info.col       = ctx.records[i].col;
            info.text      = ctx.records[i].text;
            info.author    = ctx.records[i].author;
            info.threaded  = ctx.records[i].threaded;
            info.parent_id = ctx.records[i].parent_guid;
            info.visible   = ctx.records[i].threaded ? 0 :
                             vml_lookup(&vml, info.row, info.col);
            if (cb(&info, userdata) != 0) stop = 1;
        }
    }

    free_records(ctx.records, ctx.records_count);
    free_authors(ctx.authors, ctx.authors_count);
    free(ctx.cur_text);
    free(ctx.author_text);
    free(ctx.th_pending_id);
    free(ctx.th_pending_parent);
    free(ctx.th_pending_ref);
    free(ctx.th_pending_author_id);
    free(vml.items);
    return LXR_NO_ERROR;
}
