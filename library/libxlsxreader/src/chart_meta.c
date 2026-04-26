/*
 * Chart metadata reader.
 *
 * Shallow parse: chart type / title / series ranges / drawing anchor. We do
 * NOT touch styles, axes details, or layout — the goal is enumeration and
 * migration, not rendering. The flow is:
 *
 *   sheet's _rels  -> drawing relationship (xl/drawings/drawingN.xml)
 *   drawing's _rels -> chart relationships (xl/charts/chartN.xml)
 *
 *   For each <xdr:twoCellAnchor> in the drawing that contains a
 *   <c:chart r:id=...>, we record the anchor and parse the referenced chart
 *   for type/title/series.
 */

#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"

/* Forward decl from comments.c-private helpers. To avoid a public header we
 * duplicate two tiny utilities here. */

static char *cm_strdup_or_null(const char *s) { return s ? strdup(s) : NULL; }

static char *cm_normalise(const char *base, const char *target)
{
    size_t blen, tlen;
    char  *full;
    const char *slash;
    if (!base || !target) return NULL;
    if (target[0] == '/') return strdup(target + 1);
    slash = strrchr(base, '/');
    blen = slash ? (size_t)(slash - base) + 1 : 0;
    tlen = strlen(target);
    full = (char *)malloc(blen + tlen + 1);
    if (!full) return NULL;
    if (blen) memcpy(full, base, blen);
    memcpy(full + blen, target, tlen);
    full[blen + tlen] = 0;
    {
        char *out = full, *p = full;
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

/* Reuse the rel-collector pattern: build (rid -> target) for a given file. */

typedef struct { char *id; char *target; } rel_e;
typedef struct { rel_e *items; size_t count; size_t cap; } rel_m;

static void relm_push(rel_m *m, const char *id, const char *t)
{
    if (m->count >= m->cap) {
        size_t nc = m->cap ? m->cap * 2 : 4;
        rel_e *nb = (rel_e *)realloc(m->items, nc * sizeof(*nb));
        if (!nb) return;
        m->items = nb; m->cap = nc;
    }
    m->items[m->count].id     = id ? strdup(id) : NULL;
    m->items[m->count].target = t  ? strdup(t)  : NULL;
    m->count++;
}

static void relm_free(rel_m *m)
{
    size_t i;
    for (i = 0; i < m->count; i++) { free(m->items[i].id); free(m->items[i].target); }
    free(m->items);
    m->items = NULL; m->count = m->cap = 0;
}

static const char *relm_lookup(const rel_m *m, const char *id)
{
    size_t i;
    for (i = 0; i < m->count; i++)
        if (m->items[i].id && strcmp(m->items[i].id, id) == 0) return m->items[i].target;
    return NULL;
}

static void relm_on_start(void *ud, const char *name, const char **attrs)
{
    rel_m *m = (rel_m *)ud;
    if (!lxr_xml_name_eq(name, "Relationship")) return;
    relm_push(m, lxr_xml_attr(attrs, "Id"), lxr_xml_attr(attrs, "Target"));
}

static lxr_error load_rels(lxr_zip *zip, const char *target_path, rel_m *out)
{
    char *rels_path;
    const char *slash = target_path ? strrchr(target_path, '/') : NULL;
    size_t prefix_len = slash ? (size_t)(slash - target_path) + 1 : 0;
    const char *file_part = slash ? slash + 1 : target_path;
    size_t flen = strlen(file_part);
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    lxr_error rc;

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
    lxr_xml_pump_set_handlers(pump, relm_on_start, NULL, NULL, out);
    rc = lxr_xml_pump_run(pump);
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* --- drawing anchor & chart-rid scan ------------------------------------ */

typedef struct {
    char  *rid;            /* r:id of the chart */
    long   from_col, from_row, to_col, to_row;  /* 0-based */
} drawing_anchor;

typedef struct {
    drawing_anchor *items;
    size_t count, cap;
    /* SAX state */
    int    in_two;          /* twoCellAnchor */
    int    in_one;          /* oneCellAnchor */
    int    in_from;
    int    in_to;
    int    in_col;
    int    in_row;
    int    in_chart_rid;
    long   pending_col, pending_row;
    long   from_col, from_row, to_col, to_row;
    char  *small;
    size_t small_len;
    size_t small_cap;
    drawing_anchor *cur;
} drw_ctx;

static void drw_alloc_anchor(drw_ctx *c)
{
    if (c->count >= c->cap) {
        size_t nc = c->cap ? c->cap * 2 : 4;
        drawing_anchor *nb = (drawing_anchor *)realloc(c->items, nc * sizeof(*nb));
        if (!nb) return;
        c->items = nb; c->cap = nc;
    }
    c->cur = &c->items[c->count++];
    memset(c->cur, 0, sizeof(*c->cur));
    c->from_col = c->from_row = c->to_col = c->to_row = -1;
}

static void drw_buf_append(drw_ctx *c, const char *s, size_t n)
{
    size_t need = c->small_len + n + 1;
    if (need > c->small_cap) {
        size_t nc = c->small_cap ? c->small_cap : 32;
        char *nb;
        while (nc < need) nc *= 2;
        nb = (char *)realloc(c->small, nc);
        if (!nb) return;
        c->small = nb; c->small_cap = nc;
    }
    memcpy(c->small + c->small_len, s, n);
    c->small_len += n;
    c->small[c->small_len] = 0;
}

static void drw_on_start(void *ud, const char *name, const char **attrs)
{
    drw_ctx *c = (drw_ctx *)ud;
    if (lxr_xml_name_eq(name, "twoCellAnchor")) { drw_alloc_anchor(c); c->in_two = 1; return; }
    if (lxr_xml_name_eq(name, "oneCellAnchor")) { drw_alloc_anchor(c); c->in_one = 1; return; }
    if (!c->cur) return;
    if (lxr_xml_name_eq(name, "from")) c->in_from = 1;
    if (lxr_xml_name_eq(name, "to"))   c->in_to   = 1;
    if (c->in_from || c->in_to) {
        if (lxr_xml_name_eq(name, "col")) { c->in_col = 1; c->small_len = 0; if (c->small) c->small[0] = 0; }
        if (lxr_xml_name_eq(name, "row")) { c->in_row = 1; c->small_len = 0; if (c->small) c->small[0] = 0; }
    }
    if (lxr_xml_name_eq(name, "chart")) {
        const char *rid = lxr_xml_attr(attrs, "id");
        if (!rid) {
            const char **a = attrs;
            while (a && *a) {
                if (lxr_xml_name_eq(*a, "id")) { rid = *(a + 1); break; }
                a += 2;
            }
        }
        if (rid && c->cur) c->cur->rid = strdup(rid);
    }
}

static void drw_on_text(void *ud, const char *text, int len)
{
    drw_ctx *c = (drw_ctx *)ud;
    if ((c->in_col || c->in_row) && len > 0) drw_buf_append(c, text, (size_t)len);
}

static void drw_on_end(void *ud, const char *name)
{
    drw_ctx *c = (drw_ctx *)ud;
    if (c->in_col && lxr_xml_name_eq(name, "col")) {
        long v = strtol(c->small ? c->small : "-1", NULL, 10);
        if (c->in_from) c->from_col = v; else if (c->in_to) c->to_col = v;
        c->in_col = 0;
    }
    if (c->in_row && lxr_xml_name_eq(name, "row")) {
        long v = strtol(c->small ? c->small : "-1", NULL, 10);
        if (c->in_from) c->from_row = v; else if (c->in_to) c->to_row = v;
        c->in_row = 0;
    }
    if (c->in_from && lxr_xml_name_eq(name, "from")) c->in_from = 0;
    if (c->in_to   && lxr_xml_name_eq(name, "to"))   c->in_to   = 0;
    if ((c->in_two && lxr_xml_name_eq(name, "twoCellAnchor")) ||
        (c->in_one && lxr_xml_name_eq(name, "oneCellAnchor"))) {
        if (c->cur) {
            c->cur->from_col = c->from_col;
            c->cur->from_row = c->from_row;
            c->cur->to_col   = c->to_col;
            c->cur->to_row   = c->to_row;
            /* If no chart rid was found, drop this anchor. */
            if (!c->cur->rid) c->count--;
        }
        c->cur = NULL;
        c->in_two = c->in_one = 0;
    }
}

static lxr_error scan_drawing(lxr_zip *zip, const char *path, drw_ctx *out)
{
    lxr_zip_file *zf = lxr_zip_open_entry(zip, path);
    lxr_xml_pump *pump;
    lxr_error rc;
    if (!zf) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }
    lxr_xml_pump_set_handlers(pump, drw_on_start, drw_on_end, drw_on_text, out);
    rc = lxr_xml_pump_run(pump);
    free(out->small);
    out->small = NULL; out->small_len = out->small_cap = 0;
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* --- chart XML scan ----------------------------------------------------- */

typedef struct {
    char *name;
    char *categories;
    char *values;
} chart_series_owned;

typedef struct {
    char *type;        /* one of barChart/lineChart/etc. */
    char *title;
    chart_series_owned *series;
    size_t series_count, series_cap;

    /* SAX state */
    int   in_chart;
    int   in_plotarea;
    int   in_title;
    int   in_title_text;
    int   in_ser;
    int   in_tx;        /* <c:tx><c:strRef><c:f>name</c:f></c:strRef></c:tx> */
    int   in_cat;
    int   in_val;
    int   in_f;         /* <c:f>...</c:f> */
    char *txt;
    size_t txt_len, txt_cap;
    int   capture_target;    /* 1=name, 2=categories, 3=values, 4=title */
} chart_ctx;

static void chart_buf_append(chart_ctx *c, const char *s, size_t n)
{
    size_t need = c->txt_len + n + 1;
    if (need > c->txt_cap) {
        size_t nc = c->txt_cap ? c->txt_cap : 32;
        char *nb;
        while (nc < need) nc *= 2;
        nb = (char *)realloc(c->txt, nc);
        if (!nb) return;
        c->txt = nb; c->txt_cap = nc;
    }
    memcpy(c->txt + c->txt_len, s, n);
    c->txt_len += n;
    c->txt[c->txt_len] = 0;
}
static void chart_buf_reset(chart_ctx *c) { c->txt_len = 0; if (c->txt) c->txt[0] = 0; }

static chart_series_owned *chart_push_series(chart_ctx *c)
{
    if (c->series_count >= c->series_cap) {
        size_t nc = c->series_cap ? c->series_cap * 2 : 4;
        chart_series_owned *nb = (chart_series_owned *)realloc(
            c->series, nc * sizeof(*nb));
        if (!nb) return NULL;
        c->series = nb;
        c->series_cap = nc;
    }
    {
        chart_series_owned *s = &c->series[c->series_count++];
        memset(s, 0, sizeof(*s));
        return s;
    }
}

static void chart_on_start(void *ud, const char *name, const char **attrs)
{
    chart_ctx *c = (chart_ctx *)ud;
    (void)attrs;
    if (lxr_xml_name_eq(name, "chart")) c->in_chart = 1;
    if (!c->in_chart) return;
    if (lxr_xml_name_eq(name, "plotArea")) c->in_plotarea = 1;
    /* Detect chart type by element name in the plot area. */
    if (c->in_plotarea && !c->type) {
        const char *types[] = {
            "barChart","lineChart","pieChart","areaChart","scatterChart",
            "radarChart","doughnutChart","stockChart","bubbleChart",
            "ofPieChart","surfaceChart","line3DChart","bar3DChart",
            "pie3DChart","area3DChart","surface3DChart", NULL
        };
        int i;
        for (i = 0; types[i]; i++) {
            if (lxr_xml_name_eq(name, types[i])) {
                c->type = strdup(types[i]);
                break;
            }
        }
    }
    if (lxr_xml_name_eq(name, "title")) { c->in_title = 1; }
    if (c->in_title && lxr_xml_name_eq(name, "t") && !c->title) {
        c->in_title_text = 1; c->capture_target = 4; chart_buf_reset(c);
    }
    if (lxr_xml_name_eq(name, "ser") && c->in_plotarea) {
        chart_push_series(c);
        c->in_ser = 1;
    }
    if (c->in_ser) {
        if (lxr_xml_name_eq(name, "tx"))    { c->in_tx  = 1; }
        if (lxr_xml_name_eq(name, "cat"))   { c->in_cat = 1; }
        if (lxr_xml_name_eq(name, "val"))   { c->in_val = 1; }
        if (lxr_xml_name_eq(name, "f"))     {
            c->in_f = 1; chart_buf_reset(c);
            if      (c->in_tx)  c->capture_target = 1;
            else if (c->in_cat) c->capture_target = 2;
            else if (c->in_val) c->capture_target = 3;
            else                c->capture_target = 0;
        }
    }
}

static void chart_on_text(void *ud, const char *text, int len)
{
    chart_ctx *c = (chart_ctx *)ud;
    if (len <= 0) return;
    if (c->in_f && c->capture_target > 0) chart_buf_append(c, text, (size_t)len);
    else if (c->in_title_text && c->capture_target == 4) chart_buf_append(c, text, (size_t)len);
}

static void chart_on_end(void *ud, const char *name)
{
    chart_ctx *c = (chart_ctx *)ud;
    if (c->in_f && lxr_xml_name_eq(name, "f")) {
        chart_series_owned *s = c->series_count > 0 ? &c->series[c->series_count - 1] : NULL;
        if (s && c->txt_len > 0) {
            switch (c->capture_target) {
                case 1: free(s->name);       s->name       = strdup(c->txt); break;
                case 2: free(s->categories); s->categories = strdup(c->txt); break;
                case 3: free(s->values);     s->values     = strdup(c->txt); break;
                default: break;
            }
        }
        c->in_f = 0;
        c->capture_target = 0;
    }
    if (c->in_title_text && lxr_xml_name_eq(name, "t")) {
        if (!c->title && c->txt_len > 0) c->title = strdup(c->txt);
        c->in_title_text = 0;
        c->capture_target = 0;
    }
    if (lxr_xml_name_eq(name, "tx"))     c->in_tx = 0;
    if (lxr_xml_name_eq(name, "cat"))    c->in_cat = 0;
    if (lxr_xml_name_eq(name, "val"))    c->in_val = 0;
    if (lxr_xml_name_eq(name, "ser"))    c->in_ser = 0;
    if (lxr_xml_name_eq(name, "title"))  c->in_title = 0;
    if (lxr_xml_name_eq(name, "plotArea")) c->in_plotarea = 0;
    if (lxr_xml_name_eq(name, "chart"))  c->in_chart = 0;
}

static lxr_error scan_chart(lxr_zip *zip, const char *path, chart_ctx *out)
{
    lxr_zip_file *zf = lxr_zip_open_entry(zip, path);
    lxr_xml_pump *pump;
    lxr_error rc;
    if (!zf) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;
    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }
    lxr_xml_pump_set_handlers(pump, chart_on_start, chart_on_end, chart_on_text, out);
    rc = lxr_xml_pump_run(pump);
    free(out->txt);
    out->txt = NULL; out->txt_len = out->txt_cap = 0;
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

static void chart_meta_free_owned(chart_ctx *c)
{
    size_t i;
    free(c->type);
    free(c->title);
    for (i = 0; i < c->series_count; i++) {
        free(c->series[i].name);
        free(c->series[i].categories);
        free(c->series[i].values);
    }
    free(c->series);
    free(c->txt);
    memset(c, 0, sizeof(*c));
}

/* --- public entry ------------------------------------------------------- */

lxr_error lxr_worksheet_iterate_charts(lxr_worksheet *ws,
                                       lxr_chart_cb cb, void *userdata)
{
    rel_m sheet_rels;
    char *drawing_target = NULL;
    char *drawing_path   = NULL;
    drw_ctx drw;
    rel_m chart_rels;
    size_t i;
    lxr_error rc;

    if (!ws || !ws->wb || !cb) return LXR_ERROR_NULL_PARAMETER;

    memset(&sheet_rels, 0, sizeof(sheet_rels));
    rc = load_rels(ws->wb->zip, ws->target_path, &sheet_rels);
    if (rc != LXR_NO_ERROR) { relm_free(&sheet_rels); return rc; }

    /* Find the drawing target. */
    for (i = 0; i < sheet_rels.count; i++) {
        if (sheet_rels.items[i].target &&
            strstr(sheet_rels.items[i].target, "drawing")) {
            drawing_target = cm_strdup_or_null(sheet_rels.items[i].target);
            break;
        }
    }
    relm_free(&sheet_rels);

    if (!drawing_target) return LXR_ERROR_ZIP_ENTRY_NOT_FOUND;

    drawing_path = cm_normalise(ws->target_path, drawing_target);
    free(drawing_target);
    if (!drawing_path) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    /* Scan the drawing for anchors + chart rids. */
    memset(&drw, 0, sizeof(drw));
    (void)scan_drawing(ws->wb->zip, drawing_path, &drw);

    /* Resolve drawing's _rels: rid -> chart path. */
    memset(&chart_rels, 0, sizeof(chart_rels));
    (void)load_rels(ws->wb->zip, drawing_path, &chart_rels);

    {
        int stop = 0;
        for (i = 0; i < drw.count && !stop; i++) {
            const char *target = relm_lookup(&chart_rels, drw.items[i].rid);
            char       *chart_path;
            chart_ctx   ch;
            if (!target) continue;
            chart_path = cm_normalise(drawing_path, target);
            if (!chart_path) continue;

            memset(&ch, 0, sizeof(ch));
            (void)scan_chart(ws->wb->zip, chart_path, &ch);
            free(chart_path);

            {
                lxr_chart_meta info;
                lxr_chart_series_info *info_series = NULL;
                size_t j;

                memset(&info, 0, sizeof(info));
                info.type  = ch.type;
                info.title = ch.title;
                info.from_col = drw.items[i].from_col >= 0
                    ? (size_t)(drw.items[i].from_col + 1) : 0;
                info.from_row = drw.items[i].from_row >= 0
                    ? (size_t)(drw.items[i].from_row + 1) : 0;
                info.to_col   = drw.items[i].to_col >= 0
                    ? (size_t)(drw.items[i].to_col + 1) : 0;
                info.to_row   = drw.items[i].to_row >= 0
                    ? (size_t)(drw.items[i].to_row + 1) : 0;

                if (ch.series_count > 0) {
                    info_series = (lxr_chart_series_info *)calloc(
                        ch.series_count, sizeof(*info_series));
                    if (info_series) {
                        for (j = 0; j < ch.series_count; j++) {
                            info_series[j].name       = ch.series[j].name;
                            info_series[j].categories = ch.series[j].categories;
                            info_series[j].values     = ch.series[j].values;
                        }
                    }
                }
                info.series       = info_series;
                info.series_count = ch.series_count;

                if (cb(&info, userdata) != 0) stop = 1;
                free(info_series);
            }
            chart_meta_free_owned(&ch);
        }
    }

    {
        size_t k;
        for (k = 0; k < drw.count; k++) free(drw.items[k].rid);
        free(drw.items);
    }
    relm_free(&chart_rels);
    free(drawing_path);
    return LXR_NO_ERROR;
}
