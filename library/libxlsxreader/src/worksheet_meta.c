/*
 * Phase-1 sheet metadata parser.
 *
 * libxlsxreader's streaming pump in worksheet.c is single-purpose: it walks
 * <sheetData> and emits cells. But several worksheet-level elements
 * (<mergeCells>, <hyperlinks>, <sheetProtection>, <cols>, <sheetFormatPr>,
 * plus row attrs on <row> elements themselves) live as siblings of
 * <sheetData> and may appear *after* it in the XML stream. They're also too
 * small to justify on-demand re-parsing.
 *
 * This file implements an eager scan at sheet open: open a fresh entry,
 * skip <c> children inside <sheetData> (still tokenising the bytes — minizip
 * has to decompress them anyway), and capture every metadata element into
 * the worksheet's lxr_worksheet_meta cache. Accessors below read from this
 * cache.
 *
 * Hyperlinks may carry r:id pointing at xl/worksheets/_rels/sheetN.xml.rels;
 * the rels file is parsed separately to resolve those into absolute URLs.
 */

#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"

/* ------------------------------------------------------------------------- */
/* Cell-ref helpers                                                          */
/* ------------------------------------------------------------------------- */

static void parse_a1_ref(const char *ref, size_t *out_row, size_t *out_col)
{
    size_t col = 0, row = 0;
    const char *p = ref;
    if (out_row) *out_row = 0;
    if (out_col) *out_col = 0;
    if (!p) return;
    while (*p == '$') p++;
    while (*p && ((*p >= 'A' && *p <= 'Z') || (*p >= 'a' && *p <= 'z'))) {
        col = col * 26 + (toupper((unsigned char)*p) - 'A' + 1);
        p++;
    }
    while (*p == '$') p++;
    while (*p && *p >= '0' && *p <= '9') {
        row = row * 10 + (size_t)(*p - '0');
        p++;
    }
    if (out_row) *out_row = row;
    if (out_col) *out_col = col;
}

/* Parse "A1" or "A1:B2" into a range. Single-cell refs collapse to a 1x1. */
static void parse_a1_range(const char *ref, lxr_range *out)
{
    const char *colon;
    size_t r1 = 0, c1 = 0, r2 = 0, c2 = 0;

    if (out) { out->first_row = out->first_col = out->last_row = out->last_col = 0; }
    if (!ref || !out) return;

    colon = strchr(ref, ':');
    if (colon) {
        char head[64];
        size_t n = (size_t)(colon - ref);
        if (n >= sizeof(head)) n = sizeof(head) - 1;
        memcpy(head, ref, n);
        head[n] = 0;
        parse_a1_ref(head, &r1, &c1);
        parse_a1_ref(colon + 1, &r2, &c2);
    } else {
        parse_a1_ref(ref, &r1, &c1);
        r2 = r1;
        c2 = c1;
    }
    out->first_row = r1;
    out->first_col = c1;
    out->last_row  = r2;
    out->last_col  = c2;
}

static int attr_truthy(const char *v)
{
    if (!v) return 0;
    return (strcmp(v, "1") == 0 || strcmp(v, "true") == 0) ? 1 : 0;
}

/* ------------------------------------------------------------------------- */
/* Storage helpers                                                           */
/* ------------------------------------------------------------------------- */

static int meta_push_merge(lxr_worksheet_meta *m, lxr_range r)
{
    if (m->merges_count >= m->merges_cap) {
        size_t nc = m->merges_cap ? m->merges_cap * 2 : 8;
        lxr_range *nb = (lxr_range *)realloc(m->merges, nc * sizeof(*nb));
        if (!nb) return -1;
        m->merges = nb;
        m->merges_cap = nc;
    }
    m->merges[m->merges_count++] = r;
    return 0;
}

static int meta_push_hyperlink(lxr_worksheet_meta *m,
                               lxr_range r,
                               const char *url, const char *location,
                               const char *display, const char *tooltip)
{
    struct lxr_hyperlink_owned *h;
    if (m->hyperlinks_count >= m->hyperlinks_cap) {
        size_t nc = m->hyperlinks_cap ? m->hyperlinks_cap * 2 : 8;
        struct lxr_hyperlink_owned *nb = (struct lxr_hyperlink_owned *)realloc(
            m->hyperlinks, nc * sizeof(*nb));
        if (!nb) return -1;
        m->hyperlinks = nb;
        m->hyperlinks_cap = nc;
    }
    h = &m->hyperlinks[m->hyperlinks_count++];
    h->range    = r;
    h->url      = url      ? strdup(url)      : NULL;
    h->location = location ? strdup(location) : NULL;
    h->display  = display  ? strdup(display)  : NULL;
    h->tooltip  = tooltip  ? strdup(tooltip)  : NULL;
    return 0;
}

static struct lxr_row_meta *meta_get_or_make_row(lxr_worksheet_meta *m, size_t row)
{
    /* Rows arrive in order; a quick check on the tail is enough. */
    if (m->rows_count > 0 && m->rows[m->rows_count - 1].row == row) {
        return &m->rows[m->rows_count - 1];
    }
    if (m->rows_count >= m->rows_cap) {
        size_t nc = m->rows_cap ? m->rows_cap * 2 : 16;
        struct lxr_row_meta *nb = (struct lxr_row_meta *)realloc(
            m->rows, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->rows = nb;
        m->rows_cap = nc;
    }
    {
        struct lxr_row_meta *r = &m->rows[m->rows_count++];
        memset(r, 0, sizeof(*r));
        r->row = row;
        return r;
    }
}

static struct lxr_col_meta *meta_push_col(lxr_worksheet_meta *m)
{
    if (m->cols_count >= m->cols_cap) {
        size_t nc = m->cols_cap ? m->cols_cap * 2 : 8;
        struct lxr_col_meta *nb = (struct lxr_col_meta *)realloc(
            m->cols, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->cols = nb;
        m->cols_cap = nc;
    }
    {
        struct lxr_col_meta *c = &m->cols[m->cols_count++];
        memset(c, 0, sizeof(*c));
        return c;
    }
}

/* ------------------------------------------------------------------------- */
/* Sheet rels (rId -> URL) — simple dynamic map                              */
/* ------------------------------------------------------------------------- */

typedef struct {
    char *id;
    char *target;
} rel_entry;

typedef struct {
    rel_entry *entries;
    size_t     count;
    size_t     cap;
} rel_map;

static void rel_map_init(rel_map *m) { m->entries = NULL; m->count = 0; m->cap = 0; }
static void rel_map_free(rel_map *m)
{
    size_t i;
    if (!m) return;
    for (i = 0; i < m->count; i++) {
        free(m->entries[i].id);
        free(m->entries[i].target);
    }
    free(m->entries);
    m->entries = NULL; m->count = 0; m->cap = 0;
}
static int rel_map_push(rel_map *m, const char *id, const char *target)
{
    if (m->count >= m->cap) {
        size_t nc = m->cap ? m->cap * 2 : 4;
        rel_entry *nb = (rel_entry *)realloc(m->entries, nc * sizeof(*nb));
        if (!nb) return -1;
        m->entries = nb;
        m->cap = nc;
    }
    m->entries[m->count].id     = strdup(id);
    m->entries[m->count].target = strdup(target);
    m->count++;
    return 0;
}
static const char *rel_map_lookup(const rel_map *m, const char *id)
{
    size_t i;
    if (!m || !id) return NULL;
    for (i = 0; i < m->count; i++) {
        if (m->entries[i].id && strcmp(m->entries[i].id, id) == 0)
            return m->entries[i].target;
    }
    return NULL;
}

static void rel_on_start(void *ud, const char *name, const char **attrs)
{
    rel_map *m = (rel_map *)ud;
    if (!lxr_xml_name_eq(name, "Relationship")) return;
    {
        const char *id     = lxr_xml_attr(attrs, "Id");
        const char *target = lxr_xml_attr(attrs, "Target");
        if (id && target) rel_map_push(m, id, target);
    }
}

/* Open xl/worksheets/_rels/sheetN.xml.rels (if any) for the given sheet
 * target_path (e.g. "xl/worksheets/sheet1.xml"). Always returns NO_ERROR;
 * absent rels file just yields an empty map. */
static lxr_error load_sheet_rels(lxr_zip *zip, const char *target_path,
                                 rel_map *out)
{
    char         *rels_path;
    const char   *slash;
    size_t        prefix_len, file_len;
    const char   *file_part;
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    lxr_error     rc;

    rel_map_init(out);
    if (!zip || !target_path) return LXR_NO_ERROR;

    slash = strrchr(target_path, '/');
    prefix_len = slash ? (size_t)(slash - target_path) + 1 : 0;
    file_part  = slash ? slash + 1 : target_path;
    file_len   = strlen(file_part);

    rels_path = (char *)malloc(prefix_len + 6 + file_len + 5 + 1);
    if (!rels_path) return LXR_ERROR_MEMORY_MALLOC_FAILED;
    memcpy(rels_path, target_path, prefix_len);
    memcpy(rels_path + prefix_len, "_rels/", 6);
    memcpy(rels_path + prefix_len + 6, file_part, file_len);
    memcpy(rels_path + prefix_len + 6 + file_len, ".rels", 5);
    rels_path[prefix_len + 6 + file_len + 5] = 0;

    zf = lxr_zip_open_entry(zip, rels_path);
    free(rels_path);
    if (!zf) return LXR_NO_ERROR;  /* missing rels is fine */

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }

    lxr_xml_pump_set_handlers(pump, rel_on_start, NULL, NULL, out);
    rc = lxr_xml_pump_run(pump);

    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* ------------------------------------------------------------------------- */
/* Metadata SAX                                                              */
/* ------------------------------------------------------------------------- */

typedef enum {
    M_INIT = 0,
    M_IN_WORKSHEET,
    M_IN_SHEETDATA,
    M_IN_ROW,
    M_IN_COLS,
    M_IN_MERGECELLS,
    M_IN_HYPERLINKS,
    M_IN_DATAVALIDATIONS,
    M_IN_DATAVALIDATION,    /* parsing a single <dataValidation>; capturing
                              <formula1>/<formula2> children */
    M_IN_DV_FORMULA1,
    M_IN_DV_FORMULA2,
    M_IN_AUTOFILTER,
    M_IN_FILTERCOLUMN,
    M_IN_FILTERS,           /* <filters> inside a <filterColumn> */
    M_SKIP
} m_state;

typedef struct {
    lxr_worksheet_meta *m;
    const rel_map      *rels;
    m_state             state;
    m_state             state_before_skip;
    int                 skip_depth;
    char               *skip_tag;
    /* DV text accumulator — ECMA-376 puts formulas inside <formula1>/<formula2>
     * as text children, so we collect them like definedName text. */
    char               *txt;
    size_t              txt_len;
    size_t              txt_cap;
} m_ctx;

static void txt_append(m_ctx *c, const char *s, size_t n)
{
    size_t need = c->txt_len + n + 1;
    if (need > c->txt_cap) {
        size_t nc = c->txt_cap ? c->txt_cap : 64;
        char *nb;
        while (nc < need) nc *= 2;
        nb = (char *)realloc(c->txt, nc);
        if (!nb) return;
        c->txt = nb;
        c->txt_cap = nc;
    }
    memcpy(c->txt + c->txt_len, s, n);
    c->txt_len += n;
    c->txt[c->txt_len] = 0;
}

static void txt_reset(m_ctx *c) { c->txt_len = 0; if (c->txt) c->txt[0] = 0; }

/* DV/filter helpers — append entries. */

static struct lxr_dv_owned *meta_push_dv(lxr_worksheet_meta *m)
{
    if (m->dvs_count >= m->dvs_cap) {
        size_t nc = m->dvs_cap ? m->dvs_cap * 2 : 4;
        struct lxr_dv_owned *nb = (struct lxr_dv_owned *)realloc(
            m->dvs, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->dvs = nb;
        m->dvs_cap = nc;
    }
    {
        struct lxr_dv_owned *d = &m->dvs[m->dvs_count++];
        memset(d, 0, sizeof(*d));
        return d;
    }
}

static struct lxr_filter_column_owned *meta_push_filter_column(lxr_worksheet_meta *m)
{
    if (m->filter_columns_count >= m->filter_columns_cap) {
        size_t nc = m->filter_columns_cap ? m->filter_columns_cap * 2 : 4;
        struct lxr_filter_column_owned *nb = (struct lxr_filter_column_owned *)realloc(
            m->filter_columns, nc * sizeof(*nb));
        if (!nb) return NULL;
        m->filter_columns = nb;
        m->filter_columns_cap = nc;
    }
    {
        struct lxr_filter_column_owned *fc = &m->filter_columns[m->filter_columns_count++];
        memset(fc, 0, sizeof(*fc));
        return fc;
    }
}

/* For LXR_FILTER_LIST: append to a NULL-terminated owned values array. */
static void filter_column_push_value(struct lxr_filter_column_owned *fc, const char *val)
{
    size_t n = 0;
    char **nv;
    if (!val) return;
    if (fc->values) while (fc->values[n]) n++;
    nv = (char **)realloc(fc->values, (n + 2) * sizeof(*nv));
    if (!nv) return;
    nv[n]     = strdup(val);
    nv[n + 1] = NULL;
    fc->values = nv;
}

static void enter_skip(m_ctx *c, const char *tag)
{
    free(c->skip_tag);
    c->skip_tag = strdup(tag);
    c->state_before_skip = c->state;
    c->state = M_SKIP;
    c->skip_depth = 1;
}

static void m_on_start(void *ud, const char *name, const char **attrs)
{
    m_ctx *c = (m_ctx *)ud;

    if (c->state == M_SKIP) { c->skip_depth++; return; }

    switch (c->state) {
    case M_INIT:
        if (lxr_xml_name_eq(name, "worksheet")) c->state = M_IN_WORKSHEET;
        break;

    case M_IN_WORKSHEET:
        if (lxr_xml_name_eq(name, "sheetData")) {
            c->state = M_IN_SHEETDATA;
        } else if (lxr_xml_name_eq(name, "sheetFormatPr")) {
            const char *drh = lxr_xml_attr(attrs, "defaultRowHeight");
            const char *dcw = lxr_xml_attr(attrs, "defaultColWidth");
            if (drh) {
                c->m->has_default_row_height = 1;
                c->m->default_row_height = strtod(drh, NULL);
            }
            if (dcw) {
                c->m->has_default_col_width = 1;
                c->m->default_col_width = strtod(dcw, NULL);
            }
            /* element is empty / self-closing in practice — stay in WORKSHEET */
        } else if (lxr_xml_name_eq(name, "cols")) {
            c->state = M_IN_COLS;
        } else if (lxr_xml_name_eq(name, "mergeCells")) {
            c->state = M_IN_MERGECELLS;
        } else if (lxr_xml_name_eq(name, "hyperlinks")) {
            c->state = M_IN_HYPERLINKS;
        } else if (lxr_xml_name_eq(name, "dataValidations")) {
            c->state = M_IN_DATAVALIDATIONS;
        } else if (lxr_xml_name_eq(name, "autoFilter")) {
            const char *ref = lxr_xml_attr(attrs, "ref");
            c->m->autofilter_present = 1;
            if (ref) {
                free(c->m->autofilter_range);
                c->m->autofilter_range = strdup(ref);
            }
            c->state = M_IN_AUTOFILTER;
        } else if (lxr_xml_name_eq(name, "sheetProtection")) {
            const char *v;
            c->m->prot_present = 1;
            v = lxr_xml_attr(attrs, "password");
            if (v) {
                size_t n = strlen(v);
                if (n >= sizeof(c->m->prot_hash)) n = sizeof(c->m->prot_hash) - 1;
                memcpy(c->m->prot_hash, v, n);
                c->m->prot_hash[n] = 0;
            }
            c->m->prot_sheet               = attr_truthy(lxr_xml_attr(attrs, "sheet"));
            c->m->prot_content             = attr_truthy(lxr_xml_attr(attrs, "content"));
            c->m->prot_objects             = attr_truthy(lxr_xml_attr(attrs, "objects"));
            c->m->prot_scenarios           = attr_truthy(lxr_xml_attr(attrs, "scenarios"));
            c->m->prot_format_cells        = attr_truthy(lxr_xml_attr(attrs, "formatCells"));
            c->m->prot_format_columns      = attr_truthy(lxr_xml_attr(attrs, "formatColumns"));
            c->m->prot_format_rows         = attr_truthy(lxr_xml_attr(attrs, "formatRows"));
            c->m->prot_insert_columns      = attr_truthy(lxr_xml_attr(attrs, "insertColumns"));
            c->m->prot_insert_rows         = attr_truthy(lxr_xml_attr(attrs, "insertRows"));
            c->m->prot_insert_hyperlinks   = attr_truthy(lxr_xml_attr(attrs, "insertHyperlinks"));
            c->m->prot_delete_columns      = attr_truthy(lxr_xml_attr(attrs, "deleteColumns"));
            c->m->prot_delete_rows         = attr_truthy(lxr_xml_attr(attrs, "deleteRows"));
            c->m->prot_select_locked_cells = attr_truthy(lxr_xml_attr(attrs, "selectLockedCells"));
            c->m->prot_sort                = attr_truthy(lxr_xml_attr(attrs, "sort"));
            c->m->prot_auto_filter         = attr_truthy(lxr_xml_attr(attrs, "autoFilter"));
            c->m->prot_pivot_tables        = attr_truthy(lxr_xml_attr(attrs, "pivotTables"));
            c->m->prot_select_unlocked_cells = attr_truthy(lxr_xml_attr(attrs, "selectUnlockedCells"));
        }
        break;

    case M_IN_COLS:
        if (lxr_xml_name_eq(name, "col")) {
            const char *vmin = lxr_xml_attr(attrs, "min");
            const char *vmax = lxr_xml_attr(attrs, "max");
            const char *vw   = lxr_xml_attr(attrs, "width");
            const char *cw   = lxr_xml_attr(attrs, "customWidth");
            const char *hd   = lxr_xml_attr(attrs, "hidden");
            const char *ol   = lxr_xml_attr(attrs, "outlineLevel");
            const char *cl   = lxr_xml_attr(attrs, "collapsed");
            struct lxr_col_meta *col = meta_push_col(c->m);
            if (col) {
                col->min = vmin ? (size_t)strtoul(vmin, NULL, 10) : 0;
                col->max = vmax ? (size_t)strtoul(vmax, NULL, 10) : col->min;
                if (vw) {
                    col->width = strtod(vw, NULL);
                    /* OOXML lists customWidth=1 when an explicit width was set;
                     * absent customWidth on a width attr is the default-derived
                     * value the writer emits. Treat any present width as
                     * "user-visible width". */
                    col->has_width = 1;
                    (void)cw;
                }
                col->hidden        = attr_truthy(hd);
                col->outline_level = ol ? (int)strtol(ol, NULL, 10) : 0;
                col->collapsed     = attr_truthy(cl);
            }
        }
        break;

    case M_IN_MERGECELLS:
        if (lxr_xml_name_eq(name, "mergeCell")) {
            const char *ref = lxr_xml_attr(attrs, "ref");
            lxr_range r;
            parse_a1_range(ref, &r);
            if (r.first_row && r.first_col) meta_push_merge(c->m, r);
        }
        break;

    case M_IN_DATAVALIDATIONS:
        if (lxr_xml_name_eq(name, "dataValidation")) {
            struct lxr_dv_owned *d = meta_push_dv(c->m);
            const char *v;
            if (!d) break;
            if ((v = lxr_xml_attr(attrs, "type")))               d->type        = strdup(v);
            if ((v = lxr_xml_attr(attrs, "operator")))           d->operator_   = strdup(v);
            if ((v = lxr_xml_attr(attrs, "errorStyle")))         d->error_style = strdup(v);
            if ((v = lxr_xml_attr(attrs, "prompt")))             d->prompt      = strdup(v);
            if ((v = lxr_xml_attr(attrs, "promptTitle")))        d->prompt_title= strdup(v);
            if ((v = lxr_xml_attr(attrs, "error")))              d->error       = strdup(v);
            if ((v = lxr_xml_attr(attrs, "errorTitle")))         d->error_title = strdup(v);
            if ((v = lxr_xml_attr(attrs, "sqref")))              d->sqref       = strdup(v);
            d->allow_blank          = attr_truthy(lxr_xml_attr(attrs, "allowBlank"));
            d->show_drop_down       = attr_truthy(lxr_xml_attr(attrs, "showDropDown"));
            d->show_input_message   = attr_truthy(lxr_xml_attr(attrs, "showInputMessage"));
            d->show_error_message   = attr_truthy(lxr_xml_attr(attrs, "showErrorMessage"));
            c->state = M_IN_DATAVALIDATION;
        }
        break;

    case M_IN_DATAVALIDATION:
        if (lxr_xml_name_eq(name, "formula1")) {
            txt_reset(c);
            c->state = M_IN_DV_FORMULA1;
        } else if (lxr_xml_name_eq(name, "formula2")) {
            txt_reset(c);
            c->state = M_IN_DV_FORMULA2;
        } else {
            enter_skip(c, name);
        }
        break;

    case M_IN_AUTOFILTER:
        if (lxr_xml_name_eq(name, "filterColumn")) {
            const char *cid = lxr_xml_attr(attrs, "colId");
            struct lxr_filter_column_owned *fc = meta_push_filter_column(c->m);
            if (fc) {
                fc->col_id = cid ? (int)strtol(cid, NULL, 10) : 0;
                fc->kind   = LXR_FILTER_NONE;
            }
            c->state = M_IN_FILTERCOLUMN;
        } else {
            enter_skip(c, name);
        }
        break;

    case M_IN_FILTERCOLUMN:
        if (c->m->filter_columns_count == 0) { enter_skip(c, name); break; }
        {
            struct lxr_filter_column_owned *fc =
                &c->m->filter_columns[c->m->filter_columns_count - 1];
            if (lxr_xml_name_eq(name, "filters")) {
                fc->kind = LXR_FILTER_LIST;
                c->state = M_IN_FILTERS;
            } else if (lxr_xml_name_eq(name, "customFilters")) {
                const char *and_attr = lxr_xml_attr(attrs, "and");
                fc->kind = LXR_FILTER_CUSTOM;
                fc->custom_and = attr_truthy(and_attr);
                /* children customFilter come in <customFilter operator val>;
                 * parse them inline by transitioning into a list-style scan. */
                c->state = M_IN_FILTERS;  /* reuse */
            } else if (lxr_xml_name_eq(name, "top10")) {
                const char *val_attr  = lxr_xml_attr(attrs, "val");
                const char *top_attr  = lxr_xml_attr(attrs, "top");
                const char *pct_attr  = lxr_xml_attr(attrs, "percent");
                fc->kind      = LXR_FILTER_TOP10;
                fc->top       = top_attr ? attr_truthy(top_attr) : 1;
                fc->percent   = attr_truthy(pct_attr);
                fc->top_value = val_attr ? strtod(val_attr, NULL) : 0;
                enter_skip(c, name);
            } else if (lxr_xml_name_eq(name, "dynamicFilter")) {
                fc->kind = LXR_FILTER_DYNAMIC;
                enter_skip(c, name);
            } else {
                enter_skip(c, name);
            }
        }
        break;

    case M_IN_FILTERS:
        if (c->m->filter_columns_count == 0) { enter_skip(c, name); break; }
        {
            struct lxr_filter_column_owned *fc =
                &c->m->filter_columns[c->m->filter_columns_count - 1];
            if (lxr_xml_name_eq(name, "filter")) {
                const char *v = lxr_xml_attr(attrs, "val");
                if (v) filter_column_push_value(fc, v);
                enter_skip(c, name);
            } else if (lxr_xml_name_eq(name, "customFilter")) {
                const char *op  = lxr_xml_attr(attrs, "operator");
                const char *val = lxr_xml_attr(attrs, "val");
                if (!fc->custom_op_1) {
                    fc->custom_op_1  = op  ? strdup(op)  : strdup("equal");
                    fc->custom_val_1 = val ? strdup(val) : NULL;
                } else if (!fc->custom_op_2) {
                    fc->custom_op_2  = op  ? strdup(op)  : strdup("equal");
                    fc->custom_val_2 = val ? strdup(val) : NULL;
                }
                enter_skip(c, name);
            } else {
                enter_skip(c, name);
            }
        }
        break;

    case M_IN_HYPERLINKS:
        if (lxr_xml_name_eq(name, "hyperlink")) {
            const char *ref     = lxr_xml_attr(attrs, "ref");
            const char *loc     = lxr_xml_attr(attrs, "location");
            const char *display = lxr_xml_attr(attrs, "display");
            const char *tooltip = lxr_xml_attr(attrs, "tooltip");
            const char *rid     = lxr_xml_attr(attrs, "id");
            const char *url     = NULL;
            lxr_range r;

            if (!rid) {
                /* "r:id" — namespace lookup by walking attrs. */
                const char **a = attrs;
                while (a && *a) {
                    if (lxr_xml_name_eq(*a, "id")) { rid = *(a + 1); break; }
                    a += 2;
                }
            }
            if (rid) url = rel_map_lookup(c->rels, rid);

            parse_a1_range(ref, &r);
            if (r.first_row && r.first_col)
                meta_push_hyperlink(c->m, r, url, loc, display, tooltip);
        }
        break;

    case M_IN_SHEETDATA:
        if (lxr_xml_name_eq(name, "row")) {
            const char *r_attr  = lxr_xml_attr(attrs, "r");
            const char *ht      = lxr_xml_attr(attrs, "ht");
            const char *hidden  = lxr_xml_attr(attrs, "hidden");
            const char *ol      = lxr_xml_attr(attrs, "outlineLevel");
            const char *cl      = lxr_xml_attr(attrs, "collapsed");
            const char *ch      = lxr_xml_attr(attrs, "customHeight");
            size_t row_nr = r_attr ? (size_t)strtoul(r_attr, NULL, 10) : 0;
            /* Only cache rows that carry at least one metadata attr — else
             * getRowOptions(N) would return an empty struct for rows whose
             * only purpose is holding cells. */
            int has_meta = (ht || ol ||
                            attr_truthy(hidden) ||
                            attr_truthy(cl) ||
                            attr_truthy(ch));
            if (row_nr > 0 && has_meta) {
                struct lxr_row_meta *rm = meta_get_or_make_row(c->m, row_nr);
                if (rm) {
                    if (ht) {
                        rm->has_height = 1;
                        rm->height = strtod(ht, NULL);
                    }
                    rm->hidden        = attr_truthy(hidden);
                    rm->outline_level = ol ? (int)strtol(ol, NULL, 10) : 0;
                    rm->collapsed     = attr_truthy(cl);
                    rm->custom_height = attr_truthy(ch);
                }
            }
            c->state = M_IN_ROW;
        }
        break;

    case M_IN_ROW:
        /* Skip every <c> and child — metadata only cares about row attrs. */
        enter_skip(c, name);
        break;

    default:
        break;
    }
}

static void m_on_end(void *ud, const char *name)
{
    m_ctx *c = (m_ctx *)ud;

    if (c->state == M_SKIP) {
        c->skip_depth--;
        if (c->skip_depth == 0 && c->skip_tag &&
            lxr_xml_name_eq(name, c->skip_tag)) {
            free(c->skip_tag);
            c->skip_tag = NULL;
            c->state = c->state_before_skip;
        }
        return;
    }

    switch (c->state) {
    case M_IN_ROW:
        if (lxr_xml_name_eq(name, "row")) c->state = M_IN_SHEETDATA;
        break;
    case M_IN_SHEETDATA:
        if (lxr_xml_name_eq(name, "sheetData")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_COLS:
        if (lxr_xml_name_eq(name, "cols")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_MERGECELLS:
        if (lxr_xml_name_eq(name, "mergeCells")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_HYPERLINKS:
        if (lxr_xml_name_eq(name, "hyperlinks")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_DATAVALIDATION:
        if (lxr_xml_name_eq(name, "dataValidation")) c->state = M_IN_DATAVALIDATIONS;
        break;
    case M_IN_DATAVALIDATIONS:
        if (lxr_xml_name_eq(name, "dataValidations")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_DV_FORMULA1:
        if (lxr_xml_name_eq(name, "formula1")) {
            if (c->m->dvs_count > 0 && c->txt_len > 0) {
                struct lxr_dv_owned *d = &c->m->dvs[c->m->dvs_count - 1];
                free(d->formula1);
                d->formula1 = strdup(c->txt);
            }
            c->state = M_IN_DATAVALIDATION;
        }
        break;
    case M_IN_DV_FORMULA2:
        if (lxr_xml_name_eq(name, "formula2")) {
            if (c->m->dvs_count > 0 && c->txt_len > 0) {
                struct lxr_dv_owned *d = &c->m->dvs[c->m->dvs_count - 1];
                free(d->formula2);
                d->formula2 = strdup(c->txt);
            }
            c->state = M_IN_DATAVALIDATION;
        }
        break;
    case M_IN_FILTERCOLUMN:
        if (lxr_xml_name_eq(name, "filterColumn")) c->state = M_IN_AUTOFILTER;
        break;
    case M_IN_FILTERS:
        if (lxr_xml_name_eq(name, "filters") ||
            lxr_xml_name_eq(name, "customFilters")) {
            c->state = M_IN_FILTERCOLUMN;
        }
        break;
    case M_IN_AUTOFILTER:
        if (lxr_xml_name_eq(name, "autoFilter")) c->state = M_IN_WORKSHEET;
        break;
    case M_IN_WORKSHEET:
        if (lxr_xml_name_eq(name, "worksheet")) c->state = M_INIT;
        break;
    default:
        break;
    }
}

/* Text capture for <formula1>/<formula2>. */
static void m_on_text(void *ud, const char *text, int len)
{
    m_ctx *c = (m_ctx *)ud;
    if (len <= 0) return;
    if (c->state == M_IN_DV_FORMULA1 || c->state == M_IN_DV_FORMULA2) {
        txt_append(c, text, (size_t)len);
    }
}

/* ------------------------------------------------------------------------- */
/* Public load / free                                                         */
/* ------------------------------------------------------------------------- */

lxr_error lxr_worksheet_meta_load(lxr_worksheet *ws)
{
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    rel_map       rels;
    m_ctx         ctx;
    lxr_error     rc;

    if (!ws || !ws->wb || !ws->target_path) return LXR_ERROR_NULL_PARAMETER;

    /* Sheet rels first (required to resolve external hyperlinks). */
    rc = load_sheet_rels(ws->wb->zip, ws->target_path, &rels);
    if (rc != LXR_NO_ERROR) { rel_map_free(&rels); return rc; }

    zf = lxr_zip_open_entry(ws->wb->zip, ws->target_path);
    if (!zf) { rel_map_free(&rels); return LXR_ERROR_ZIP_ENTRY_NOT_FOUND; }

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxr_zip_close_entry(zf);
        rel_map_free(&rels);
        return LXR_ERROR_MEMORY_MALLOC_FAILED;
    }

    memset(&ctx, 0, sizeof(ctx));
    ctx.m     = &ws->meta;
    ctx.rels  = &rels;
    ctx.state = M_INIT;

    lxr_xml_pump_set_handlers(pump, m_on_start, m_on_end, m_on_text, &ctx);
    rc = lxr_xml_pump_run(pump);

    free(ctx.skip_tag);
    free(ctx.txt);
    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    rel_map_free(&rels);
    return rc;
}

void lxr_worksheet_meta_free(lxr_worksheet_meta *m)
{
    size_t i;
    if (!m) return;
    free(m->merges);
    if (m->hyperlinks) {
        for (i = 0; i < m->hyperlinks_count; i++) {
            free(m->hyperlinks[i].url);
            free(m->hyperlinks[i].location);
            free(m->hyperlinks[i].display);
            free(m->hyperlinks[i].tooltip);
        }
        free(m->hyperlinks);
    }
    free(m->rows);
    free(m->cols);
    if (m->dvs) {
        for (i = 0; i < m->dvs_count; i++) {
            free(m->dvs[i].type);
            free(m->dvs[i].operator_);
            free(m->dvs[i].error_style);
            free(m->dvs[i].formula1);
            free(m->dvs[i].formula2);
            free(m->dvs[i].prompt);
            free(m->dvs[i].prompt_title);
            free(m->dvs[i].error);
            free(m->dvs[i].error_title);
            free(m->dvs[i].sqref);
        }
        free(m->dvs);
    }
    free(m->autofilter_range);
    if (m->filter_columns) {
        for (i = 0; i < m->filter_columns_count; i++) {
            if (m->filter_columns[i].values) {
                size_t j = 0;
                while (m->filter_columns[i].values[j]) {
                    free(m->filter_columns[i].values[j]);
                    j++;
                }
                free(m->filter_columns[i].values);
            }
            free(m->filter_columns[i].custom_op_1);
            free(m->filter_columns[i].custom_val_1);
            free(m->filter_columns[i].custom_op_2);
            free(m->filter_columns[i].custom_val_2);
        }
        free(m->filter_columns);
    }
    free(m->filter_columns_pub);
    memset(m, 0, sizeof(*m));
}

/* ------------------------------------------------------------------------- */
/* Public accessors                                                          */
/* ------------------------------------------------------------------------- */

size_t lxr_worksheet_merged_count(const lxr_worksheet *ws)
{
    return ws ? ws->meta.merges_count : 0;
}

int lxr_worksheet_merged_get(const lxr_worksheet *ws, size_t idx, lxr_range *out)
{
    if (!ws || !out || idx >= ws->meta.merges_count) return 0;
    *out = ws->meta.merges[idx];
    return 1;
}

int lxr_worksheet_in_merge_follow(const lxr_worksheet *ws, size_t row, size_t col)
{
    size_t i;
    if (!ws) return 0;
    for (i = 0; i < ws->meta.merges_count; i++) {
        const lxr_range *r = &ws->meta.merges[i];
        if (row >= r->first_row && row <= r->last_row &&
            col >= r->first_col && col <= r->last_col) {
            /* Inside this range. Master is (first_row, first_col). */
            return (row == r->first_row && col == r->first_col) ? 0 : 1;
        }
    }
    return 0;
}

size_t lxr_worksheet_hyperlink_count(const lxr_worksheet *ws)
{
    return ws ? ws->meta.hyperlinks_count : 0;
}

int lxr_worksheet_hyperlink_get(const lxr_worksheet *ws, size_t idx,
                                lxr_hyperlink *out)
{
    const struct lxr_hyperlink_owned *h;
    if (!ws || !out || idx >= ws->meta.hyperlinks_count) return 0;
    h = &ws->meta.hyperlinks[idx];
    out->range    = h->range;
    out->url      = h->url;
    out->location = h->location;
    out->display  = h->display;
    out->tooltip  = h->tooltip;
    return 1;
}

const char *lxr_worksheet_hyperlink_url(const lxr_worksheet *ws,
                                        size_t row, size_t col)
{
    size_t i;
    if (!ws) return NULL;
    for (i = 0; i < ws->meta.hyperlinks_count; i++) {
        const struct lxr_hyperlink_owned *h = &ws->meta.hyperlinks[i];
        if (row >= h->range.first_row && row <= h->range.last_row &&
            col >= h->range.first_col && col <= h->range.last_col) {
            return h->url ? h->url : h->location;
        }
    }
    return NULL;
}

int lxr_worksheet_protection(const lxr_worksheet *ws, lxr_protection *out)
{
    const lxr_worksheet_meta *m;
    if (!ws || !out) return 0;
    m = &ws->meta;
    out->is_present                = m->prot_present;
    {
        size_t n = strlen(m->prot_hash);
        if (n >= sizeof(out->password_hash)) n = sizeof(out->password_hash) - 1;
        memcpy(out->password_hash, m->prot_hash, n);
        out->password_hash[n] = 0;
    }
    out->sheet                = m->prot_sheet;
    out->content              = m->prot_content;
    out->objects              = m->prot_objects;
    out->scenarios            = m->prot_scenarios;
    out->format_cells         = m->prot_format_cells;
    out->format_columns       = m->prot_format_columns;
    out->format_rows          = m->prot_format_rows;
    out->insert_columns       = m->prot_insert_columns;
    out->insert_rows          = m->prot_insert_rows;
    out->insert_hyperlinks    = m->prot_insert_hyperlinks;
    out->delete_columns       = m->prot_delete_columns;
    out->delete_rows          = m->prot_delete_rows;
    out->select_locked_cells  = m->prot_select_locked_cells;
    out->sort                 = m->prot_sort;
    out->auto_filter          = m->prot_auto_filter;
    out->pivot_tables         = m->prot_pivot_tables;
    out->select_unlocked_cells = m->prot_select_unlocked_cells;
    return m->prot_present;
}

int lxr_worksheet_row_options(const lxr_worksheet *ws, size_t row,
                              lxr_row_options *out)
{
    size_t i;
    if (!ws || !out || row == 0) return 0;
    for (i = 0; i < ws->meta.rows_count; i++) {
        const struct lxr_row_meta *r = &ws->meta.rows[i];
        if (r->row == row) {
            out->has_height    = r->has_height;
            out->height        = r->height;
            out->hidden        = r->hidden;
            out->outline_level = r->outline_level;
            out->collapsed     = r->collapsed;
            out->custom_height = r->custom_height;
            return 1;
        }
    }
    return 0;
}

int lxr_worksheet_col_options(const lxr_worksheet *ws, size_t col,
                              lxr_col_options *out)
{
    size_t i;
    if (!ws || !out || col == 0) return 0;
    for (i = 0; i < ws->meta.cols_count; i++) {
        const struct lxr_col_meta *c = &ws->meta.cols[i];
        if (col >= c->min && col <= c->max) {
            out->has_width     = c->has_width;
            out->width         = c->width;
            out->hidden        = c->hidden;
            out->outline_level = c->outline_level;
            out->collapsed     = c->collapsed;
            return 1;
        }
    }
    return 0;
}

int lxr_worksheet_default_row_height(const lxr_worksheet *ws, double *out)
{
    if (!ws || !out || !ws->meta.has_default_row_height) return 0;
    *out = ws->meta.default_row_height;
    return 1;
}

int lxr_worksheet_default_col_width(const lxr_worksheet *ws, double *out)
{
    if (!ws || !out || !ws->meta.has_default_col_width) return 0;
    *out = ws->meta.default_col_width;
    return 1;
}

/* ---- Data validation accessors ----------------------------------------- */

size_t lxr_worksheet_data_validation_count(const lxr_worksheet *ws)
{
    return ws ? ws->meta.dvs_count : 0;
}

int lxr_worksheet_data_validation_get(const lxr_worksheet *ws, size_t idx,
                                      lxr_data_validation *out)
{
    const struct lxr_dv_owned *d;
    if (!ws || !out || idx >= ws->meta.dvs_count) return 0;
    d = &ws->meta.dvs[idx];
    out->type               = d->type;
    out->operator_          = d->operator_;
    out->error_style        = d->error_style;
    out->formula1           = d->formula1;
    out->formula2           = d->formula2;
    out->prompt             = d->prompt;
    out->prompt_title       = d->prompt_title;
    out->error              = d->error;
    out->error_title        = d->error_title;
    out->sqref              = d->sqref;
    out->allow_blank        = d->allow_blank;
    out->show_drop_down     = d->show_drop_down;
    out->show_input_message = d->show_input_message;
    out->show_error_message = d->show_error_message;
    return 1;
}

/* ---- AutoFilter accessor ------------------------------------------------ */

int lxr_worksheet_autofilter(const lxr_worksheet *ws, lxr_autofilter *out)
{
    lxr_worksheet_meta *m;
    size_t i;
    if (!ws || !out) return 0;
    m = (lxr_worksheet_meta *)&ws->meta;  /* mutable for cache materialisation */
    if (!m->autofilter_present) return 0;

    /* Lazily materialise a stable public-shape array of filter columns so
     * out->columns has a predictable lifetime (same as the worksheet). */
    if (!m->filter_columns_pub && m->filter_columns_count > 0) {
        m->filter_columns_pub = (lxr_filter_column *)
            calloc(m->filter_columns_count, sizeof(lxr_filter_column));
        if (m->filter_columns_pub) {
            for (i = 0; i < m->filter_columns_count; i++) {
                const struct lxr_filter_column_owned *src = &m->filter_columns[i];
                lxr_filter_column *dst = &m->filter_columns_pub[i];
                dst->col_id        = src->col_id;
                dst->kind          = (lxr_filter_kind)src->kind;
                dst->values        = (const char **)src->values;
                dst->custom_and    = src->custom_and;
                dst->custom_op_1   = src->custom_op_1;
                dst->custom_val_1  = src->custom_val_1;
                dst->custom_op_2   = src->custom_op_2;
                dst->custom_val_2  = src->custom_val_2;
                dst->top           = src->top;
                dst->percent       = src->percent;
                dst->top_value     = src->top_value;
            }
        }
    }

    out->range         = m->autofilter_range;
    out->columns       = m->filter_columns_pub;
    out->columns_count = m->filter_columns_count;
    return 1;
}
