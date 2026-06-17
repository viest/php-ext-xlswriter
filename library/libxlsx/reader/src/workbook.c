#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxlsx_reader_internal.h"
#include "lxlsx_reader_xml_pump.h"

/* Well-known Open Packaging Conventions strings. */
static const char *CT_WORKBOOK_XLSX =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
static const char *CT_WORKBOOK_XLSM =
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
static const char *CT_WORKBOOK_XLTX =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
static const char *CT_WORKBOOK_XLTM =
    "application/vnd.ms-excel.template.macroEnabled.main+xml";

static const char *REL_TYPE_WORKSHEET =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
static const char *REL_TYPE_SHARED_STRINGS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
static const char *REL_TYPE_STYLES =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

/* --- string utilities ---------------------------------------------------- */

static char *str_dup(const char *s)
{
    if (!s) return NULL;
    return strdup(s);
}

static char *normalize_zip_path(char *path)
{
    char **segments;
    char *p;
    size_t count = 0;
    size_t i;
    size_t len;
    size_t out_len = 0;
    char *out;
    char *write;

    if (!path) return NULL;
    len = strlen(path);
    segments = (char **)calloc(len + 1, sizeof(*segments));
    if (!segments) {
        free(path);
        return NULL;
    }

    p = path;
    while (*p) {
        char *seg;
        while (*p == '/') p++;
        if (!*p) break;
        seg = p;
        while (*p && *p != '/') p++;
        if (*p == '/') *p++ = 0;

        if (strcmp(seg, ".") == 0 || seg[0] == 0) {
            continue;
        } else if (strcmp(seg, "..") == 0) {
            if (count > 0) count--;
        } else {
            segments[count++] = seg;
        }
    }

    for (i = 0; i < count; i++)
        out_len += strlen(segments[i]) + (i ? 1 : 0);

    out = (char *)malloc(out_len + 1);
    if (!out) {
        free(segments);
        free(path);
        return NULL;
    }

    write = out;
    for (i = 0; i < count; i++) {
        size_t seg_len = strlen(segments[i]);
        if (i) *write++ = '/';
        memcpy(write, segments[i], seg_len);
        write += seg_len;
    }
    *write = 0;

    free(segments);
    free(path);
    return out;
}

/* Normalize a workbook-relative path to a zip-absolute path.
 * Examples:
 *   base="xl/", target="worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
 *   base="xl/", target="./worksheets/../worksheets/custom.xml" -> "xl/worksheets/custom.xml"
 *   base="xl/", target="/xl/sharedStrings.xml" -> "xl/sharedStrings.xml"
 */
static char *join_base(const char *base, const char *target)
{
    size_t bl, tl;
    char  *out;
    if (!target) return NULL;
    if (target[0] == '/') return normalize_zip_path(str_dup(target + 1));
    bl = base ? strlen(base) : 0;
    tl = strlen(target);
    out = (char *)malloc(bl + tl + 1);
    if (!out) return NULL;
    if (bl) memcpy(out, base, bl);
    memcpy(out + bl, target, tl);
    out[bl + tl] = 0;
    return normalize_zip_path(out);
}

static char *base_path_of(const char *path)
{
    const char *slash = path ? strrchr(path, '/') : NULL;
    char *out;
    size_t n;
    if (!slash) return str_dup("");
    n = (size_t)(slash - path) + 1;
    out = (char *)malloc(n + 1);
    if (!out) return NULL;
    memcpy(out, path, n);
    out[n] = 0;
    return out;
}

static char *rels_path_of(const char *path)
{
    /* "xl/workbook.xml" -> "xl/_rels/workbook.xml.rels"
     * "/sheet1.xml"     -> "_rels/sheet1.xml.rels" */
    const char *slash = path ? strrchr(path, '/') : NULL;
    size_t prefix_len = slash ? (size_t)(slash - path) + 1 : 0;
    const char *file = slash ? slash + 1 : path;
    size_t flen = strlen(file);
    char *out = (char *)malloc(prefix_len + 6 + flen + 5 + 1);
    if (!out) return NULL;
    memcpy(out, path, prefix_len);
    memcpy(out + prefix_len, "_rels/", 6);
    memcpy(out + prefix_len + 6, file, flen);
    memcpy(out + prefix_len + 6 + flen, ".rels", 5);
    out[prefix_len + 6 + flen + 5] = 0;
    return out;
}

static int sheet_array_push(lxlsx_reader_workbook *wb, const char *name, const char *rid,
                            lxlsx_reader_sheet_visibility vis)
{
    if (wb->sheet_count >= wb->sheet_cap) {
        size_t cap = wb->sheet_cap ? wb->sheet_cap * 2 : 8;
        lxlsx_reader_sheet_info *nb = (lxlsx_reader_sheet_info *)realloc(
            wb->sheets, cap * sizeof(*nb));
        if (!nb) return -1;
        wb->sheets    = nb;
        wb->sheet_cap = cap;
    }
    wb->sheets[wb->sheet_count].name       = str_dup(name);
    wb->sheets[wb->sheet_count].rel_id     = str_dup(rid);
    wb->sheets[wb->sheet_count].target     = NULL;
    wb->sheets[wb->sheet_count].visibility = (int)vis;
    wb->sheet_count++;
    return 0;
}

/* ------------------------------------------------------------------------- */
/* [Content_Types].xml                                                       */
/* ------------------------------------------------------------------------- */

typedef struct {
    char *workbook_path;
} ct_ctx;

static void ct_on_start(void *ud, const char *name, const char **attrs)
{
    ct_ctx *c = (ct_ctx *)ud;
    if (!lxlsx_reader_xml_name_eq(name, "Override")) return;
    {
        const char *part_name = lxlsx_reader_xml_attr(attrs, "PartName");
        const char *content   = lxlsx_reader_xml_attr(attrs, "ContentType");
        if (!part_name || !content) return;

        if (strcmp(content, CT_WORKBOOK_XLSX) == 0 ||
            strcmp(content, CT_WORKBOOK_XLSM) == 0 ||
            strcmp(content, CT_WORKBOOK_XLTX) == 0 ||
            strcmp(content, CT_WORKBOOK_XLTM) == 0) {
            free(c->workbook_path);
            /* PartName starts with leading '/'; strip it. */
            c->workbook_path = str_dup(part_name[0] == '/' ? part_name + 1 : part_name);
        }
    }
}

static lxlsx_reader_error find_workbook_path(lxlsx_reader_zip *zip, char **out_path)
{
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    ct_ctx        ctx = {NULL};
    lxlsx_reader_error     rc;

    zf = lxlsx_reader_zip_open_entry(zip, "[Content_Types].xml");
    if (!zf) return LXLSX_READER_ERROR_FILE_NOT_XLSX;

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxlsx_reader_zip_close_entry(zf);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }

    lxlsx_reader_xml_pump_set_handlers(pump, ct_on_start, NULL, NULL, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);

    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);

    if (rc != LXLSX_READER_NO_ERROR) {
        free(ctx.workbook_path);
        return rc;
    }
    if (!ctx.workbook_path) return LXLSX_READER_ERROR_FILE_NOT_XLSX;
    *out_path = ctx.workbook_path;
    return LXLSX_READER_NO_ERROR;
}

/* ------------------------------------------------------------------------- */
/* xl/workbook.xml                                                           */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxlsx_reader_workbook *wb;
    int  in_sheets;
    int  in_defined_names;
    int  in_defined_name;       /* inside a single <definedName> */
    char *dn_name;              /* current entry's name attr */
    int   dn_local_sheet;       /* localSheetId or -1 */
    int   dn_hidden;
    char *dn_text;              /* accumulated formula body */
    size_t dn_text_len;
    size_t dn_text_cap;
} wb_ctx;

static int dn_array_push(lxlsx_reader_workbook *wb, const char *name, const char *formula,
                         int scope_sheet_index, int hidden)
{
    if (wb->defined_name_count >= wb->defined_name_cap) {
        size_t cap = wb->defined_name_cap ? wb->defined_name_cap * 2 : 4;
        lxlsx_reader_defined_name_entry *nb = (lxlsx_reader_defined_name_entry *)realloc(
            wb->defined_names, cap * sizeof(*nb));
        if (!nb) return -1;
        wb->defined_names    = nb;
        wb->defined_name_cap = cap;
    }
    wb->defined_names[wb->defined_name_count].name              = str_dup(name);
    wb->defined_names[wb->defined_name_count].formula           = str_dup(formula);
    wb->defined_names[wb->defined_name_count].scope_sheet_index = scope_sheet_index;
    wb->defined_names[wb->defined_name_count].hidden            = hidden;
    wb->defined_name_count++;
    return 0;
}

static void wb_on_start(void *ud, const char *name, const char **attrs)
{
    wb_ctx *c = (wb_ctx *)ud;

    if (lxlsx_reader_xml_name_eq(name, "workbookPr")) {
        const char *d1904 = lxlsx_reader_xml_attr(attrs, "date1904");
        if (d1904 && (strcmp(d1904, "1") == 0 || strcmp(d1904, "true") == 0)) {
            c->wb->uses_1904 = 1;
        }
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "sheets")) {
        c->in_sheets = 1;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "definedNames")) {
        c->in_defined_names = 1;
        return;
    }
    if (c->in_defined_names && lxlsx_reader_xml_name_eq(name, "definedName")) {
        const char *nm  = lxlsx_reader_xml_attr(attrs, "name");
        const char *lsi = lxlsx_reader_xml_attr(attrs, "localSheetId");
        const char *hd  = lxlsx_reader_xml_attr(attrs, "hidden");
        c->in_defined_name = 1;
        free(c->dn_name);
        c->dn_name = nm ? strdup(nm) : NULL;
        c->dn_local_sheet = lsi ? (int)strtol(lsi, NULL, 10) : -1;
        c->dn_hidden = (hd && (strcmp(hd, "1") == 0 || strcmp(hd, "true") == 0)) ? 1 : 0;
        c->dn_text_len = 0;
        if (c->dn_text) c->dn_text[0] = 0;
        return;
    }
    if (c->in_sheets && lxlsx_reader_xml_name_eq(name, "sheet")) {
        const char *sn    = lxlsx_reader_xml_attr(attrs, "name");
        const char *rid   = lxlsx_reader_xml_attr(attrs, "id");
        const char *state = lxlsx_reader_xml_attr(attrs, "state");
        lxlsx_reader_sheet_visibility vis = LXLSX_READER_SHEET_VISIBLE;
        /* attr name in workbook.xml is "r:id" — namespace-stripping needed.
         * Walk attrs explicitly to find any id=*. */
        if (!rid) {
            const char **a = attrs;
            while (a && *a) {
                if (lxlsx_reader_xml_name_eq(*a, "id")) { rid = *(a + 1); break; }
                a += 2;
            }
        }
        if (state) {
            if (strcmp(state, "hidden") == 0)           vis = LXLSX_READER_SHEET_HIDDEN;
            else if (strcmp(state, "veryHidden") == 0)  vis = LXLSX_READER_SHEET_VERY_HIDDEN;
        }
        if (sn && rid) sheet_array_push(c->wb, sn, rid, vis);
    }
}

static void wb_on_text(void *ud, const char *text, int len)
{
    wb_ctx *c = (wb_ctx *)ud;
    if (!c->in_defined_name || len <= 0) return;
    {
        size_t need = c->dn_text_len + (size_t)len + 1;
        if (need > c->dn_text_cap) {
            size_t nc = c->dn_text_cap ? c->dn_text_cap : 64;
            char *nb;
            while (nc < need) nc *= 2;
            nb = (char *)realloc(c->dn_text, nc);
            if (!nb) return;
            c->dn_text = nb;
            c->dn_text_cap = nc;
        }
        memcpy(c->dn_text + c->dn_text_len, text, (size_t)len);
        c->dn_text_len += (size_t)len;
        c->dn_text[c->dn_text_len] = 0;
    }
}

static void wb_on_end(void *ud, const char *name)
{
    wb_ctx *c = (wb_ctx *)ud;
    if (lxlsx_reader_xml_name_eq(name, "sheets")) c->in_sheets = 0;
    if (lxlsx_reader_xml_name_eq(name, "definedNames")) c->in_defined_names = 0;
    if (c->in_defined_name && lxlsx_reader_xml_name_eq(name, "definedName")) {
        if (c->dn_name) {
            dn_array_push(c->wb, c->dn_name,
                          c->dn_text_len > 0 ? c->dn_text : "",
                          c->dn_local_sheet, c->dn_hidden);
        }
        c->in_defined_name = 0;
        free(c->dn_name); c->dn_name = NULL;
        c->dn_text_len = 0;
        c->dn_local_sheet = -1;
        c->dn_hidden = 0;
    }
}

static lxlsx_reader_error parse_workbook_xml(lxlsx_reader_workbook *wb)
{
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    wb_ctx        ctx;
    lxlsx_reader_error     rc;

    zf = lxlsx_reader_zip_open_entry(wb->zip, wb->workbook_path);
    if (!zf) return LXLSX_READER_ERROR_FILE_CORRUPTED;

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) { lxlsx_reader_zip_close_entry(zf); return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED; }

    memset(&ctx, 0, sizeof(ctx));
    ctx.wb = wb;
    ctx.dn_local_sheet = -1;
    lxlsx_reader_xml_pump_set_handlers(pump, wb_on_start, wb_on_end, wb_on_text, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);

    free(ctx.dn_name);
    free(ctx.dn_text);
    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    return rc;
}

/* ------------------------------------------------------------------------- */
/* xl/_rels/workbook.xml.rels                                                */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxlsx_reader_workbook *wb;
} rel_ctx;

static void rel_on_start(void *ud, const char *name, const char **attrs)
{
    rel_ctx *c = (rel_ctx *)ud;
    const char *id, *type, *target;
    if (!lxlsx_reader_xml_name_eq(name, "Relationship")) return;

    id     = lxlsx_reader_xml_attr(attrs, "Id");
    type   = lxlsx_reader_xml_attr(attrs, "Type");
    target = lxlsx_reader_xml_attr(attrs, "Target");
    if (!id || !type || !target) return;

    if (strcmp(type, REL_TYPE_WORKSHEET) == 0) {
        size_t i;
        for (i = 0; i < c->wb->sheet_count; i++) {
            if (c->wb->sheets[i].rel_id &&
                strcmp(c->wb->sheets[i].rel_id, id) == 0) {
                free(c->wb->sheets[i].target);
                c->wb->sheets[i].target = join_base(c->wb->base_path, target);
                break;
            }
        }
    } else if (strcmp(type, REL_TYPE_SHARED_STRINGS) == 0) {
        free(c->wb->sst_path);
        c->wb->sst_path = join_base(c->wb->base_path, target);
    } else if (strcmp(type, REL_TYPE_STYLES) == 0) {
        free(c->wb->styles_path);
        c->wb->styles_path = join_base(c->wb->base_path, target);
    }
}

static lxlsx_reader_error parse_workbook_rels(lxlsx_reader_workbook *wb)
{
    char         *rpath;
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    rel_ctx       ctx;
    lxlsx_reader_error     rc;

    rpath = rels_path_of(wb->workbook_path);
    if (!rpath) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    zf = lxlsx_reader_zip_open_entry(wb->zip, rpath);
    free(rpath);
    if (!zf) return LXLSX_READER_ERROR_FILE_CORRUPTED;

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) { lxlsx_reader_zip_close_entry(zf); return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED; }

    ctx.wb = wb;
    lxlsx_reader_xml_pump_set_handlers(pump, rel_on_start, NULL, NULL, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);

    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    return rc;
}

/* ------------------------------------------------------------------------- */
/* Public open/close                                                         */
/* ------------------------------------------------------------------------- */

static lxlsx_reader_error workbook_finalize_open(lxlsx_reader_workbook *wb)
{
    lxlsx_reader_error rc;

    rc = find_workbook_path(wb->zip, &wb->workbook_path);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    wb->base_path = base_path_of(wb->workbook_path);
    if (!wb->base_path) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    rc = parse_workbook_xml(wb);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    rc = parse_workbook_rels(wb);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    rc = lxlsx_reader_styles_load(wb->zip, wb->styles_path, &wb->styles);
    if (rc != LXLSX_READER_NO_ERROR) return rc;

    if (wb->sst_path) {
        rc = lxlsx_reader_sst_open(wb->zip, wb->sst_path, wb->opts.sst_mode, &wb->sst);
        if (rc != LXLSX_READER_NO_ERROR) return rc;
    }

    return LXLSX_READER_NO_ERROR;
}

lxlsx_reader_error lxlsx_reader_workbook_open_ex(const char            *filename,
                               const lxlsx_reader_open_options *opts,
                               lxlsx_reader_workbook          **out)
{
    lxlsx_reader_workbook *wb;
    lxlsx_reader_error     rc;

    if (!filename || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;

    wb = (lxlsx_reader_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    if (opts) wb->opts = *opts;

    wb->zip = lxlsx_reader_zip_open_path(filename);
    if (!wb->zip) { free(wb); return LXLSX_READER_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXLSX_READER_NO_ERROR;
}

lxlsx_reader_error lxlsx_reader_workbook_open(const char *filename, lxlsx_reader_workbook **out)
{
    lxlsx_reader_open_options opts = {0};
    opts.sst_mode = LXLSX_READER_SST_MODE_FULL;
    return lxlsx_reader_workbook_open_ex(filename, &opts, out);
}

lxlsx_reader_error lxlsx_reader_workbook_open_fd(int fd, lxlsx_reader_workbook **out)
{
    lxlsx_reader_workbook *wb;
    lxlsx_reader_open_options opts = {0};
    lxlsx_reader_error     rc;

    if (!out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    wb = (lxlsx_reader_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    wb->opts = opts;

    wb->zip = lxlsx_reader_zip_open_fd(fd);
    if (!wb->zip) { free(wb); return LXLSX_READER_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXLSX_READER_NO_ERROR;
}

lxlsx_reader_error lxlsx_reader_workbook_open_memory(const void *data, size_t len, lxlsx_reader_workbook **out)
{
    lxlsx_reader_workbook *wb;
    lxlsx_reader_open_options opts = {0};
    lxlsx_reader_error     rc;

    if (!data || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    wb = (lxlsx_reader_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    wb->opts = opts;

    wb->zip = lxlsx_reader_zip_open_memory(data, len);
    if (!wb->zip) { free(wb); return LXLSX_READER_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXLSX_READER_NO_ERROR) {
        lxlsx_reader_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXLSX_READER_NO_ERROR;
}

void lxlsx_reader_workbook_close(lxlsx_reader_workbook *wb)
{
    size_t i;
    if (!wb) return;
    for (i = 0; i < wb->sheet_count; i++) {
        free(wb->sheets[i].name);
        free(wb->sheets[i].rel_id);
        free(wb->sheets[i].target);
    }
    for (i = 0; i < wb->defined_name_count; i++) {
        free(wb->defined_names[i].name);
        free(wb->defined_names[i].formula);
    }
    free(wb->defined_names);
    free(wb->sheets);
    free(wb->workbook_path);
    free(wb->base_path);
    free(wb->sst_path);
    free(wb->styles_path);
    if (wb->sst)    lxlsx_reader_sst_close(wb->sst);
    if (wb->styles) lxlsx_reader_styles_free(wb->styles);
    if (wb->zip)    lxlsx_reader_zip_close(wb->zip);
    free(wb);
}

/* --- accessors ----------------------------------------------------------- */

size_t lxlsx_reader_workbook_sheet_count(const lxlsx_reader_workbook *wb)
{
    return wb ? wb->sheet_count : 0;
}

const char *lxlsx_reader_workbook_sheet_name(const lxlsx_reader_workbook *wb, size_t index)
{
    if (!wb || index >= wb->sheet_count) return NULL;
    return wb->sheets[index].name;
}

lxlsx_reader_sheet_visibility lxlsx_reader_workbook_sheet_visibility(const lxlsx_reader_workbook *wb, size_t index)
{
    if (!wb || index >= wb->sheet_count) return LXLSX_READER_SHEET_VISIBLE;
    return (lxlsx_reader_sheet_visibility)wb->sheets[index].visibility;
}

size_t lxlsx_reader_workbook_defined_name_count(const lxlsx_reader_workbook *wb)
{
    return wb ? wb->defined_name_count : 0;
}

int lxlsx_reader_workbook_defined_name_get(const lxlsx_reader_workbook *wb, size_t idx,
                                  lxlsx_reader_defined_name *out)
{
    const lxlsx_reader_defined_name_entry *e;
    if (!wb || !out || idx >= wb->defined_name_count) return 0;
    e = &wb->defined_names[idx];
    out->name              = e->name;
    out->formula           = e->formula;
    out->scope_sheet_index = e->scope_sheet_index;
    out->hidden            = e->hidden;
    return 1;
}

int lxlsx_reader_workbook_uses_1904_dates(const lxlsx_reader_workbook *wb)
{
    return wb ? wb->uses_1904 : 0;
}

const lxlsx_reader_styles *lxlsx_reader_workbook_get_styles(const lxlsx_reader_workbook *wb)
{
    return wb ? wb->styles : NULL;
}

lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_name(lxlsx_reader_workbook *wb, const char *name,
                                             uint32_t flags, lxlsx_reader_worksheet **out)
{
    size_t i;
    if (!wb || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    *out = NULL;

    /* NULL name selects the first sheet, matching libxlsxio behaviour. */
    if (!name) {
        if (wb->sheet_count == 0 || !wb->sheets[0].target)
            return LXLSX_READER_ERROR_SHEET_NOT_FOUND;
        return lxlsx_reader_worksheet_open_internal(wb, wb->sheets[0].target, flags, out);
    }
    for (i = 0; i < wb->sheet_count; i++) {
        if (wb->sheets[i].name && strcmp(wb->sheets[i].name, name) == 0) {
            if (!wb->sheets[i].target) return LXLSX_READER_ERROR_SHEET_NOT_FOUND;
            return lxlsx_reader_worksheet_open_internal(wb, wb->sheets[i].target, flags, out);
        }
    }
    return LXLSX_READER_ERROR_SHEET_NOT_FOUND;
}

lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_index(lxlsx_reader_workbook *wb, size_t index,
                                              uint32_t flags, lxlsx_reader_worksheet **out)
{
    if (!wb || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    *out = NULL;
    if (index >= wb->sheet_count) return LXLSX_READER_ERROR_SHEET_NOT_FOUND;
    if (!wb->sheets[index].target) return LXLSX_READER_ERROR_SHEET_NOT_FOUND;
    return lxlsx_reader_worksheet_open_internal(wb, wb->sheets[index].target, flags, out);
}
