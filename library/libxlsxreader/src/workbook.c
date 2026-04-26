#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "lxr_internal.h"
#include "lxr_xml_pump.h"

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

/* Normalize a workbook-relative path to a zip-absolute path.
 * Examples:
 *   base="xl/", target="worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
 *   base="xl/", target="/xl/sharedStrings.xml" -> "xl/sharedStrings.xml"
 */
static char *join_base(const char *base, const char *target)
{
    size_t bl, tl;
    char  *out;
    if (!target) return NULL;
    if (target[0] == '/') return str_dup(target + 1);
    bl = base ? strlen(base) : 0;
    tl = strlen(target);
    out = (char *)malloc(bl + tl + 1);
    if (!out) return NULL;
    if (bl) memcpy(out, base, bl);
    memcpy(out + bl, target, tl);
    out[bl + tl] = 0;
    return out;
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

static int sheet_array_push(lxr_workbook *wb, const char *name, const char *rid)
{
    if (wb->sheet_count >= wb->sheet_cap) {
        size_t cap = wb->sheet_cap ? wb->sheet_cap * 2 : 8;
        lxr_sheet_info *nb = (lxr_sheet_info *)realloc(
            wb->sheets, cap * sizeof(*nb));
        if (!nb) return -1;
        wb->sheets    = nb;
        wb->sheet_cap = cap;
    }
    wb->sheets[wb->sheet_count].name   = str_dup(name);
    wb->sheets[wb->sheet_count].rel_id = str_dup(rid);
    wb->sheets[wb->sheet_count].target = NULL;
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
    if (!lxr_xml_name_eq(name, "Override")) return;
    {
        const char *part_name = lxr_xml_attr(attrs, "PartName");
        const char *content   = lxr_xml_attr(attrs, "ContentType");
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

static lxr_error find_workbook_path(lxr_zip *zip, char **out_path)
{
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    ct_ctx        ctx = {NULL};
    lxr_error     rc;

    zf = lxr_zip_open_entry(zip, "[Content_Types].xml");
    if (!zf) return LXR_ERROR_FILE_NOT_XLSX;

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxr_zip_close_entry(zf);
        return LXR_ERROR_MEMORY_MALLOC_FAILED;
    }

    lxr_xml_pump_set_handlers(pump, ct_on_start, NULL, NULL, &ctx);
    rc = lxr_xml_pump_run(pump);

    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);

    if (rc != LXR_NO_ERROR) {
        free(ctx.workbook_path);
        return rc;
    }
    if (!ctx.workbook_path) return LXR_ERROR_FILE_NOT_XLSX;
    *out_path = ctx.workbook_path;
    return LXR_NO_ERROR;
}

/* ------------------------------------------------------------------------- */
/* xl/workbook.xml                                                           */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxr_workbook *wb;
    int in_sheets;
} wb_ctx;

static void wb_on_start(void *ud, const char *name, const char **attrs)
{
    wb_ctx *c = (wb_ctx *)ud;

    if (lxr_xml_name_eq(name, "workbookPr")) {
        const char *d1904 = lxr_xml_attr(attrs, "date1904");
        if (d1904 && (strcmp(d1904, "1") == 0 || strcmp(d1904, "true") == 0)) {
            c->wb->uses_1904 = 1;
        }
        return;
    }
    if (lxr_xml_name_eq(name, "sheets")) {
        c->in_sheets = 1;
        return;
    }
    if (c->in_sheets && lxr_xml_name_eq(name, "sheet")) {
        const char *sn  = lxr_xml_attr(attrs, "name");
        const char *rid = lxr_xml_attr(attrs, "id");
        /* attr name in workbook.xml is "r:id" — namespace-stripping needed.
         * Walk attrs explicitly to find any id=*. */
        if (!rid) {
            const char **a = attrs;
            while (a && *a) {
                if (lxr_xml_name_eq(*a, "id")) { rid = *(a + 1); break; }
                a += 2;
            }
        }
        if (sn && rid) sheet_array_push(c->wb, sn, rid);
    }
}

static void wb_on_end(void *ud, const char *name)
{
    wb_ctx *c = (wb_ctx *)ud;
    if (lxr_xml_name_eq(name, "sheets")) c->in_sheets = 0;
}

static lxr_error parse_workbook_xml(lxr_workbook *wb)
{
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    wb_ctx        ctx;
    lxr_error     rc;

    zf = lxr_zip_open_entry(wb->zip, wb->workbook_path);
    if (!zf) return LXR_ERROR_FILE_CORRUPTED;

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }

    ctx.wb = wb;
    ctx.in_sheets = 0;
    lxr_xml_pump_set_handlers(pump, wb_on_start, wb_on_end, NULL, &ctx);
    rc = lxr_xml_pump_run(pump);

    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* ------------------------------------------------------------------------- */
/* xl/_rels/workbook.xml.rels                                                */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxr_workbook *wb;
} rel_ctx;

static void rel_on_start(void *ud, const char *name, const char **attrs)
{
    rel_ctx *c = (rel_ctx *)ud;
    const char *id, *type, *target;
    if (!lxr_xml_name_eq(name, "Relationship")) return;

    id     = lxr_xml_attr(attrs, "Id");
    type   = lxr_xml_attr(attrs, "Type");
    target = lxr_xml_attr(attrs, "Target");
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

static lxr_error parse_workbook_rels(lxr_workbook *wb)
{
    char         *rpath;
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    rel_ctx       ctx;
    lxr_error     rc;

    rpath = rels_path_of(wb->workbook_path);
    if (!rpath) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    zf = lxr_zip_open_entry(wb->zip, rpath);
    free(rpath);
    if (!zf) return LXR_ERROR_FILE_CORRUPTED;

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) { lxr_zip_close_entry(zf); return LXR_ERROR_MEMORY_MALLOC_FAILED; }

    ctx.wb = wb;
    lxr_xml_pump_set_handlers(pump, rel_on_start, NULL, NULL, &ctx);
    rc = lxr_xml_pump_run(pump);

    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);
    return rc;
}

/* ------------------------------------------------------------------------- */
/* Public open/close                                                         */
/* ------------------------------------------------------------------------- */

static lxr_error workbook_finalize_open(lxr_workbook *wb)
{
    lxr_error rc;

    rc = find_workbook_path(wb->zip, &wb->workbook_path);
    if (rc != LXR_NO_ERROR) return rc;

    wb->base_path = base_path_of(wb->workbook_path);
    if (!wb->base_path) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    rc = parse_workbook_xml(wb);
    if (rc != LXR_NO_ERROR) return rc;

    rc = parse_workbook_rels(wb);
    if (rc != LXR_NO_ERROR) return rc;

    rc = lxr_styles_load(wb->zip, wb->styles_path, &wb->styles);
    if (rc != LXR_NO_ERROR) return rc;

    if (wb->sst_path) {
        rc = lxr_sst_open(wb->zip, wb->sst_path, wb->opts.sst_mode, &wb->sst);
        if (rc != LXR_NO_ERROR) return rc;
    }

    return LXR_NO_ERROR;
}

lxr_error lxr_workbook_open_ex(const char            *filename,
                               const lxr_open_options *opts,
                               lxr_workbook          **out)
{
    lxr_workbook *wb;
    lxr_error     rc;

    if (!filename || !out) return LXR_ERROR_NULL_PARAMETER;

    wb = (lxr_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXR_ERROR_MEMORY_MALLOC_FAILED;
    if (opts) wb->opts = *opts;

    wb->zip = lxr_zip_open_path(filename);
    if (!wb->zip) { free(wb); return LXR_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXR_NO_ERROR) {
        lxr_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXR_NO_ERROR;
}

lxr_error lxr_workbook_open(const char *filename, lxr_workbook **out)
{
    lxr_open_options opts = {0};
    opts.sst_mode = LXR_SST_MODE_FULL;
    return lxr_workbook_open_ex(filename, &opts, out);
}

lxr_error lxr_workbook_open_fd(int fd, lxr_workbook **out)
{
    lxr_workbook *wb;
    lxr_open_options opts = {0};
    lxr_error     rc;

    if (!out) return LXR_ERROR_NULL_PARAMETER;
    wb = (lxr_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXR_ERROR_MEMORY_MALLOC_FAILED;
    wb->opts = opts;

    wb->zip = lxr_zip_open_fd(fd);
    if (!wb->zip) { free(wb); return LXR_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXR_NO_ERROR) {
        lxr_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXR_NO_ERROR;
}

lxr_error lxr_workbook_open_memory(const void *data, size_t len, lxr_workbook **out)
{
    lxr_workbook *wb;
    lxr_open_options opts = {0};
    lxr_error     rc;

    if (!data || !out) return LXR_ERROR_NULL_PARAMETER;
    wb = (lxr_workbook *)calloc(1, sizeof(*wb));
    if (!wb) return LXR_ERROR_MEMORY_MALLOC_FAILED;
    wb->opts = opts;

    wb->zip = lxr_zip_open_memory(data, len);
    if (!wb->zip) { free(wb); return LXR_ERROR_FILE_OPEN_FAILED; }

    rc = workbook_finalize_open(wb);
    if (rc != LXR_NO_ERROR) {
        lxr_workbook_close(wb);
        *out = NULL;
        return rc;
    }
    *out = wb;
    return LXR_NO_ERROR;
}

void lxr_workbook_close(lxr_workbook *wb)
{
    size_t i;
    if (!wb) return;
    for (i = 0; i < wb->sheet_count; i++) {
        free(wb->sheets[i].name);
        free(wb->sheets[i].rel_id);
        free(wb->sheets[i].target);
    }
    free(wb->sheets);
    free(wb->workbook_path);
    free(wb->base_path);
    free(wb->sst_path);
    free(wb->styles_path);
    if (wb->sst)    lxr_sst_close(wb->sst);
    if (wb->styles) lxr_styles_free(wb->styles);
    if (wb->zip)    lxr_zip_close(wb->zip);
    free(wb);
}

/* --- accessors ----------------------------------------------------------- */

size_t lxr_workbook_sheet_count(const lxr_workbook *wb)
{
    return wb ? wb->sheet_count : 0;
}

const char *lxr_workbook_sheet_name(const lxr_workbook *wb, size_t index)
{
    if (!wb || index >= wb->sheet_count) return NULL;
    return wb->sheets[index].name;
}

int lxr_workbook_uses_1904_dates(const lxr_workbook *wb)
{
    return wb ? wb->uses_1904 : 0;
}

const lxr_styles *lxr_workbook_get_styles(const lxr_workbook *wb)
{
    return wb ? wb->styles : NULL;
}

lxr_error lxr_workbook_get_worksheet_by_name(lxr_workbook *wb, const char *name,
                                             uint32_t flags, lxr_worksheet **out)
{
    size_t i;
    if (!wb || !out) return LXR_ERROR_NULL_PARAMETER;
    *out = NULL;

    /* NULL name selects the first sheet, matching libxlsxio behaviour. */
    if (!name) {
        if (wb->sheet_count == 0 || !wb->sheets[0].target)
            return LXR_ERROR_SHEET_NOT_FOUND;
        return lxr_worksheet_open_internal(wb, wb->sheets[0].target, flags, out);
    }
    for (i = 0; i < wb->sheet_count; i++) {
        if (wb->sheets[i].name && strcmp(wb->sheets[i].name, name) == 0) {
            if (!wb->sheets[i].target) return LXR_ERROR_SHEET_NOT_FOUND;
            return lxr_worksheet_open_internal(wb, wb->sheets[i].target, flags, out);
        }
    }
    return LXR_ERROR_SHEET_NOT_FOUND;
}

lxr_error lxr_workbook_get_worksheet_by_index(lxr_workbook *wb, size_t index,
                                              uint32_t flags, lxr_worksheet **out)
{
    if (!wb || !out) return LXR_ERROR_NULL_PARAMETER;
    *out = NULL;
    if (index >= wb->sheet_count) return LXR_ERROR_SHEET_NOT_FOUND;
    if (!wb->sheets[index].target) return LXR_ERROR_SHEET_NOT_FOUND;
    return lxr_worksheet_open_internal(wb, wb->sheets[index].target, flags, out);
}
