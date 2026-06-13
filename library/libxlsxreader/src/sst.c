#include <stdlib.h>
#include <string.h>

#include "lxr_sst.h"
#include "lxr_xml_pump.h"

/* Per-SI run accumulator — populated while scanning <r>...</r>. */
typedef struct {
    lxr_sst_run *items;
    size_t       count;
    size_t       cap;
} lxr_run_list;

struct lxr_sst {
    lxr_sst_mode mode;

    char  **strings;
    size_t  loaded_count;
    size_t  capacity;

    /* Parallel run table: one list per SST index. The outer pointer is
     * lazily allocated alongside strings so existing callers that don't use
     * runs pay nothing extra. */
    lxr_run_list *runs;
    size_t        runs_capacity;

    lxr_zip      *zip;
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    int           eof;

    /* SAX parse state */
    int           in_si;
    int           in_t;
    int           skip_depth;       /* >0 while inside <rPh>/<phoneticPr> */
    char         *skip_tag;          /* tag name we are skipping at depth 1 */
    char         *cur_buf;
    size_t        cur_len;
    size_t        cur_cap;

    /* Run-tracking state. */
    int           in_r;           /* inside <r> */
    int           in_rpr;         /* inside <rPr> */
    int           in_run_t;       /* inside <r>...<t>...</t> */
    /* Buffer for the run's text. */
    char         *run_text_buf;
    size_t        run_text_len;
    size_t        run_text_cap;
    /* Pending properties for the run being assembled. */
    lxr_sst_run   pending;
    /* Runs collected for the current <si>. */
    lxr_run_list  cur_runs;
};

/* --- buffer helpers ------------------------------------------------------ */

static int append_text(lxr_sst *s, const char *src, size_t n)
{
    size_t need = s->cur_len + n + 1;
    if (need > s->cur_cap) {
        size_t new_cap = s->cur_cap ? s->cur_cap : 64;
        char  *nb;
        while (new_cap < need) new_cap *= 2;
        nb = (char *)realloc(s->cur_buf, new_cap);
        if (!nb) return -1;
        s->cur_buf = nb;
        s->cur_cap = new_cap;
    }
    memcpy(s->cur_buf + s->cur_len, src, n);
    s->cur_len += n;
    s->cur_buf[s->cur_len] = 0;
    return 0;
}

static int append_run_text(lxr_sst *s, const char *src, size_t n)
{
    size_t need = s->run_text_len + n + 1;
    if (need > s->run_text_cap) {
        size_t new_cap = s->run_text_cap ? s->run_text_cap : 32;
        char  *nb;
        while (new_cap < need) new_cap *= 2;
        nb = (char *)realloc(s->run_text_buf, new_cap);
        if (!nb) return -1;
        s->run_text_buf = nb;
        s->run_text_cap = new_cap;
    }
    memcpy(s->run_text_buf + s->run_text_len, src, n);
    s->run_text_len += n;
    s->run_text_buf[s->run_text_len] = 0;
    return 0;
}

static void run_clear_pending(lxr_sst *s)
{
    free(s->pending.text);
    free(s->pending.font_name);
    free(s->pending.color);
    memset(&s->pending, 0, sizeof(s->pending));
}

static int run_list_push(lxr_run_list *l, lxr_sst_run r)
{
    if (l->count >= l->cap) {
        size_t nc = l->cap ? l->cap * 2 : 4;
        lxr_sst_run *nb = (lxr_sst_run *)realloc(l->items, nc * sizeof(*nb));
        if (!nb) return -1;
        l->items = nb;
        l->cap   = nc;
    }
    l->items[l->count++] = r;
    return 0;
}

static int sst_runs_ensure(lxr_sst *s, size_t at_least)
{
    if (at_least <= s->runs_capacity) return 0;
    {
        size_t nc = s->runs_capacity ? s->runs_capacity * 2 : 16;
        lxr_run_list *nb;
        while (nc < at_least) nc *= 2;
        nb = (lxr_run_list *)realloc(s->runs, nc * sizeof(*nb));
        if (!nb) return -1;
        memset(nb + s->runs_capacity, 0,
               (nc - s->runs_capacity) * sizeof(*nb));
        s->runs = nb;
        s->runs_capacity = nc;
    }
    return 0;
}

static int commit_string(lxr_sst *s)
{
    char *str;
    if (s->loaded_count >= s->capacity) {
        size_t new_cap = s->capacity ? s->capacity * 2 : 16;
        char **ns = (char **)realloc(s->strings, new_cap * sizeof(*ns));
        if (!ns) return -1;
        s->strings  = ns;
        s->capacity = new_cap;
    }
    str = s->cur_buf ? strdup(s->cur_buf) : strdup("");
    if (!str) return -1;
    s->strings[s->loaded_count] = str;

    /* Migrate any collected runs into the parallel runs[] table. */
    if (s->cur_runs.count > 0) {
        if (sst_runs_ensure(s, s->loaded_count + 1) == 0) {
            s->runs[s->loaded_count] = s->cur_runs;
        } else {
            /* OOM — drop the runs but keep the string. */
            size_t i;
            for (i = 0; i < s->cur_runs.count; i++) {
                free(s->cur_runs.items[i].text);
                free(s->cur_runs.items[i].font_name);
                free(s->cur_runs.items[i].color);
            }
            free(s->cur_runs.items);
        }
        memset(&s->cur_runs, 0, sizeof(s->cur_runs));
    }

    s->loaded_count++;

    /* reset accumulator (keep the buffer for reuse) */
    s->cur_len = 0;
    if (s->cur_buf) s->cur_buf[0] = 0;
    return 0;
}

/* --- SAX handlers -------------------------------------------------------- */

static void on_start(void *ud, const char *name, const char **attrs)
{
    lxr_sst *s = (lxr_sst *)ud;

    if (s->skip_depth > 0) {
        s->skip_depth++;
        return;
    }
    if (lxr_xml_name_eq(name, "rPh") || lxr_xml_name_eq(name, "phoneticPr")) {
        free(s->skip_tag);
        s->skip_tag   = strdup(name);
        s->skip_depth = 1;
        return;
    }
    if (lxr_xml_name_eq(name, "si")) {
        s->in_si   = 1;
        s->cur_len = 0;
        if (s->cur_buf) s->cur_buf[0] = 0;
        memset(&s->cur_runs, 0, sizeof(s->cur_runs));
        return;
    }
    if (!s->in_si) return;

    /* Inside <si>: handle <r>, <t>, and rPr children. */
    if (s->in_r) {
        if (lxr_xml_name_eq(name, "rPr")) {
            s->in_rpr = 1;
            return;
        }
        if (s->in_rpr) {
            const char *v;
            if (lxr_xml_name_eq(name, "rFont")) {
                if ((v = lxr_xml_attr(attrs, "val"))) {
                    free(s->pending.font_name);
                    s->pending.font_name = strdup(v);
                }
            } else if (lxr_xml_name_eq(name, "sz")) {
                if ((v = lxr_xml_attr(attrs, "val"))) s->pending.font_size = strtod(v, NULL);
            } else if (lxr_xml_name_eq(name, "b")) {
                v = lxr_xml_attr(attrs, "val");
                s->pending.bold = !v || strcmp(v, "0") != 0;
            } else if (lxr_xml_name_eq(name, "i")) {
                v = lxr_xml_attr(attrs, "val");
                s->pending.italic = !v || strcmp(v, "0") != 0;
            } else if (lxr_xml_name_eq(name, "strike")) {
                v = lxr_xml_attr(attrs, "val");
                s->pending.strike = !v || strcmp(v, "0") != 0;
            } else if (lxr_xml_name_eq(name, "u")) {
                v = lxr_xml_attr(attrs, "val");
                if (!v || strcmp(v, "single") == 0)             s->pending.underline = 1;
                else if (strcmp(v, "double") == 0)              s->pending.underline = 2;
                else if (strcmp(v, "singleAccounting") == 0)    s->pending.underline = 3;
                else if (strcmp(v, "doubleAccounting") == 0)    s->pending.underline = 4;
            } else if (lxr_xml_name_eq(name, "color")) {
                if ((v = lxr_xml_attr(attrs, "rgb"))) {
                    free(s->pending.color);
                    s->pending.color = strdup(v);
                }
            }
            return;
        }
        if (lxr_xml_name_eq(name, "t")) {
            s->in_run_t  = 1;
            s->in_t      = 1;
            s->run_text_len = 0;
            if (s->run_text_buf) s->run_text_buf[0] = 0;
        }
        return;
    }
    if (lxr_xml_name_eq(name, "r")) {
        s->in_r = 1;
        run_clear_pending(s);
        s->run_text_len = 0;
        if (s->run_text_buf) s->run_text_buf[0] = 0;
        return;
    }
    if (lxr_xml_name_eq(name, "t")) {
        s->in_t = 1;
    }
}

static void on_end(void *ud, const char *name)
{
    lxr_sst *s = (lxr_sst *)ud;

    if (s->skip_depth > 0) {
        s->skip_depth--;
        if (s->skip_depth == 0) {
            free(s->skip_tag);
            s->skip_tag = NULL;
        }
        return;
    }
    if (lxr_xml_name_eq(name, "t")) {
        s->in_t = 0;
        if (s->in_run_t) s->in_run_t = 0;
    } else if (lxr_xml_name_eq(name, "rPr")) {
        s->in_rpr = 0;
    } else if (lxr_xml_name_eq(name, "r") && s->in_r) {
        /* Commit the run. */
        s->pending.text = s->run_text_len > 0 ? strdup(s->run_text_buf) : strdup("");
        run_list_push(&s->cur_runs, s->pending);
        memset(&s->pending, 0, sizeof(s->pending));
        s->in_r = 0;
        s->run_text_len = 0;
        if (s->run_text_buf) s->run_text_buf[0] = 0;
    } else if (lxr_xml_name_eq(name, "si") && s->in_si) {
        commit_string(s);
        s->in_si = 0;
        if (s->mode == LXR_SST_MODE_STREAMING) {
            lxr_xml_pump_suspend(s->pump);
        }
    } else if (lxr_xml_name_eq(name, "sst")) {
        s->eof = 1;
    }
}

static void on_text(void *ud, const char *text, int len)
{
    lxr_sst *s = (lxr_sst *)ud;
    if (s->skip_depth || len <= 0) return;
    if (s->in_t) {
        append_text(s, text, (size_t)len);
    }
    if (s->in_run_t) {
        append_run_text(s, text, (size_t)len);
    }
}

/* --- public -------------------------------------------------------------- */

lxr_error lxr_sst_open(lxr_zip *zip, const char *path, lxr_sst_mode mode, lxr_sst **out)
{
    lxr_sst *s;
    if (!zip || !path || !out) return LXR_ERROR_NULL_PARAMETER;

    s = (lxr_sst *)calloc(1, sizeof(*s));
    if (!s) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    s->mode = mode;
    s->zip  = zip;

    s->zf = lxr_zip_open_entry(zip, path);
    if (!s->zf) { free(s); return LXR_ERROR_ZIP_ENTRY_NOT_FOUND; }

    s->pump = lxr_xml_pump_create_zip_file(s->zf);
    if (!s->pump) {
        lxr_zip_close_entry(s->zf);
        free(s);
        return LXR_ERROR_MEMORY_MALLOC_FAILED;
    }
    lxr_xml_pump_set_handlers(s->pump, on_start, on_end, on_text, s);

    if (mode == LXR_SST_MODE_FULL) {
        lxr_error rc = lxr_xml_pump_run(s->pump);
        if (rc != LXR_NO_ERROR) {
            lxr_sst_close(s);
            return rc;
        }
    } else {
        /* For STREAMING we do not pre-parse; SAX will run on demand. */
    }

    *out = s;
    return LXR_NO_ERROR;
}

const char *lxr_sst_get(lxr_sst *s, uint32_t index)
{
    if (!s) return NULL;
    if (index < s->loaded_count) return s->strings[index];

    if (s->mode == LXR_SST_MODE_FULL || s->eof) return NULL;

    /* STREAMING: drive the pump until enough strings are loaded or EOF. */
    while ((uint32_t)s->loaded_count <= index && !s->eof) {
        lxr_error rc = lxr_xml_pump_resume(s->pump);
        if (rc != LXR_NO_ERROR) return NULL;
        if (!lxr_xml_pump_is_suspended(s->pump)) break;
    }
    if (index < s->loaded_count) return s->strings[index];
    return NULL;
}

size_t lxr_sst_count(const lxr_sst *s)
{
    if (!s) return 0;
    if (s->mode == LXR_SST_MODE_FULL) return s->loaded_count;
    /* STREAMING: only known precisely after EOF. Return loaded_count for
       partial knowledge; callers can distinguish via lxr_sst_loaded_count. */
    return s->loaded_count;
}

size_t lxr_sst_loaded_count(const lxr_sst *s)
{
    return s ? s->loaded_count : 0;
}

const lxr_sst_run *lxr_sst_get_runs(lxr_sst *s, uint32_t index, size_t *out_count)
{
    if (out_count) *out_count = 0;
    if (!s) return NULL;
    /* Drive the pump if needed (STREAMING mode) — same logic as lxr_sst_get. */
    if (index >= s->loaded_count && s->mode == LXR_SST_MODE_STREAMING && !s->eof) {
        while ((uint32_t)s->loaded_count <= index && !s->eof) {
            if (lxr_xml_pump_resume(s->pump) != LXR_NO_ERROR) return NULL;
            if (!lxr_xml_pump_is_suspended(s->pump)) break;
        }
    }
    if (index >= s->loaded_count) return NULL;
    if (!s->runs || index >= s->runs_capacity) return NULL;
    if (out_count) *out_count = s->runs[index].count;
    return s->runs[index].items;
}

void lxr_sst_close(lxr_sst *s)
{
    size_t i, j;
    if (!s) return;
    for (i = 0; i < s->loaded_count; i++) free(s->strings[i]);
    free(s->strings);
    if (s->runs) {
        for (i = 0; i < s->runs_capacity; i++) {
            for (j = 0; j < s->runs[i].count; j++) {
                free(s->runs[i].items[j].text);
                free(s->runs[i].items[j].font_name);
                free(s->runs[i].items[j].color);
            }
            free(s->runs[i].items);
        }
        free(s->runs);
    }
    /* Free any in-progress run state (e.g. STREAMING mode never finished). */
    for (i = 0; i < s->cur_runs.count; i++) {
        free(s->cur_runs.items[i].text);
        free(s->cur_runs.items[i].font_name);
        free(s->cur_runs.items[i].color);
    }
    free(s->cur_runs.items);
    free(s->pending.text);
    free(s->pending.font_name);
    free(s->pending.color);
    free(s->run_text_buf);
    free(s->cur_buf);
    free(s->skip_tag);
    if (s->pump) lxr_xml_pump_destroy(s->pump);
    if (s->zf)   lxr_zip_close_entry(s->zf);
    free(s);
}
