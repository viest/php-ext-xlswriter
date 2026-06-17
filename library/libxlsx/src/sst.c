#include <stdlib.h>
#include <string.h>

#include "sst.h"
#include "xlsx_util.h"
#include "xml_pump.h"

/* Per-SI run accumulator — populated while scanning <r>...</r>. */
typedef struct {
    lxlsx_reader_sst_run *items;
    size_t       count;
    size_t       cap;
} lxlsx_reader_run_list;

struct lxlsx_reader_sst {
    lxlsx_reader_sst_mode mode;

    char  **strings;
    size_t  loaded_count;
    size_t  capacity;

    /* Parallel run table: one list per SST index. The outer pointer is
     * lazily allocated alongside strings so existing callers that don't use
     * runs pay nothing extra. */
    lxlsx_reader_run_list *runs;
    size_t        runs_capacity;

    lxlsx_reader_zip      *zip;
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    lxlsx_reader_error error;
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
    lxlsx_reader_sst_run   pending;
    /* Runs collected for the current <si>. */
    lxlsx_reader_run_list  cur_runs;
};

static void sst_fail(lxlsx_reader_sst *s, lxlsx_reader_error err)
{
    if (!s->error)
        s->error = err;
    if (s->pump)
        lxlsx_reader_xml_pump_suspend(s->pump);
}

/* --- buffer helpers ------------------------------------------------------ */

static int append_text(lxlsx_reader_sst *s, const char *src, size_t n)
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

static int append_run_text(lxlsx_reader_sst *s, const char *src, size_t n)
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

static void run_clear_pending(lxlsx_reader_sst *s)
{
    free(s->pending.text);
    free(s->pending.font_name);
    free(s->pending.color);
    memset(&s->pending, 0, sizeof(s->pending));
}

static int run_list_push(lxlsx_reader_run_list *l, lxlsx_reader_sst_run r)
{
    if (l->count >= l->cap) {
        size_t nc = l->cap ? l->cap * 2 : 4;
        lxlsx_reader_sst_run *nb = (lxlsx_reader_sst_run *)realloc(l->items, nc * sizeof(*nb));
        if (!nb) return -1;
        l->items = nb;
        l->cap   = nc;
    }
    l->items[l->count++] = r;
    return 0;
}

static int sst_runs_ensure(lxlsx_reader_sst *s, size_t at_least)
{
    if (at_least <= s->runs_capacity) return 0;
    {
        size_t nc = s->runs_capacity ? s->runs_capacity * 2 : 16;
        lxlsx_reader_run_list *nb;
        while (nc < at_least) nc *= 2;
        nb = (lxlsx_reader_run_list *)realloc(s->runs, nc * sizeof(*nb));
        if (!nb) return -1;
        memset(nb + s->runs_capacity, 0,
               (nc - s->runs_capacity) * sizeof(*nb));
        s->runs = nb;
        s->runs_capacity = nc;
    }
    return 0;
}

static int commit_string(lxlsx_reader_sst *s)
{
    char *str;
    if (s->loaded_count >= s->capacity) {
        size_t new_cap = s->capacity ? s->capacity * 2 : 16;
        char **ns = (char **)realloc(s->strings, new_cap * sizeof(*ns));
        if (!ns) return -1;
        s->strings  = ns;
        s->capacity = new_cap;
    }
    str = s->cur_buf
        ? lxlsx_reader_strndup(s->cur_buf, s->cur_len)
        : lxlsx_reader_strndup("", 0);
    if (!str) return -1;

    /* Migrate any collected runs into the parallel runs[] table. */
    if (s->cur_runs.count > 0) {
        if (sst_runs_ensure(s, s->loaded_count + 1) != 0) {
            free(str);
            return -1;
        }
        s->runs[s->loaded_count] = s->cur_runs;
        memset(&s->cur_runs, 0, sizeof(s->cur_runs));
    }

    s->strings[s->loaded_count] = str;
    s->loaded_count++;

    /* reset accumulator (keep the buffer for reuse) */
    s->cur_len = 0;
    if (s->cur_buf) s->cur_buf[0] = 0;
    return 0;
}

/* --- SAX handlers -------------------------------------------------------- */

static void on_start(void *ud, const char *name, const char **attrs)
{
    lxlsx_reader_sst *s = (lxlsx_reader_sst *)ud;

    if (s->error)
        return;

    if (s->skip_depth > 0) {
        s->skip_depth++;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "rPh") || lxlsx_reader_xml_name_eq(name, "phoneticPr")) {
        free(s->skip_tag);
        s->skip_tag   = strdup(name);
        if (!s->skip_tag) {
            sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
            return;
        }
        s->skip_depth = 1;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "si")) {
        s->in_si   = 1;
        s->cur_len = 0;
        if (s->cur_buf) s->cur_buf[0] = 0;
        memset(&s->cur_runs, 0, sizeof(s->cur_runs));
        return;
    }
    if (!s->in_si) return;

    /* Inside <si>: handle <r>, <t>, and rPr children. */
    if (s->in_r) {
        if (lxlsx_reader_xml_name_eq(name, "rPr")) {
            s->in_rpr = 1;
            return;
        }
        if (s->in_rpr) {
            const char *v;
            if (lxlsx_reader_xml_name_eq(name, "rFont")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "val"))) {
                    free(s->pending.font_name);
                    s->pending.font_name = strdup(v);
                    if (!s->pending.font_name) {
                        sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                        return;
                    }
                }
            } else if (lxlsx_reader_xml_name_eq(name, "sz")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "val"))) s->pending.font_size = strtod(v, NULL);
            } else if (lxlsx_reader_xml_name_eq(name, "b")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                s->pending.bold = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "i")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                s->pending.italic = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "strike")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                s->pending.strike = !v || strcmp(v, "0") != 0;
            } else if (lxlsx_reader_xml_name_eq(name, "u")) {
                v = lxlsx_reader_xml_attr(attrs, "val");
                if (!v || strcmp(v, "single") == 0)             s->pending.underline = 1;
                else if (strcmp(v, "double") == 0)              s->pending.underline = 2;
                else if (strcmp(v, "singleAccounting") == 0)    s->pending.underline = 3;
                else if (strcmp(v, "doubleAccounting") == 0)    s->pending.underline = 4;
            } else if (lxlsx_reader_xml_name_eq(name, "color")) {
                if ((v = lxlsx_reader_xml_attr(attrs, "rgb"))) {
                    free(s->pending.color);
                    s->pending.color = strdup(v);
                    if (!s->pending.color) {
                        sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
                        return;
                    }
                }
            }
            return;
        }
        if (lxlsx_reader_xml_name_eq(name, "t")) {
            s->in_run_t  = 1;
            s->in_t      = 1;
            s->run_text_len = 0;
            if (s->run_text_buf) s->run_text_buf[0] = 0;
        }
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "r")) {
        s->in_r = 1;
        run_clear_pending(s);
        s->run_text_len = 0;
        if (s->run_text_buf) s->run_text_buf[0] = 0;
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "t")) {
        s->in_t = 1;
    }
}

static void on_end(void *ud, const char *name)
{
    lxlsx_reader_sst *s = (lxlsx_reader_sst *)ud;

    if (s->error)
        return;

    if (s->skip_depth > 0) {
        s->skip_depth--;
        if (s->skip_depth == 0) {
            free(s->skip_tag);
            s->skip_tag = NULL;
        }
        return;
    }
    if (lxlsx_reader_xml_name_eq(name, "t")) {
        s->in_t = 0;
        if (s->in_run_t) s->in_run_t = 0;
    } else if (lxlsx_reader_xml_name_eq(name, "rPr")) {
        s->in_rpr = 0;
    } else if (lxlsx_reader_xml_name_eq(name, "r") && s->in_r) {
        /* Commit the run. */
        s->pending.text = s->run_text_len > 0 ? strdup(s->run_text_buf) : strdup("");
        if (!s->pending.text || run_list_push(&s->cur_runs, s->pending) != 0) {
            sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
            return;
        }
        memset(&s->pending, 0, sizeof(s->pending));
        s->in_r = 0;
        s->run_text_len = 0;
        if (s->run_text_buf) s->run_text_buf[0] = 0;
    } else if (lxlsx_reader_xml_name_eq(name, "si") && s->in_si) {
        if (commit_string(s) != 0) {
            sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
            return;
        }
        s->in_si = 0;
        if (s->mode == LXLSX_READER_SST_MODE_STREAMING) {
            lxlsx_reader_xml_pump_suspend(s->pump);
        }
    } else if (lxlsx_reader_xml_name_eq(name, "sst")) {
        s->eof = 1;
    }
}

static void on_text(void *ud, const char *text, int len)
{
    lxlsx_reader_sst *s = (lxlsx_reader_sst *)ud;
    if (s->error || s->skip_depth || len <= 0) return;
    if (s->in_t) {
        if (append_text(s, text, (size_t)len) != 0) {
            sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
            return;
        }
    }
    if (s->in_run_t) {
        if (append_run_text(s, text, (size_t)len) != 0) {
            sst_fail(s, LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED);
            return;
        }
    }
}

/* --- public -------------------------------------------------------------- */

lxlsx_reader_error lxlsx_reader_sst_open(lxlsx_reader_zip *zip, const char *path, lxlsx_reader_sst_mode mode, lxlsx_reader_sst **out)
{
    lxlsx_reader_sst *s;
    if (!zip || !path || !out) return LXLSX_READER_ERROR_NULL_PARAMETER;

    s = (lxlsx_reader_sst *)calloc(1, sizeof(*s));
    if (!s) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    s->mode = mode;
    s->zip  = zip;

    s->zf = lxlsx_reader_zip_open_entry(zip, path);
    if (!s->zf) { free(s); return LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND; }

    s->pump = lxlsx_reader_xml_pump_create_zip_file(s->zf);
    if (!s->pump) {
        lxlsx_reader_zip_close_entry(s->zf);
        free(s);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }
    lxlsx_reader_xml_pump_set_handlers(s->pump, on_start, on_end, on_text, s);

    if (mode == LXLSX_READER_SST_MODE_FULL) {
        lxlsx_reader_error rc = lxlsx_reader_xml_pump_run(s->pump);
        if (rc != LXLSX_READER_NO_ERROR) {
            lxlsx_reader_sst_close(s);
            return rc;
        }
        if (s->error != LXLSX_READER_NO_ERROR) {
            rc = s->error;
            lxlsx_reader_sst_close(s);
            return rc;
        }
    } else {
        /* For STREAMING we do not pre-parse; SAX will run on demand. */
    }

    *out = s;
    return LXLSX_READER_NO_ERROR;
}

const char *lxlsx_reader_sst_get(lxlsx_reader_sst *s, uint32_t index)
{
    if (!s) return NULL;
    if (s->error) return NULL;
    if (index < s->loaded_count) return s->strings[index];

    if (s->mode == LXLSX_READER_SST_MODE_FULL || s->eof) return NULL;

    /* STREAMING: drive the pump until enough strings are loaded or EOF. */
    while ((uint32_t)s->loaded_count <= index && !s->eof) {
        lxlsx_reader_error rc = lxlsx_reader_xml_pump_resume(s->pump);
        if (rc != LXLSX_READER_NO_ERROR) return NULL;
        if (s->error) return NULL;
        if (!lxlsx_reader_xml_pump_is_suspended(s->pump)) break;
    }
    if (index < s->loaded_count) return s->strings[index];
    return NULL;
}

size_t lxlsx_reader_sst_count(const lxlsx_reader_sst *s)
{
    if (!s) return 0;
    return s->loaded_count;
}

size_t lxlsx_reader_sst_loaded_count(const lxlsx_reader_sst *s)
{
    return s ? s->loaded_count : 0;
}

const lxlsx_reader_sst_run *lxlsx_reader_sst_get_runs(lxlsx_reader_sst *s, uint32_t index, size_t *out_count)
{
    if (out_count) *out_count = 0;
    if (!s) return NULL;
    if (s->error) return NULL;
    /* Drive the pump if needed (STREAMING mode) — same logic as lxlsx_reader_sst_get. */
    if (index >= s->loaded_count && s->mode == LXLSX_READER_SST_MODE_STREAMING && !s->eof) {
        while ((uint32_t)s->loaded_count <= index && !s->eof) {
            if (lxlsx_reader_xml_pump_resume(s->pump) != LXLSX_READER_NO_ERROR) return NULL;
            if (s->error) return NULL;
            if (!lxlsx_reader_xml_pump_is_suspended(s->pump)) break;
        }
    }
    if (index >= s->loaded_count) return NULL;
    if (!s->runs || index >= s->runs_capacity) return NULL;
    if (out_count) *out_count = s->runs[index].count;
    return s->runs[index].items;
}

void lxlsx_reader_sst_close(lxlsx_reader_sst *s)
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
    if (s->pump) lxlsx_reader_xml_pump_destroy(s->pump);
    if (s->zf)   lxlsx_reader_zip_close_entry(s->zf);
    free(s);
}
