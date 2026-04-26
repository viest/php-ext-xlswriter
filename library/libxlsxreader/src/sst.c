#include <stdlib.h>
#include <string.h>

#include "lxr_sst.h"
#include "lxr_xml_pump.h"

struct lxr_sst {
    lxr_sst_mode mode;

    char  **strings;
    size_t  loaded_count;
    size_t  capacity;

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
    s->strings[s->loaded_count++] = str;

    /* reset accumulator (keep the buffer for reuse) */
    s->cur_len = 0;
    if (s->cur_buf) s->cur_buf[0] = 0;
    return 0;
}

/* --- SAX handlers -------------------------------------------------------- */

static void on_start(void *ud, const char *name, const char **attrs)
{
    lxr_sst *s = (lxr_sst *)ud;
    (void)attrs;

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
    } else if (lxr_xml_name_eq(name, "t") && s->in_si) {
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
    if (s->skip_depth == 0 && s->in_t && len > 0) {
        append_text(s, text, (size_t)len);
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

void lxr_sst_close(lxr_sst *s)
{
    size_t i;
    if (!s) return;
    for (i = 0; i < s->loaded_count; i++) free(s->strings[i]);
    free(s->strings);
    free(s->cur_buf);
    free(s->skip_tag);
    if (s->pump) lxr_xml_pump_destroy(s->pump);
    if (s->zf)   lxr_zip_close_entry(s->zf);
    free(s);
}
