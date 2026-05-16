#include <stdlib.h>
#include <string.h>

#include <expat.h>

#include "lxr_xml_pump.h"

#define LXR_PARSE_BUFFER_SIZE (64 * 1024)

typedef enum {
    LXR_SOURCE_CALLBACK,
    LXR_SOURCE_ZIP_FILE,
    LXR_SOURCE_BUFFER
} lxr_source_kind;

struct lxr_xml_pump {
    XML_Parser parser;

    lxr_source_kind source;
    lxr_xml_read_fn read_fn;
    void           *read_ud;

    /* buffer-source state */
    const char *buf_data;
    size_t      buf_len;
    size_t      buf_pos;

    /* zip-source state */
    lxr_zip_file *zf;

    int suspended;
    int eof;
    int error;
};

/* ------------------------------------------------------------------------- */
/* Built-in source readers                                                   */
/* ------------------------------------------------------------------------- */

static ssize_t read_from_zip_file(void *ud, void *buf, size_t n)
{
    return lxr_zip_read((lxr_zip_file *)ud, buf, n);
}

static ssize_t read_from_buffer(void *ud, void *buf, size_t n)
{
    lxr_xml_pump *p = (lxr_xml_pump *)ud;
    size_t avail = p->buf_len - p->buf_pos;
    if (avail > n) avail = n;
    if (avail == 0) return 0;
    memcpy(buf, p->buf_data + p->buf_pos, avail);
    p->buf_pos += avail;
    return (ssize_t)avail;
}

/* ------------------------------------------------------------------------- */
/* Construction                                                              */
/* ------------------------------------------------------------------------- */

static lxr_xml_pump *new_pump(void)
{
    lxr_xml_pump *p = (lxr_xml_pump *)calloc(1, sizeof(*p));
    if (!p) return NULL;
    p->parser = XML_ParserCreate(NULL);
    if (!p->parser) {
        free(p);
        return NULL;
    }
    return p;
}

lxr_xml_pump *lxr_xml_pump_create_callback(lxr_xml_read_fn fn, void *ud)
{
    lxr_xml_pump *p;
    if (!fn) return NULL;
    p = new_pump();
    if (!p) return NULL;
    p->source  = LXR_SOURCE_CALLBACK;
    p->read_fn = fn;
    p->read_ud = ud;
    return p;
}

lxr_xml_pump *lxr_xml_pump_create_zip_file(lxr_zip_file *zf)
{
    lxr_xml_pump *p;
    if (!zf) return NULL;
    p = new_pump();
    if (!p) return NULL;
    p->source  = LXR_SOURCE_ZIP_FILE;
    p->read_fn = read_from_zip_file;
    p->read_ud = zf;
    p->zf      = zf;
    return p;
}

lxr_xml_pump *lxr_xml_pump_create_buffer(const char *data, size_t len)
{
    lxr_xml_pump *p;
    if (!data) return NULL;
    p = new_pump();
    if (!p) return NULL;
    p->source   = LXR_SOURCE_BUFFER;
    p->buf_data = data;
    p->buf_len  = len;
    p->buf_pos  = 0;
    p->read_fn  = read_from_buffer;
    p->read_ud  = p;
    return p;
}

void lxr_xml_pump_destroy(lxr_xml_pump *p)
{
    if (!p) return;
    if (p->parser) {
        /* set_handlers stashes a bridge struct as the parser's user-data
         * for callback dispatch; expat does not own it, so free it before
         * tearing the parser down. */
        void *bridge = XML_GetUserData(p->parser);
        free(bridge);
        XML_ParserFree(p->parser);
    }
    free(p);
}

/* ------------------------------------------------------------------------- */
/* Handler bridging                                                          */
/* ------------------------------------------------------------------------- */

typedef struct {
    lxr_xml_pump    *pump;
    lxr_xml_start_fn start;
    lxr_xml_end_fn   end;
    lxr_xml_text_fn  text;
    void            *ud;
} lxr_xml_bridge;

static void XMLCALL bridge_start(void *ud, const XML_Char *name, const XML_Char **attrs)
{
    lxr_xml_bridge *b = (lxr_xml_bridge *)ud;
    if (b->start) b->start(b->ud, (const char *)name, (const char **)attrs);
}

static void XMLCALL bridge_end(void *ud, const XML_Char *name)
{
    lxr_xml_bridge *b = (lxr_xml_bridge *)ud;
    if (b->end) b->end(b->ud, (const char *)name);
}

static void XMLCALL bridge_text(void *ud, const XML_Char *text, int len)
{
    lxr_xml_bridge *b = (lxr_xml_bridge *)ud;
    if (b->text) b->text(b->ud, (const char *)text, len);
}

void lxr_xml_pump_set_handlers(lxr_xml_pump    *p,
                               lxr_xml_start_fn s,
                               lxr_xml_end_fn   e,
                               lxr_xml_text_fn  t,
                               void            *ud)
{
    /* The bridge struct lives in user-data slot; we allocate it lazily and
     * leak it by attaching to the parser. For simplicity, just keep one
     * bridge per pump.  Reuse if already allocated. */
    lxr_xml_bridge *b;
    if (!p) return;

    b = (lxr_xml_bridge *)XML_GetUserData(p->parser);
    if (!b) {
        b = (lxr_xml_bridge *)calloc(1, sizeof(*b));
        if (!b) return;
        b->pump = p;
        XML_SetUserData(p->parser, b);
    }
    b->start = s;
    b->end   = e;
    b->text  = t;
    b->ud    = ud;
    XML_SetElementHandler(p->parser, bridge_start, bridge_end);
    XML_SetCharacterDataHandler(p->parser, bridge_text);
}

/* ------------------------------------------------------------------------- */
/* Parsing loop                                                              */
/* ------------------------------------------------------------------------- */

static lxr_error pump_loop(lxr_xml_pump *p, int starting_fresh)
{
    enum XML_Status status;

    if (!p) return LXR_ERROR_NULL_PARAMETER;

    if (p->error) return LXR_ERROR_XML_PARSE;

    /* If a previous suspend left us paused, resume first. */
    if (!starting_fresh && p->suspended) {
        p->suspended = 0;
        status = XML_ResumeParser(p->parser);
        if (status == XML_STATUS_SUSPENDED) {
            p->suspended = 1;
            return LXR_NO_ERROR;
        }
        if (status == XML_STATUS_ERROR) {
            /* If parser was not suspended, ResumeParser sets a specific
             * "not suspended" error which we treat as benign and fall through
             * to the normal feed loop. */
            if (XML_GetErrorCode(p->parser) != XML_ERROR_NOT_SUSPENDED) {
                p->error = 1;
                return LXR_ERROR_XML_PARSE;
            }
        }
    }

    while (!p->eof && !p->suspended) {
        ssize_t got;
        void *buf = XML_GetBuffer(p->parser, LXR_PARSE_BUFFER_SIZE);
        if (!buf) {
            p->error = 1;
            return LXR_ERROR_MEMORY_MALLOC_FAILED;
        }

        got = p->read_fn(p->read_ud, buf, LXR_PARSE_BUFFER_SIZE);
        if (got < 0) {
            p->error = 1;
            return LXR_ERROR_FILE_CORRUPTED;
        }

        if (got == 0) p->eof = 1;

        status = XML_ParseBuffer(p->parser, (int)got, p->eof ? 1 : 0);

        if (status == XML_STATUS_ERROR) {
            p->error = 1;
            return LXR_ERROR_XML_PARSE;
        }
        if (status == XML_STATUS_SUSPENDED) {
            p->suspended = 1;
            return LXR_NO_ERROR;
        }
    }
    return LXR_NO_ERROR;
}

lxr_error lxr_xml_pump_run(lxr_xml_pump *p)
{
    return pump_loop(p, 1);
}

lxr_error lxr_xml_pump_resume(lxr_xml_pump *p)
{
    return pump_loop(p, 0);
}

void lxr_xml_pump_suspend(lxr_xml_pump *p)
{
    if (p && p->parser) XML_StopParser(p->parser, XML_TRUE);
}

int lxr_xml_pump_is_suspended(const lxr_xml_pump *p)
{
    return p ? p->suspended : 0;
}

int lxr_xml_pump_is_eof(const lxr_xml_pump *p)
{
    return p ? p->eof : 0;
}

/* ------------------------------------------------------------------------- */
/* Helpers                                                                   */
/* ------------------------------------------------------------------------- */

const char *lxr_xml_attr(const char **attrs, const char *name)
{
    if (!attrs || !name) return NULL;
    while (*attrs) {
        if (lxr_xml_name_eq(*attrs, name)) return *(attrs + 1);
        attrs += 2;
    }
    return NULL;
}

int lxr_xml_name_eq(const char *xml_name, const char *want)
{
    size_t xn_len, w_len;
    const char *colon;
    if (!xml_name || !want) return 0;
    xn_len = strlen(xml_name);
    w_len  = strlen(want);

    if (xn_len == w_len) return strcmp(xml_name, want) == 0;
    if (xn_len < w_len)  return 0;

    /* Allow prefix:want */
    colon = xml_name + (xn_len - w_len) - 1;
    if (*colon != ':') return 0;
    return strcmp(colon + 1, want) == 0;
}
