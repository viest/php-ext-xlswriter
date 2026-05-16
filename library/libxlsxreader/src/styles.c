#include <stdlib.h>
#include <string.h>

#include "lxr_styles_priv.h"
#include "lxr_numfmt.h"
#include "lxr_xml_pump.h"

typedef struct {
    uint16_t  id;
    char     *fmt;
} lxr_user_numfmt;

struct lxr_styles {
    lxr_user_numfmt *user_fmts;
    size_t           user_fmts_count;
    size_t           user_fmts_cap;

    lxr_xf          *xfs;
    size_t           xfs_count;
    size_t           xfs_cap;

    lxr_font        *fonts;
    size_t           fonts_count;
    size_t           fonts_cap;

    lxr_fill        *fills;
    size_t           fills_count;
    size_t           fills_cap;

    lxr_border      *borders;
    size_t           borders_count;
    size_t           borders_cap;
};

/* ========================================================================= */
/* Public lookup                                                             */
/* ========================================================================= */

const lxr_xf *lxr_styles_get_xf(const lxr_styles *st, uint32_t style_id)
{
    if (!st) return NULL;
    if ((size_t)style_id >= st->xfs_count) return NULL;
    return &st->xfs[style_id];
}

const lxr_font *lxr_styles_get_font(const lxr_styles *st, uint32_t font_id)
{
    if (!st) return NULL;
    if ((size_t)font_id >= st->fonts_count) return NULL;
    return &st->fonts[font_id];
}

const lxr_fill *lxr_styles_get_fill(const lxr_styles *st, uint32_t fill_id)
{
    if (!st) return NULL;
    if ((size_t)fill_id >= st->fills_count) return NULL;
    return &st->fills[fill_id];
}

const lxr_border *lxr_styles_get_border(const lxr_styles *st, uint32_t border_id)
{
    if (!st) return NULL;
    if ((size_t)border_id >= st->borders_count) return NULL;
    return &st->borders[border_id];
}

size_t lxr_styles_count(const lxr_styles *st)
{
    return st ? st->xfs_count : 0;
}

void lxr_styles_free(lxr_styles *st)
{
    size_t i;
    if (!st) return;

    for (i = 0; i < st->user_fmts_count; i++) free(st->user_fmts[i].fmt);
    free(st->user_fmts);

    for (i = 0; i < st->fonts_count; i++) free((void *)st->fonts[i].name);
    free(st->fonts);

    for (i = 0; i < st->fills_count; i++) free((void *)st->fills[i].pattern_type);
    free(st->fills);

    for (i = 0; i < st->borders_count; i++) {
        free((void *)st->borders[i].left.style);
        free((void *)st->borders[i].right.style);
        free((void *)st->borders[i].top.style);
        free((void *)st->borders[i].bottom.style);
    }
    free(st->borders);

    for (i = 0; i < st->xfs_count; i++) {
        free((void *)st->xfs[i].alignment.horizontal);
        free((void *)st->xfs[i].alignment.vertical);
    }
    free(st->xfs);

    free(st);
}

/* ========================================================================= */
/* numFmt                                                                    */
/* ========================================================================= */

static const char *resolve_format(lxr_styles *st, uint16_t id, lxr_fmt_category *cat_out)
{
    size_t      i;
    const char *builtin = lxr_numfmt_builtin_format(id);
    if (builtin) {
        if (cat_out) *cat_out = lxr_numfmt_builtin_category(id);
        return builtin;
    }
    for (i = 0; i < st->user_fmts_count; i++) {
        if (st->user_fmts[i].id == id) {
            const char *f = st->user_fmts[i].fmt;
            if (cat_out) *cat_out = lxr_numfmt_classify(f);
            return f;
        }
    }
    if (cat_out) *cat_out = LXR_FMT_CATEGORY_GENERAL;
    return NULL;
}

static int append_user_fmt(lxr_styles *st, uint16_t id, const char *fmt)
{
    if (st->user_fmts_count >= st->user_fmts_cap) {
        size_t cap = st->user_fmts_cap ? st->user_fmts_cap * 2 : 8;
        lxr_user_numfmt *nb = (lxr_user_numfmt *)realloc(
            st->user_fmts, cap * sizeof(*nb));
        if (!nb) return -1;
        st->user_fmts     = nb;
        st->user_fmts_cap = cap;
    }
    st->user_fmts[st->user_fmts_count].id  = id;
    st->user_fmts[st->user_fmts_count].fmt = strdup(fmt ? fmt : "");
    if (!st->user_fmts[st->user_fmts_count].fmt) return -1;
    st->user_fmts_count++;
    return 0;
}

/* ========================================================================= */
/* Array growers                                                             */
/* ========================================================================= */

#define GROW_ARRAY(arr, count, cap, type, init_cap)                          \
    do {                                                                      \
        if ((count) >= (cap)) {                                               \
            size_t _nc = (cap) ? (cap) * 2 : (init_cap);                      \
            type *_nb  = (type *)realloc((arr), _nc * sizeof(*(arr)));        \
            if (!_nb) return -1;                                              \
            (arr) = _nb;                                                      \
            (cap) = _nc;                                                      \
        }                                                                     \
    } while (0)

static int begin_font(lxr_styles *st)
{
    GROW_ARRAY(st->fonts, st->fonts_count, st->fonts_cap, lxr_font, 8);
    memset(&st->fonts[st->fonts_count], 0, sizeof(lxr_font));
    st->fonts[st->fonts_count].underline = LXR_UNDERLINE_NONE;
    return 0;
}

static int begin_fill(lxr_styles *st)
{
    GROW_ARRAY(st->fills, st->fills_count, st->fills_cap, lxr_fill, 8);
    memset(&st->fills[st->fills_count], 0, sizeof(lxr_fill));
    return 0;
}

static int begin_border(lxr_styles *st)
{
    GROW_ARRAY(st->borders, st->borders_count, st->borders_cap, lxr_border, 8);
    memset(&st->borders[st->borders_count], 0, sizeof(lxr_border));
    return 0;
}

static int begin_xf(lxr_styles *st, const char **attrs)
{
    GROW_ARRAY(st->xfs, st->xfs_count, st->xfs_cap, lxr_xf, 32);
    {
        lxr_xf *x = &st->xfs[st->xfs_count];
        memset(x, 0, sizeof(*x));
        x->locked = 1;  /* Excel default — protection.locked is on unless set off */

        const char *id_s     = lxr_xml_attr(attrs, "numFmtId");
        const char *font_s   = lxr_xml_attr(attrs, "fontId");
        const char *fill_s   = lxr_xml_attr(attrs, "fillId");
        const char *border_s = lxr_xml_attr(attrs, "borderId");

        x->num_fmt_id = id_s     ? (uint16_t)strtoul(id_s,     NULL, 10) : 0;
        x->font_id    = font_s   ? (uint32_t)strtoul(font_s,   NULL, 10) : 0;
        x->fill_id    = fill_s   ? (uint32_t)strtoul(fill_s,   NULL, 10) : 0;
        x->border_id  = border_s ? (uint32_t)strtoul(border_s, NULL, 10) : 0;
        x->format_string = resolve_format(st, x->num_fmt_id, &x->category);
    }
    return 0;
}

/* ========================================================================= */
/* color extraction                                                          */
/* ========================================================================= */

static void extract_color(const char **attrs, char out[LXR_COLOR_LEN])
{
    out[0] = 0;
    if (!attrs) return;
    const char *rgb = lxr_xml_attr(attrs, "rgb");
    if (rgb) {
        size_t n = strlen(rgb);
        if (n >= LXR_COLOR_LEN) n = LXR_COLOR_LEN - 1;
        memcpy(out, rgb, n);
        out[n] = 0;
    }
    /* theme/indexed/auto colors: leave as "" — we don't resolve them. */
}

static int parse_bool_attr(const char **attrs)
{
    const char *v = lxr_xml_attr(attrs, "val");
    if (!v) return 1;          /* presence alone == true */
    if (*v == '0' || *v == 'f' || *v == 'F') return 0;
    return 1;
}

/* ========================================================================= */
/* SAX state machine                                                         */
/* ========================================================================= */

typedef enum {
    SC_NONE,
    SC_NUMFMTS,
    SC_FONTS,
    SC_FONT,
    SC_FILLS,
    SC_FILL,
    SC_PATTERN_FILL,
    SC_BORDERS,
    SC_BORDER,
    SC_BORDER_LEFT,
    SC_BORDER_RIGHT,
    SC_BORDER_TOP,
    SC_BORDER_BOTTOM,
    SC_CELL_XFS,
    SC_CELL_STYLE_XFS,
    SC_XF
} scope_t;

typedef struct {
    lxr_styles *st;
    scope_t     stack[16];
    int         depth;
    int         in_alignment;        /* parsed inside SC_XF */
    int         in_protection;       /* parsed inside SC_XF */
} parse_ctx;

static scope_t top(parse_ctx *c) { return c->depth > 0 ? c->stack[c->depth - 1] : SC_NONE; }
static void push(parse_ctx *c, scope_t s) {
    if (c->depth < (int)(sizeof(c->stack)/sizeof(c->stack[0]))) c->stack[c->depth++] = s;
}
static void pop(parse_ctx *c) { if (c->depth > 0) c->depth--; }

static lxr_border_side *current_side(parse_ctx *c)
{
    lxr_border *b = &c->st->borders[c->st->borders_count];
    switch (top(c)) {
    case SC_BORDER_LEFT:   return &b->left;
    case SC_BORDER_RIGHT:  return &b->right;
    case SC_BORDER_TOP:    return &b->top;
    case SC_BORDER_BOTTOM: return &b->bottom;
    default:               return NULL;
    }
}

static void on_start(void *ud, const char *name, const char **attrs)
{
    parse_ctx *c  = (parse_ctx *)ud;
    scope_t    sc = top(c);

    /* --- numFmts ------------------------------------------------------ */
    if (lxr_xml_name_eq(name, "numFmts"))    { push(c, SC_NUMFMTS);   return; }

    if (sc == SC_NUMFMTS && lxr_xml_name_eq(name, "numFmt")) {
        const char *id_s  = lxr_xml_attr(attrs, "numFmtId");
        const char *fmt_s = lxr_xml_attr(attrs, "formatCode");
        if (id_s && fmt_s) append_user_fmt(c->st, (uint16_t)strtoul(id_s, NULL, 10), fmt_s);
        return;
    }

    /* --- fonts -------------------------------------------------------- */
    if (lxr_xml_name_eq(name, "fonts"))      { push(c, SC_FONTS);     return; }
    if (sc == SC_FONTS && lxr_xml_name_eq(name, "font")) {
        begin_font(c->st);
        push(c, SC_FONT);
        return;
    }
    if (sc == SC_FONT) {
        lxr_font *f = &c->st->fonts[c->st->fonts_count];
        if (lxr_xml_name_eq(name, "sz")) {
            const char *v = lxr_xml_attr(attrs, "val");
            if (v) f->size = strtod(v, NULL);
        } else if (lxr_xml_name_eq(name, "name") || lxr_xml_name_eq(name, "rFont")) {
            const char *v = lxr_xml_attr(attrs, "val");
            if (v) { free((void *)f->name); f->name = strdup(v); }
        } else if (lxr_xml_name_eq(name, "color")) {
            extract_color(attrs, f->color);
        } else if (lxr_xml_name_eq(name, "b"))      { f->bold      = parse_bool_attr(attrs); }
        else if   (lxr_xml_name_eq(name, "i"))      { f->italic    = parse_bool_attr(attrs); }
        else if   (lxr_xml_name_eq(name, "strike")) { f->strike    = parse_bool_attr(attrs); }
        else if   (lxr_xml_name_eq(name, "u"))      {
            const char *v = lxr_xml_attr(attrs, "val");
            if (!v || strcmp(v, "single") == 0)               f->underline = LXR_UNDERLINE_SINGLE;
            else if (strcmp(v, "double") == 0)                f->underline = LXR_UNDERLINE_DOUBLE;
            else if (strcmp(v, "singleAccounting") == 0)      f->underline = LXR_UNDERLINE_SINGLE_ACCOUNTING;
            else if (strcmp(v, "doubleAccounting") == 0)      f->underline = LXR_UNDERLINE_DOUBLE_ACCOUNTING;
            else                                              f->underline = LXR_UNDERLINE_SINGLE;
        }
        return;
    }

    /* --- fills -------------------------------------------------------- */
    if (lxr_xml_name_eq(name, "fills"))      { push(c, SC_FILLS);     return; }
    if (sc == SC_FILLS && lxr_xml_name_eq(name, "fill")) {
        begin_fill(c->st);
        push(c, SC_FILL);
        return;
    }
    if (sc == SC_FILL && lxr_xml_name_eq(name, "patternFill")) {
        lxr_fill   *fl  = &c->st->fills[c->st->fills_count];
        const char *pt  = lxr_xml_attr(attrs, "patternType");
        if (pt) { free((void *)fl->pattern_type); fl->pattern_type = strdup(pt); }
        push(c, SC_PATTERN_FILL);
        return;
    }
    if (sc == SC_PATTERN_FILL) {
        lxr_fill *fl = &c->st->fills[c->st->fills_count];
        if      (lxr_xml_name_eq(name, "fgColor")) extract_color(attrs, fl->fg_color);
        else if (lxr_xml_name_eq(name, "bgColor")) extract_color(attrs, fl->bg_color);
        return;
    }

    /* --- borders ------------------------------------------------------ */
    if (lxr_xml_name_eq(name, "borders"))    { push(c, SC_BORDERS);   return; }
    if (sc == SC_BORDERS && lxr_xml_name_eq(name, "border")) {
        begin_border(c->st);
        push(c, SC_BORDER);
        return;
    }
    if (sc == SC_BORDER) {
        lxr_border *b   = &c->st->borders[c->st->borders_count];
        lxr_border_side *side = NULL;
        scope_t          ns   = SC_BORDER;

        if      (lxr_xml_name_eq(name, "left"))   { side = &b->left;   ns = SC_BORDER_LEFT; }
        else if (lxr_xml_name_eq(name, "right"))  { side = &b->right;  ns = SC_BORDER_RIGHT; }
        else if (lxr_xml_name_eq(name, "top"))    { side = &b->top;    ns = SC_BORDER_TOP; }
        else if (lxr_xml_name_eq(name, "bottom")) { side = &b->bottom; ns = SC_BORDER_BOTTOM; }

        if (side) {
            const char *st_attr = lxr_xml_attr(attrs, "style");
            if (st_attr) { free((void *)side->style); side->style = strdup(st_attr); }
            push(c, ns);
        }
        return;
    }
    if (sc == SC_BORDER_LEFT || sc == SC_BORDER_RIGHT
        || sc == SC_BORDER_TOP || sc == SC_BORDER_BOTTOM) {
        if (lxr_xml_name_eq(name, "color")) {
            lxr_border_side *side = current_side(c);
            if (side) extract_color(attrs, side->color);
        }
        return;
    }

    /* --- cellXfs ------------------------------------------------------ */
    if (lxr_xml_name_eq(name, "cellXfs"))      { push(c, SC_CELL_XFS);       return; }
    if (lxr_xml_name_eq(name, "cellStyleXfs")) { push(c, SC_CELL_STYLE_XFS); return; }

    if (sc == SC_CELL_XFS && lxr_xml_name_eq(name, "xf")) {
        begin_xf(c->st, attrs);
        push(c, SC_XF);
        return;
    }

    if (sc == SC_XF) {
        lxr_xf *x = &c->st->xfs[c->st->xfs_count];
        if (lxr_xml_name_eq(name, "alignment")) {
            const char *h = lxr_xml_attr(attrs, "horizontal");
            const char *v = lxr_xml_attr(attrs, "vertical");
            const char *w = lxr_xml_attr(attrs, "wrapText");
            const char *i = lxr_xml_attr(attrs, "indent");
            const char *r = lxr_xml_attr(attrs, "textRotation");
            x->has_alignment = 1;
            if (h) { free((void *)x->alignment.horizontal); x->alignment.horizontal = strdup(h); }
            if (v) { free((void *)x->alignment.vertical);   x->alignment.vertical   = strdup(v); }
            if (w) x->alignment.wrap_text = (strcmp(w, "1") == 0 || strcmp(w, "true") == 0);
            if (i) x->alignment.indent    = (int)strtol(i, NULL, 10);
            if (r) x->alignment.rotation  = (int)strtol(r, NULL, 10);
            return;
        }
        if (lxr_xml_name_eq(name, "protection")) {
            const char *l = lxr_xml_attr(attrs, "locked");
            const char *h = lxr_xml_attr(attrs, "hidden");
            if (l) x->locked = (strcmp(l, "1") == 0 || strcmp(l, "true") == 0);
            if (h) x->hidden = (strcmp(h, "1") == 0 || strcmp(h, "true") == 0);
            return;
        }
    }
}

static void on_end(void *ud, const char *name)
{
    parse_ctx *c = (parse_ctx *)ud;
    scope_t    sc = top(c);

    if (lxr_xml_name_eq(name, "numFmts")    && sc == SC_NUMFMTS)    pop(c);
    else if (lxr_xml_name_eq(name, "font")  && sc == SC_FONT)       { c->st->fonts_count++;   pop(c); }
    else if (lxr_xml_name_eq(name, "fonts") && sc == SC_FONTS)      pop(c);
    else if (lxr_xml_name_eq(name, "patternFill") && sc == SC_PATTERN_FILL) pop(c);
    else if (lxr_xml_name_eq(name, "fill")  && sc == SC_FILL)       { c->st->fills_count++;   pop(c); }
    else if (lxr_xml_name_eq(name, "fills") && sc == SC_FILLS)      pop(c);
    else if ((lxr_xml_name_eq(name, "left")   && sc == SC_BORDER_LEFT)
          || (lxr_xml_name_eq(name, "right")  && sc == SC_BORDER_RIGHT)
          || (lxr_xml_name_eq(name, "top")    && sc == SC_BORDER_TOP)
          || (lxr_xml_name_eq(name, "bottom") && sc == SC_BORDER_BOTTOM)) {
        pop(c);
    }
    else if (lxr_xml_name_eq(name, "border")  && sc == SC_BORDER)   { c->st->borders_count++; pop(c); }
    else if (lxr_xml_name_eq(name, "borders") && sc == SC_BORDERS)  pop(c);
    else if (lxr_xml_name_eq(name, "xf") && sc == SC_XF)            { c->st->xfs_count++;     pop(c); }
    else if (lxr_xml_name_eq(name, "cellXfs") && sc == SC_CELL_XFS) pop(c);
    else if (lxr_xml_name_eq(name, "cellStyleXfs") && sc == SC_CELL_STYLE_XFS) pop(c);
}

/* ========================================================================= */
/* entry point                                                               */
/* ========================================================================= */

lxr_error lxr_styles_load(lxr_zip *zip, const char *path, lxr_styles **out)
{
    lxr_zip_file *zf;
    lxr_xml_pump *pump;
    lxr_styles   *st;
    parse_ctx     ctx;
    lxr_error     rc;

    if (!zip || !out) return LXR_ERROR_NULL_PARAMETER;

    st = (lxr_styles *)calloc(1, sizeof(*st));
    if (!st) return LXR_ERROR_MEMORY_MALLOC_FAILED;

    if (!path) {
        *out = st;
        return LXR_NO_ERROR;
    }

    zf = lxr_zip_open_entry(zip, path);
    if (!zf) {
        *out = st;
        return LXR_NO_ERROR;
    }

    pump = lxr_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxr_zip_close_entry(zf);
        free(st);
        return LXR_ERROR_MEMORY_MALLOC_FAILED;
    }

    memset(&ctx, 0, sizeof(ctx));
    ctx.st = st;

    lxr_xml_pump_set_handlers(pump, on_start, on_end, NULL, &ctx);
    rc = lxr_xml_pump_run(pump);

    lxr_xml_pump_destroy(pump);
    lxr_zip_close_entry(zf);

    if (rc != LXR_NO_ERROR) {
        lxr_styles_free(st);
        return rc;
    }
    *out = st;
    return LXR_NO_ERROR;
}
