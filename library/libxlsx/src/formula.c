/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 *
 * Excel formula evaluation engine: lex -> parse (recursive descent) -> walk.
 * Cell references resolve through a caller callback (see formula.h). No php.h.
 */
#include <ctype.h>
#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "libxlsx/formula.h"
#include "libxlsx/utility.h"

/* ===================================================================== *
 * Value helpers
 * ===================================================================== */

void lxlsx_value_free(lxlsx_value *value)
{
    if (!value)
        return;
    if (value->kind == LXLSX_VAL_STRING)
        free(value->string);
    value->string = NULL;
    value->kind = LXLSX_VAL_BLANK;
    value->number = 0.0;
    value->error = LXLSX_FERR_NONE;
}

const char *lxlsx_formula_error_string(lxlsx_formula_error error)
{
    switch (error) {
    case LXLSX_FERR_NULL:  return "#NULL!";
    case LXLSX_FERR_DIV0:  return "#DIV/0!";
    case LXLSX_FERR_VALUE: return "#VALUE!";
    case LXLSX_FERR_REF:   return "#REF!";
    case LXLSX_FERR_NAME:  return "#NAME?";
    case LXLSX_FERR_NUM:   return "#NUM!";
    case LXLSX_FERR_NA:    return "#N/A";
    default:               return "";
    }
}

static lxlsx_value val_number(double n)
{
    lxlsx_value v; v.kind = LXLSX_VAL_NUMBER; v.number = n;
    v.string = NULL; v.error = LXLSX_FERR_NONE; return v;
}
static lxlsx_value val_bool(int b)
{
    lxlsx_value v; v.kind = LXLSX_VAL_BOOL; v.number = b ? 1.0 : 0.0;
    v.string = NULL; v.error = LXLSX_FERR_NONE; return v;
}
static lxlsx_value val_blank(void)
{
    lxlsx_value v; v.kind = LXLSX_VAL_BLANK; v.number = 0.0;
    v.string = NULL; v.error = LXLSX_FERR_NONE; return v;
}
static lxlsx_value val_error(lxlsx_formula_error e)
{
    lxlsx_value v; v.kind = LXLSX_VAL_ERROR; v.number = 0.0;
    v.string = NULL; v.error = e; return v;
}
static lxlsx_value val_string_take(char *s)  /* takes ownership */
{
    lxlsx_value v; v.kind = LXLSX_VAL_STRING; v.number = 0.0;
    v.string = s; v.error = LXLSX_FERR_NONE; return v;
}
static lxlsx_value val_string_copy(const char *s, size_t n)
{
    char *p = malloc(n + 1);
    if (!p) return val_error(LXLSX_FERR_VALUE);
    memcpy(p, s, n); p[n] = '\0';
    return val_string_take(p);
}

/* ===================================================================== *
 * Lexer (ported from the zval AST builder, value-free)
 * ===================================================================== */

enum {
    TK_END = 0, TK_NUMBER, TK_STRING, TK_BOOL, TK_ERROR, TK_NAME,
    TK_LPAREN, TK_RPAREN, TK_LBRACE, TK_RBRACE, TK_COMMA, TK_SEMI, TK_OP
};

typedef struct {
    int kind;
    const char *start;
    size_t len;
    double num;
    int boolean;
    lxlsx_formula_error err;
    char op[3];
    int opn;
} tok;

typedef struct {
    const char *src;
    size_t pos, n;
    tok cur;
    int failed;     /* parse error */
} pstate;

static int is_idstart(char c) {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c == '$';
}
static int is_idcont(char c) {
    return is_idstart(c) || (c >= '0' && c <= '9') || c == '.';
}

static void lex_next(pstate *p)
{
    while (p->pos < p->n) {
        char c = p->src[p->pos];
        if (c == ' ' || c == '\t' || c == '\n' || c == '\r') p->pos++;
        else break;
    }
    p->cur.opn = 0;
    if (p->pos >= p->n) { p->cur.kind = TK_END; return; }

    char c = p->src[p->pos];
    p->cur.start = p->src + p->pos;

    if (c == '(') { p->cur.kind = TK_LPAREN; p->pos++; return; }
    if (c == ')') { p->cur.kind = TK_RPAREN; p->pos++; return; }
    if (c == '{') { p->cur.kind = TK_LBRACE; p->pos++; return; }
    if (c == '}') { p->cur.kind = TK_RBRACE; p->pos++; return; }
    if (c == ',') { p->cur.kind = TK_COMMA;  p->pos++; return; }
    if (c == ';') { p->cur.kind = TK_SEMI;   p->pos++; return; }

    if (c == '"') {              /* string literal, "" escapes a quote */
        size_t start = ++p->pos;
        while (p->pos < p->n) {
            if (p->src[p->pos] == '"') {
                if (p->pos + 1 < p->n && p->src[p->pos + 1] == '"') p->pos += 2;
                else break;
            } else p->pos++;
        }
        p->cur.kind = TK_STRING;
        p->cur.start = p->src + start;
        p->cur.len = p->pos - start;
        if (p->pos < p->n) p->pos++;  /* closing quote */
        return;
    }

    if (c == '#') {              /* error literal, e.g. #DIV/0! */
        size_t start = p->pos++;
        while (p->pos < p->n) {
            char e = p->src[p->pos];
            if (e == '!' || e == '?') { p->pos++; break; }
            if (e == ' ' || e == ')' || e == ',') break;
            p->pos++;
        }
        size_t len = p->pos - start;
        p->cur.kind = TK_ERROR;
        p->cur.err = LXLSX_FERR_VALUE;
        if (len >= 4 && strncmp(p->src + start, "#REF", 4) == 0) p->cur.err = LXLSX_FERR_REF;
        else if (strncmp(p->src + start, "#DIV", 4) == 0) p->cur.err = LXLSX_FERR_DIV0;
        else if (strncmp(p->src + start, "#NAME", 5) == 0) p->cur.err = LXLSX_FERR_NAME;
        else if (strncmp(p->src + start, "#NULL", 5) == 0) p->cur.err = LXLSX_FERR_NULL;
        else if (strncmp(p->src + start, "#NUM", 4) == 0) p->cur.err = LXLSX_FERR_NUM;
        else if (strncmp(p->src + start, "#N/A", 4) == 0) p->cur.err = LXLSX_FERR_NA;
        else if (strncmp(p->src + start, "#VALUE", 6) == 0) p->cur.err = LXLSX_FERR_VALUE;
        return;
    }

    if ((c >= '0' && c <= '9') || c == '.') {   /* number */
        char *end = NULL;
        p->cur.num = strtod(p->src + p->pos, &end);
        p->cur.kind = TK_NUMBER;
        p->pos = (size_t)(end - p->src);
        return;
    }

    if (is_idstart(c)) {                         /* name / cell ref / bool */
        size_t start = p->pos;
        /* Allow a single sheet-qualifier '!': Sheet1!A1 (sheet ignored). */
        while (p->pos < p->n && (is_idcont(p->src[p->pos]) || p->src[p->pos] == '!'))
            p->pos++;
        size_t len = p->pos - start;
        p->cur.start = p->src + start;
        p->cur.len = len;
        if ((len == 4 && lxlsx_strncasecmp(p->src + start, "TRUE", 4) == 0)) {
            p->cur.kind = TK_BOOL; p->cur.boolean = 1; return;
        }
        if ((len == 5 && lxlsx_strncasecmp(p->src + start, "FALSE", 5) == 0)) {
            p->cur.kind = TK_BOOL; p->cur.boolean = 0; return;
        }
        p->cur.kind = TK_NAME;
        return;
    }

    /* operator: one or two chars from + - * / ^ & = < > % : */
    p->cur.kind = TK_OP;
    p->cur.op[0] = c; p->cur.opn = 1; p->pos++;
    if (p->pos < p->n) {
        char d = p->src[p->pos];
        if ((c == '<' && (d == '>' || d == '=')) || (c == '>' && d == '=')) {
            p->cur.op[1] = d; p->cur.opn = 2; p->pos++;
        }
    }
    p->cur.op[p->cur.opn] = '\0';
}

static int op_is(pstate *p, const char *s)
{
    return p->cur.kind == TK_OP && strcmp(p->cur.op, s) == 0;
}

/* ===================================================================== *
 * AST
 * ===================================================================== */

typedef enum {
    N_NUM, N_STR, N_BOOL, N_ERR, N_REF, N_RANGE, N_UNARY, N_BINARY, N_FUNC, N_NAME
} node_kind;

typedef struct node {
    node_kind kind;
    double num;
    char *str;
    int boolean;
    lxlsx_formula_error err;
    lxlsx_row_t row, row2;
    lxlsx_col_t col, col2;
    char op[3];
    struct node *a, *b;
    char *fname;
    struct node **args;
    int nargs;
} node;

static node *node_new(node_kind k)
{
    node *n = calloc(1, sizeof(node));
    if (n) n->kind = k;
    return n;
}

static void node_free(node *n)
{
    int i;
    if (!n) return;
    free(n->str);
    free(n->fname);
    node_free(n->a);
    node_free(n->b);
    for (i = 0; i < n->nargs; i++) node_free(n->args[i]);
    free(n->args);
    free(n);
}

/* Parse "A1" / "$A$1" / "Sheet1!A1" into 0-based row/col. Returns 1 on success. */
static int parse_ref(const char *s, size_t len, lxlsx_row_t *row, lxlsx_col_t *col)
{
    size_t i = 0;
    /* drop sheet qualifier */
    const char *bang = memchr(s, '!', len);
    if (bang) { size_t off = (size_t)(bang - s) + 1; s += off; len -= off; }

    unsigned long c = 0; int have_col = 0;
    if (i < len && s[i] == '$') i++;
    while (i < len) {
        char ch = s[i];
        if (ch >= 'a' && ch <= 'z') ch = (char)(ch - 'a' + 'A');
        if (ch >= 'A' && ch <= 'Z') { c = c * 26 + (unsigned long)(ch - 'A' + 1); have_col = 1; i++; }
        else break;
    }
    if (!have_col) return 0;
    if (i < len && s[i] == '$') i++;
    unsigned long r = 0; int have_row = 0;
    while (i < len && s[i] >= '0' && s[i] <= '9') { r = r * 10 + (unsigned long)(s[i] - '0'); have_row = 1; i++; }
    if (!have_row || i != len) return 0;
    if (c == 0 || r == 0) return 0;
    *col = (lxlsx_col_t)(c - 1);
    *row = (lxlsx_row_t)(r - 1);
    return 1;
}

/* forward */
static node *parse_expr(pstate *p);

static node *parse_args(pstate *p, node *fn)
{
    /* current token is LPAREN */
    lex_next(p);
    if (p->cur.kind == TK_RPAREN) { lex_next(p); return fn; }
    for (;;) {
        node *arg = parse_expr(p);
        if (!arg) { p->failed = 1; return fn; }
        node **na = realloc(fn->args, sizeof(node *) * (size_t)(fn->nargs + 1));
        if (!na) { node_free(arg); p->failed = 1; return fn; }
        fn->args = na;
        fn->args[fn->nargs++] = arg;
        if (p->cur.kind == TK_COMMA || p->cur.kind == TK_SEMI) { lex_next(p); continue; }
        break;
    }
    if (p->cur.kind == TK_RPAREN) lex_next(p);
    else p->failed = 1;
    return fn;
}

static node *parse_atom(pstate *p)
{
    if (p->cur.kind == TK_NUMBER) {
        node *n = node_new(N_NUM); if (!n) { p->failed = 1; return NULL; }
        n->num = p->cur.num; lex_next(p); return n;
    }
    if (p->cur.kind == TK_STRING) {
        node *n = node_new(N_STR); if (!n) { p->failed = 1; return NULL; }
        n->str = malloc(p->cur.len + 1);
        if (!n->str) { node_free(n); p->failed = 1; return NULL; }
        /* collapse "" -> " */
        { size_t i, j = 0; for (i = 0; i < p->cur.len; i++) {
            n->str[j++] = p->cur.start[i];
            if (p->cur.start[i] == '"' && i + 1 < p->cur.len && p->cur.start[i + 1] == '"') i++;
        } n->str[j] = '\0'; }
        lex_next(p); return n;
    }
    if (p->cur.kind == TK_BOOL) {
        node *n = node_new(N_BOOL); if (!n) { p->failed = 1; return NULL; }
        n->boolean = p->cur.boolean; lex_next(p); return n;
    }
    if (p->cur.kind == TK_ERROR) {
        node *n = node_new(N_ERR); if (!n) { p->failed = 1; return NULL; }
        n->err = p->cur.err; lex_next(p); return n;
    }
    if (p->cur.kind == TK_LPAREN) {
        lex_next(p);
        node *e = parse_expr(p);
        if (p->cur.kind == TK_RPAREN) lex_next(p); else p->failed = 1;
        return e;
    }
    if (p->cur.kind == TK_NAME) {
        const char *s = p->cur.start; size_t len = p->cur.len;
        /* function call? */
        /* peek: save state, advance */
        pstate save = *p;
        lex_next(p);
        if (p->cur.kind == TK_LPAREN) {
            node *fn = node_new(N_FUNC); if (!fn) { p->failed = 1; return NULL; }
            fn->fname = malloc(len + 1);
            if (!fn->fname) { node_free(fn); p->failed = 1; return NULL; }
            memcpy(fn->fname, s, len); fn->fname[len] = '\0';
            return parse_args(p, fn);
        }
        /* not a call: restore and treat as ref or name */
        *p = save;
        lxlsx_row_t row; lxlsx_col_t col;
        if (parse_ref(s, len, &row, &col)) {
            node *n = node_new(N_REF); if (!n) { p->failed = 1; return NULL; }
            n->row = row; n->col = col; lex_next(p); return n;
        }
        node *n = node_new(N_NAME); if (!n) { p->failed = 1; return NULL; }
        lex_next(p); return n;  /* unknown name -> #NAME? at eval */
    }
    p->failed = 1;
    return NULL;
}

static node *parse_range(pstate *p)
{
    node *left = parse_atom(p);
    if (left && left->kind == N_REF && op_is(p, ":")) {
        lex_next(p);
        node *right = parse_atom(p);
        if (right && right->kind == N_REF) {
            node *rng = node_new(N_RANGE);
            if (!rng) { node_free(left); node_free(right); p->failed = 1; return NULL; }
            rng->row = left->row; rng->col = left->col;
            rng->row2 = right->row; rng->col2 = right->col;
            node_free(left); node_free(right);
            return rng;
        }
        node_free(right);
        p->failed = 1;
    }
    return left;
}

static node *make_unary(pstate *p, const char *op, node *operand)
{
    node *n = node_new(N_UNARY);
    if (!n) { node_free(operand); p->failed = 1; return NULL; }
    strcpy(n->op, op); n->a = operand; return n;
}
static node *make_binary(pstate *p, const char *op, node *a, node *b)
{
    node *n = node_new(N_BINARY);
    if (!n) { node_free(a); node_free(b); p->failed = 1; return NULL; }
    strcpy(n->op, op); n->a = a; n->b = b; return n;
}

static node *parse_postfix(pstate *p)
{
    node *n = parse_range(p);
    while (op_is(p, "%")) { lex_next(p); n = make_unary(p, "%", n); }
    return n;
}
static node *parse_unary(pstate *p)
{
    if (op_is(p, "-") || op_is(p, "+")) {
        char op[3]; strcpy(op, p->cur.op); lex_next(p);
        return make_unary(p, op, parse_unary(p));
    }
    return parse_postfix(p);
}
static node *parse_pow(pstate *p)
{
    node *a = parse_unary(p);
    while (op_is(p, "^")) { lex_next(p); a = make_binary(p, "^", a, parse_unary(p)); }
    return a;
}
static node *parse_muldiv(pstate *p)
{
    node *a = parse_pow(p);
    while (op_is(p, "*") || op_is(p, "/")) {
        char op[3]; strcpy(op, p->cur.op); lex_next(p);
        a = make_binary(p, op, a, parse_pow(p));
    }
    return a;
}
static node *parse_addsub(pstate *p)
{
    node *a = parse_muldiv(p);
    while (op_is(p, "+") || op_is(p, "-")) {
        char op[3]; strcpy(op, p->cur.op); lex_next(p);
        a = make_binary(p, op, a, parse_muldiv(p));
    }
    return a;
}
static node *parse_concat(pstate *p)
{
    node *a = parse_addsub(p);
    while (op_is(p, "&")) { lex_next(p); a = make_binary(p, "&", a, parse_addsub(p)); }
    return a;
}
static node *parse_expr(pstate *p)
{
    node *a = parse_concat(p);
    while (op_is(p, "=") || op_is(p, "<>") || op_is(p, "<") ||
           op_is(p, ">") || op_is(p, "<=") || op_is(p, ">=")) {
        char op[3]; strcpy(op, p->cur.op); lex_next(p);
        a = make_binary(p, op, a, parse_concat(p));
    }
    return a;
}

/* ===================================================================== *
 * Evaluator
 * ===================================================================== */

typedef struct {
    lxlsx_formula_resolver resolver;
    void *ctx;
} ev;

static lxlsx_value eval(ev *e, node *n);

/* Coerce a value to a number; *err set on failure. Frees nothing. */
static double to_number(const lxlsx_value *v, lxlsx_formula_error *err)
{
    *err = LXLSX_FERR_NONE;
    switch (v->kind) {
    case LXLSX_VAL_NUMBER:
    case LXLSX_VAL_BOOL:   return v->number;
    case LXLSX_VAL_BLANK:  return 0.0;
    case LXLSX_VAL_ERROR:  *err = v->error; return 0.0;
    case LXLSX_VAL_STRING: {
        if (!v->string || !*v->string) return 0.0;
        char *end = NULL;
        double d = strtod(v->string, &end);
        while (end && (*end == ' ' || *end == '\t')) end++;
        if (end && *end == '\0') return d;
        *err = LXLSX_FERR_VALUE; return 0.0;
    }
    }
    return 0.0;
}

/* Coerce a value to an owned string. */
static char *to_string(const lxlsx_value *v)
{
    char buf[64];
    switch (v->kind) {
    case LXLSX_VAL_STRING:
        return v->string ? strdup(v->string) : strdup("");
    case LXLSX_VAL_BOOL:
        return strdup(v->number != 0.0 ? "TRUE" : "FALSE");
    case LXLSX_VAL_BLANK:
        return strdup("");
    case LXLSX_VAL_ERROR:
        return strdup(lxlsx_formula_error_string(v->error));
    case LXLSX_VAL_NUMBER:
    default: {
        double d = v->number;
        if (d == (double)(long long)d && fabs(d) < 1e15)
            snprintf(buf, sizeof(buf), "%lld", (long long)d);
        else
            snprintf(buf, sizeof(buf), "%.15g", d);
        return strdup(buf);
    }
    }
}

static int value_is_error(const lxlsx_value *v) { return v->kind == LXLSX_VAL_ERROR; }

static lxlsx_value resolve_cell(ev *e, lxlsx_row_t row, lxlsx_col_t col)
{
    lxlsx_value out = val_blank();
    if (e->resolver) e->resolver(e->ctx, row, col, &out);
    return out;
}

/* Iterate a range, calling cb for each cell value (cb takes ownership-free
 * borrow; the value is freed by the iterator). Stops early if cb returns 0. */
typedef int (*range_cb)(void *acc, const lxlsx_value *v);

static void iterate_range(ev *e, node *n, range_cb cb, void *acc)
{
    lxlsx_row_t r;
    lxlsx_col_t c;
    lxlsx_row_t r1 = n->row, r2 = n->row2;
    lxlsx_col_t c1 = n->col, c2 = n->col2;
    if (r1 > r2) { lxlsx_row_t t = r1; r1 = r2; r2 = t; }
    if (c1 > c2) { lxlsx_col_t t = c1; c1 = c2; c2 = t; }
    for (r = r1; r <= r2; r++) {
        for (c = c1; c <= c2; c++) {
            lxlsx_value v = resolve_cell(e, r, c);
            int cont = cb(acc, &v);
            lxlsx_value_free(&v);
            if (!cont) return;
        }
    }
}

/* --- aggregate accumulators --- */
typedef struct { double sum; long count; double minv, maxv; long counta;
                 lxlsx_formula_error err; int has; } agg;

static int agg_cb(void *a, const lxlsx_value *v)
{
    agg *g = a;
    if (v->kind == LXLSX_VAL_ERROR) { g->err = v->error; return 0; }
    if (v->kind != LXLSX_VAL_BLANK) g->counta++;
    /* SUM/AVG/MIN/MAX/COUNT consider only numbers (Excel ignores text/bool in ranges) */
    if (v->kind == LXLSX_VAL_NUMBER) {
        g->sum += v->number; g->count++;
        if (!g->has) { g->minv = g->maxv = v->number; g->has = 1; }
        else { if (v->number < g->minv) g->minv = v->number;
               if (v->number > g->maxv) g->maxv = v->number; }
    }
    return 1;
}

/* Feed one argument node into an aggregate: ranges iterate, scalars count once
 * (scalars DO coerce text/bool to number for SUM, per Excel direct-arg rules). */
static void agg_feed(ev *e, node *arg, agg *g, int numeric_scalar)
{
    if (g->err) return;
    if (arg->kind == N_RANGE) { iterate_range(e, arg, agg_cb, g); return; }
    lxlsx_value v = eval(e, arg);
    if (v.kind == LXLSX_VAL_ERROR) { g->err = v.error; lxlsx_value_free(&v); return; }
    if (v.kind != LXLSX_VAL_BLANK) g->counta++;
    if (numeric_scalar) {
        lxlsx_formula_error er;
        double d = to_number(&v, &er);
        if (er) { g->err = er; lxlsx_value_free(&v); return; }
        if (v.kind != LXLSX_VAL_BLANK) {
            g->sum += d; g->count++;
            if (!g->has) { g->minv = g->maxv = d; g->has = 1; }
            else { if (d < g->minv) g->minv = d; if (d > g->maxv) g->maxv = d; }
        }
    } else if (v.kind == LXLSX_VAL_NUMBER) {
        g->sum += v.number; g->count++;
        if (!g->has) { g->minv = g->maxv = v.number; g->has = 1; }
        else { if (v.number < g->minv) g->minv = v.number;
               if (v.number > g->maxv) g->maxv = v.number; }
    }
    lxlsx_value_free(&v);
}

static int streq_ci(const char *a, const char *b)
{
    return lxlsx_strcasecmp(a, b) == 0;
}

/* Evaluate function call. */
static lxlsx_value eval_func(ev *e, node *n)
{
    const char *fn = n->fname;
    int i;

    /* Aggregates over args/ranges. */
    if (streq_ci(fn, "SUM") || streq_ci(fn, "AVERAGE") || streq_ci(fn, "MIN") ||
        streq_ci(fn, "MAX") || streq_ci(fn, "COUNT") || streq_ci(fn, "COUNTA")) {
        agg g; memset(&g, 0, sizeof(g));
        int numeric_scalar = !streq_ci(fn, "COUNT") && !streq_ci(fn, "COUNTA");
        for (i = 0; i < n->nargs; i++) agg_feed(e, n->args[i], &g, numeric_scalar);
        if (g.err) return val_error(g.err);
        if (streq_ci(fn, "SUM")) return val_number(g.sum);
        if (streq_ci(fn, "COUNT")) return val_number((double)g.count);
        if (streq_ci(fn, "COUNTA")) return val_number((double)g.counta);
        if (streq_ci(fn, "AVERAGE"))
            return g.count ? val_number(g.sum / (double)g.count) : val_error(LXLSX_FERR_DIV0);
        if (streq_ci(fn, "MIN")) return val_number(g.has ? g.minv : 0.0);
        if (streq_ci(fn, "MAX")) return val_number(g.has ? g.maxv : 0.0);
    }

    /* IF(cond, a, [b]) */
    if (streq_ci(fn, "IF")) {
        if (n->nargs < 2) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value c = eval(e, n->args[0]);
        if (value_is_error(&c)) return c;
        lxlsx_formula_error er; double d = to_number(&c, &er);
        lxlsx_value_free(&c);
        if (er) return val_error(er);
        if (d != 0.0) return eval(e, n->args[1]);
        if (n->nargs >= 3) return eval(e, n->args[2]);
        return val_bool(0);
    }
    if (streq_ci(fn, "IFERROR")) {
        if (n->nargs < 2) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) { lxlsx_value_free(&v); return eval(e, n->args[1]); }
        return v;
    }

    /* Logical AND/OR/NOT */
    if (streq_ci(fn, "AND") || streq_ci(fn, "OR")) {
        int is_and = streq_ci(fn, "AND");
        int result = is_and ? 1 : 0;
        for (i = 0; i < n->nargs; i++) {
            lxlsx_value v = eval(e, n->args[i]);
            if (value_is_error(&v)) return v;
            lxlsx_formula_error er; double d = to_number(&v, &er);
            lxlsx_value_free(&v);
            if (er) return val_error(er);
            if (is_and) { if (d == 0.0) result = 0; }
            else { if (d != 0.0) result = 1; }
        }
        return val_bool(result);
    }
    if (streq_ci(fn, "NOT")) {
        if (n->nargs != 1) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) return v;
        lxlsx_formula_error er; double d = to_number(&v, &er);
        lxlsx_value_free(&v);
        if (er) return val_error(er);
        return val_bool(d == 0.0);
    }

    /* Single-number math functions */
    if (streq_ci(fn, "ABS") || streq_ci(fn, "INT") || streq_ci(fn, "SQRT") ||
        streq_ci(fn, "ROUND") || streq_ci(fn, "ROUNDUP") || streq_ci(fn, "ROUNDDOWN") ||
        streq_ci(fn, "MOD") || streq_ci(fn, "POWER")) {
        if (n->nargs < 1) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value a = eval(e, n->args[0]);
        if (value_is_error(&a)) return a;
        lxlsx_formula_error er; double x = to_number(&a, &er); lxlsx_value_free(&a);
        if (er) return val_error(er);
        if (streq_ci(fn, "ABS")) return val_number(fabs(x));
        if (streq_ci(fn, "INT")) return val_number(floor(x));
        if (streq_ci(fn, "SQRT")) return x < 0 ? val_error(LXLSX_FERR_NUM) : val_number(sqrt(x));
        double y = 0.0;
        if (n->nargs >= 2) {
            lxlsx_value b = eval(e, n->args[1]);
            if (value_is_error(&b)) return b;
            y = to_number(&b, &er); lxlsx_value_free(&b);
            if (er) return val_error(er);
        }
        if (streq_ci(fn, "POWER")) return val_number(pow(x, y));
        if (streq_ci(fn, "MOD")) {
            if (y == 0.0) return val_error(LXLSX_FERR_DIV0);
            double m = fmod(x, y); if (m != 0.0 && ((m < 0) != (y < 0))) m += y;
            return val_number(m);
        }
        /* ROUND family: y = digits */
        double f = pow(10.0, y);
        if (streq_ci(fn, "ROUND"))     return val_number(round(x * f) / f);
        if (streq_ci(fn, "ROUNDUP"))   return val_number((x < 0 ? floor(x * f) : ceil(x * f)) / f);
        if (streq_ci(fn, "ROUNDDOWN")) return val_number((x < 0 ? ceil(x * f) : floor(x * f)) / f);
    }

    /* Text functions */
    if (streq_ci(fn, "CONCATENATE") || streq_ci(fn, "CONCAT")) {
        size_t cap = 16, len = 0; char *buf = malloc(cap);
        if (!buf) return val_error(LXLSX_FERR_VALUE);
        buf[0] = '\0';
        for (i = 0; i < n->nargs; i++) {
            lxlsx_value v = eval(e, n->args[i]);
            if (value_is_error(&v)) { free(buf); return v; }
            char *s = to_string(&v); lxlsx_value_free(&v);
            size_t sl = s ? strlen(s) : 0;
            if (len + sl + 1 > cap) { cap = (len + sl + 1) * 2; char *nb = realloc(buf, cap);
                if (!nb) { free(buf); free(s); return val_error(LXLSX_FERR_VALUE); } buf = nb; }
            if (s) { memcpy(buf + len, s, sl); len += sl; buf[len] = '\0'; free(s); }
        }
        return val_string_take(buf);
    }
    if (streq_ci(fn, "LEN")) {
        if (n->nargs != 1) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) return v;
        char *s = to_string(&v); lxlsx_value_free(&v);
        double l = s ? (double)strlen(s) : 0; free(s);
        return val_number(l);
    }
    if (streq_ci(fn, "UPPER") || streq_ci(fn, "LOWER") || streq_ci(fn, "TRIM")) {
        if (n->nargs != 1) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) return v;
        char *s = to_string(&v); lxlsx_value_free(&v);
        if (!s) return val_error(LXLSX_FERR_VALUE);
        if (streq_ci(fn, "UPPER")) { char *q = s; for (; *q; q++) *q = (char)toupper((unsigned char)*q); }
        else if (streq_ci(fn, "LOWER")) { char *q = s; for (; *q; q++) *q = (char)tolower((unsigned char)*q); }
        else { /* TRIM: collapse runs of spaces, strip ends */
            char *r = s, *w = s; int sp = 1;
            for (; *r; r++) { if (*r == ' ') { if (!sp) { *w++ = ' '; sp = 1; } }
                              else { *w++ = *r; sp = 0; } }
            if (w > s && w[-1] == ' ') w--; *w = '\0';
        }
        return val_string_take(s);
    }
    if (streq_ci(fn, "LEFT") || streq_ci(fn, "RIGHT")) {
        if (n->nargs < 1) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) return v;
        char *s = to_string(&v); lxlsx_value_free(&v);
        if (!s) return val_error(LXLSX_FERR_VALUE);
        long k = 1;
        if (n->nargs >= 2) { lxlsx_value b = eval(e, n->args[1]);
            if (value_is_error(&b)) { free(s); return b; }
            lxlsx_formula_error er; k = (long)to_number(&b, &er); lxlsx_value_free(&b);
            if (er) { free(s); return val_error(er); } }
        long sl = (long)strlen(s);
        if (k < 0) { free(s); return val_error(LXLSX_FERR_VALUE); }
        if (k > sl) k = sl;
        lxlsx_value out;
        if (streq_ci(fn, "LEFT")) out = val_string_copy(s, (size_t)k);
        else out = val_string_copy(s + (sl - k), (size_t)k);
        free(s);
        return out;
    }
    if (streq_ci(fn, "MID")) {
        if (n->nargs != 3) return val_error(LXLSX_FERR_VALUE);
        lxlsx_value v = eval(e, n->args[0]);
        if (value_is_error(&v)) return v;
        char *s = to_string(&v); lxlsx_value_free(&v);
        if (!s) return val_error(LXLSX_FERR_VALUE);
        lxlsx_formula_error er;
        lxlsx_value sv = eval(e, n->args[1]); if (value_is_error(&sv)) { free(s); return sv; }
        long start = (long)to_number(&sv, &er); lxlsx_value_free(&sv);
        if (er) { free(s); return val_error(er); }
        lxlsx_value cv = eval(e, n->args[2]); if (value_is_error(&cv)) { free(s); return cv; }
        long cnt = (long)to_number(&cv, &er); lxlsx_value_free(&cv);
        if (er) { free(s); return val_error(er); }
        long sl = (long)strlen(s);
        if (start < 1 || cnt < 0) { free(s); return val_error(LXLSX_FERR_VALUE); }
        long from = start - 1; if (from > sl) from = sl;
        if (from + cnt > sl) cnt = sl - from;
        lxlsx_value out = val_string_copy(s + from, (size_t)cnt);
        free(s); return out;
    }

    return val_error(LXLSX_FERR_NAME);  /* unknown function */
}

static lxlsx_value eval(ev *e, node *n)
{
    if (!n) return val_error(LXLSX_FERR_VALUE);
    switch (n->kind) {
    case N_NUM:  return val_number(n->num);
    case N_STR:  return val_string_copy(n->str ? n->str : "", n->str ? strlen(n->str) : 0);
    case N_BOOL: return val_bool(n->boolean);
    case N_ERR:  return val_error(n->err);
    case N_NAME: return val_error(LXLSX_FERR_NAME);
    case N_REF:  return resolve_cell(e, n->row, n->col);
    case N_RANGE: {
        /* A bare range used as a scalar resolves to its top-left cell. */
        return resolve_cell(e, n->row, n->col);
    }
    case N_FUNC: return eval_func(e, n);
    case N_UNARY: {
        if (strcmp(n->op, "%") == 0) {
            lxlsx_value a = eval(e, n->a);
            if (value_is_error(&a)) return a;
            lxlsx_formula_error er; double d = to_number(&a, &er); lxlsx_value_free(&a);
            return er ? val_error(er) : val_number(d / 100.0);
        }
        lxlsx_value a = eval(e, n->a);
        if (value_is_error(&a)) return a;
        lxlsx_formula_error er; double d = to_number(&a, &er); lxlsx_value_free(&a);
        if (er) return val_error(er);
        return val_number(strcmp(n->op, "-") == 0 ? -d : d);
    }
    case N_BINARY: {
        lxlsx_value a = eval(e, n->a);
        if (value_is_error(&a)) return a;
        lxlsx_value b = eval(e, n->b);
        if (value_is_error(&b)) { lxlsx_value_free(&a); return b; }

        if (strcmp(n->op, "&") == 0) {
            char *sa = to_string(&a), *sb = to_string(&b);
            lxlsx_value_free(&a); lxlsx_value_free(&b);
            size_t la = sa ? strlen(sa) : 0, lb = sb ? strlen(sb) : 0;
            char *r = malloc(la + lb + 1);
            if (!r) { free(sa); free(sb); return val_error(LXLSX_FERR_VALUE); }
            if (sa) memcpy(r, sa, la); if (sb) memcpy(r + la, sb, lb); r[la + lb] = '\0';
            free(sa); free(sb);
            return val_string_take(r);
        }

        /* comparisons */
        if (strcmp(n->op, "=") == 0 || strcmp(n->op, "<>") == 0 ||
            strcmp(n->op, "<") == 0 || strcmp(n->op, ">") == 0 ||
            strcmp(n->op, "<=") == 0 || strcmp(n->op, ">=") == 0) {
            int cmp; int numeric = (a.kind == LXLSX_VAL_NUMBER || a.kind == LXLSX_VAL_BOOL ||
                                    a.kind == LXLSX_VAL_BLANK) &&
                                   (b.kind == LXLSX_VAL_NUMBER || b.kind == LXLSX_VAL_BOOL ||
                                    b.kind == LXLSX_VAL_BLANK);
            if (numeric) {
                lxlsx_formula_error er; double x = to_number(&a, &er), y = to_number(&b, &er);
                lxlsx_value_free(&a); lxlsx_value_free(&b);
                cmp = (x < y) ? -1 : (x > y) ? 1 : 0;
            } else {
                char *sa = to_string(&a), *sb = to_string(&b);
                lxlsx_value_free(&a); lxlsx_value_free(&b);
                cmp = lxlsx_strcasecmp(sa ? sa : "", sb ? sb : "");
                free(sa); free(sb);
            }
            int r = 0;
            if (strcmp(n->op, "=") == 0) r = (cmp == 0);
            else if (strcmp(n->op, "<>") == 0) r = (cmp != 0);
            else if (strcmp(n->op, "<") == 0) r = (cmp < 0);
            else if (strcmp(n->op, ">") == 0) r = (cmp > 0);
            else if (strcmp(n->op, "<=") == 0) r = (cmp <= 0);
            else r = (cmp >= 0);
            return val_bool(r);
        }

        /* arithmetic */
        lxlsx_formula_error er; double x = to_number(&a, &er);
        if (er) { lxlsx_value_free(&a); lxlsx_value_free(&b); return val_error(er); }
        double y = to_number(&b, &er);
        lxlsx_value_free(&a); lxlsx_value_free(&b);
        if (er) return val_error(er);
        if (strcmp(n->op, "+") == 0) return val_number(x + y);
        if (strcmp(n->op, "-") == 0) return val_number(x - y);
        if (strcmp(n->op, "*") == 0) return val_number(x * y);
        if (strcmp(n->op, "/") == 0) return y == 0.0 ? val_error(LXLSX_FERR_DIV0) : val_number(x / y);
        if (strcmp(n->op, "^") == 0) return val_number(pow(x, y));
        return val_error(LXLSX_FERR_VALUE);
    }
    }
    return val_error(LXLSX_FERR_VALUE);
}

/* ===================================================================== *
 * Public entry
 * ===================================================================== */

lxlsx_error lxlsx_formula_eval(const char *formula,
                               lxlsx_formula_resolver resolver, void *ctx,
                               lxlsx_value *out)
{
    pstate p;
    ev e;
    node *root;

    if (!formula || !out)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    *out = val_blank();

    /* Skip a leading '=' if present. */
    if (*formula == '=') formula++;

    p.src = formula;
    p.pos = 0;
    p.n = strlen(formula);
    p.failed = 0;
    lex_next(&p);

    root = parse_expr(&p);
    if (!root || p.failed || p.cur.kind != TK_END) {
        node_free(root);
        *out = val_error(LXLSX_FERR_NAME);  /* unparseable -> surfaced in-band */
        return LXLSX_NO_ERROR;
    }

    e.resolver = resolver;
    e.ctx = ctx;
    *out = eval(&e, root);
    node_free(root);
    return LXLSX_NO_ERROR;
}
