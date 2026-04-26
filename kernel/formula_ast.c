/*
  +----------------------------------------------------------------------+
  | XlsWriter Extension — Excel formula AST builder.                      |
  +----------------------------------------------------------------------+
  | Tokenises and parses an Excel formula string into a recursive zval    |
  | tree shaped like {kind: 'literal'|'ref'|'op'|'fn'|'array', ...}.      |
  | NO evaluation — by design (see plans/upgrade.md §10.1: never built).  |
  | Operator precedence follows Excel's: : (range)  ,  unary -+, %, ^,    |
  | * /, + -, &, comparisons. Implementation is recursive descent.        |
  +----------------------------------------------------------------------+
*/

#include "xlswriter.h"

/* --- token kinds --------------------------------------------------------- */
enum {
    TK_END = 0,
    TK_NUMBER, TK_STRING, TK_BOOL, TK_ERROR,
    TK_NAME,           /* function name or identifier (when followed by ( ) */
    TK_REF,            /* cell or range ref token (after identification) */
    TK_LPAREN, TK_RPAREN,
    TK_LBRACE, TK_RBRACE,
    TK_COMMA, TK_SEMI,
    TK_OP              /* +-*\/^&=<>%: */
};

typedef struct {
    int   kind;
    const char *start;
    size_t      len;
    /* For TK_OP, op[0..opn-1] holds the operator characters. */
    char  op[3];
    int   opn;
} tok;

typedef struct {
    const char *src;
    size_t      pos;
    size_t      n;
    tok         cur;
} pstate;

/* skip whitespace */
static void skipws(pstate *p) {
    while (p->pos < p->n) {
        char c = p->src[p->pos];
        if (c == ' ' || c == '\t' || c == '\n' || c == '\r') p->pos++;
        else break;
    }
}

/* Is this a valid identifier-start (letter or underscore). */
static int is_idstart(char c) {
    return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c == '$';
}
static int is_idcont(char c) {
    return is_idstart(c) || (c >= '0' && c <= '9') || c == '.' || c == '!';
}

/* --- lexer --------------------------------------------------------------- */
static void next_tok(pstate *p) {
    skipws(p);
    p->cur.start = p->src + p->pos;
    p->cur.len = 0;
    p->cur.opn = 0;
    if (p->pos >= p->n) { p->cur.kind = TK_END; return; }
    char c = p->src[p->pos];

    if (c == '(') { p->cur.kind = TK_LPAREN; p->pos++; p->cur.len = 1; return; }
    if (c == ')') { p->cur.kind = TK_RPAREN; p->pos++; p->cur.len = 1; return; }
    if (c == '{') { p->cur.kind = TK_LBRACE; p->pos++; p->cur.len = 1; return; }
    if (c == '}') { p->cur.kind = TK_RBRACE; p->pos++; p->cur.len = 1; return; }
    if (c == ',') { p->cur.kind = TK_COMMA;  p->pos++; p->cur.len = 1; return; }
    if (c == ';') { p->cur.kind = TK_SEMI;   p->pos++; p->cur.len = 1; return; }

    /* string literal */
    if (c == '"') {
        size_t s = p->pos++;
        while (p->pos < p->n) {
            if (p->src[p->pos] == '"') {
                /* "" is an embedded quote */
                if (p->pos + 1 < p->n && p->src[p->pos + 1] == '"') p->pos += 2;
                else { p->pos++; break; }
            } else {
                p->pos++;
            }
        }
        p->cur.kind  = TK_STRING;
        p->cur.start = p->src + s;
        p->cur.len   = p->pos - s;
        return;
    }
    /* error literal #NAME?, #DIV/0!, #N/A, etc. */
    if (c == '#') {
        size_t s = p->pos++;
        while (p->pos < p->n) {
            char x = p->src[p->pos];
            if (x == '!' || x == '?') { p->pos++; break; }
            if (x == ',' || x == ')' || x == ' ' || x == '\t') break;
            p->pos++;
        }
        p->cur.kind  = TK_ERROR;
        p->cur.start = p->src + s;
        p->cur.len   = p->pos - s;
        return;
    }
    /* number */
    if ((c >= '0' && c <= '9') ||
        (c == '.' && p->pos + 1 < p->n && p->src[p->pos + 1] >= '0' && p->src[p->pos + 1] <= '9')) {
        size_t s = p->pos++;
        int saw_dot = (c == '.');
        while (p->pos < p->n) {
            char x = p->src[p->pos];
            if (x >= '0' && x <= '9') p->pos++;
            else if (x == '.' && !saw_dot) { saw_dot = 1; p->pos++; }
            else if ((x == 'e' || x == 'E') &&
                     p->pos + 1 < p->n &&
                     (p->src[p->pos + 1] == '+' || p->src[p->pos + 1] == '-' ||
                      (p->src[p->pos + 1] >= '0' && p->src[p->pos + 1] <= '9'))) {
                p->pos++;
                if (p->src[p->pos] == '+' || p->src[p->pos] == '-') p->pos++;
            }
            else break;
        }
        p->cur.kind  = TK_NUMBER;
        p->cur.start = p->src + s;
        p->cur.len   = p->pos - s;
        return;
    }
    /* sheet-qualified ref: 'Name with space'!… or Name!… falls into NAME path */
    if (c == '\'') {
        size_t s = p->pos++;
        while (p->pos < p->n) {
            if (p->src[p->pos] == '\'') {
                if (p->pos + 1 < p->n && p->src[p->pos + 1] == '\'') p->pos += 2;
                else { p->pos++; break; }
            } else p->pos++;
        }
        /* Expect '!' next, fold into NAME until the '!' is consumed and the
         * following identifier starts. Lex the qualified ref as a single
         * NAME token. */
        if (p->pos < p->n && p->src[p->pos] == '!') {
            p->pos++;
            while (p->pos < p->n && is_idcont(p->src[p->pos])) p->pos++;
        }
        p->cur.kind  = TK_NAME;
        p->cur.start = p->src + s;
        p->cur.len   = p->pos - s;
        return;
    }
    /* operators */
    if (c == '+' || c == '-' || c == '*' || c == '/' || c == '^' ||
        c == '&' || c == '=' || c == '<' || c == '>' || c == '%' || c == ':') {
        p->cur.op[0] = c; p->cur.opn = 1; p->pos++;
        if ((c == '<' || c == '>') && p->pos < p->n) {
            if (p->src[p->pos] == '=' || (c == '<' && p->src[p->pos] == '>')) {
                p->cur.op[1] = p->src[p->pos]; p->cur.opn = 2; p->pos++;
            }
        }
        p->cur.kind = TK_OP;
        p->cur.len  = (size_t)p->cur.opn;
        return;
    }
    /* identifier / name / ref / boolean */
    if (is_idstart(c)) {
        size_t s = p->pos++;
        while (p->pos < p->n && is_idcont(p->src[p->pos])) p->pos++;
        p->cur.start = p->src + s;
        p->cur.len   = p->pos - s;
        /* boolean? */
        if ((p->cur.len == 4 &&
             (p->cur.start[0] == 'T' || p->cur.start[0] == 't') &&
             (p->cur.start[1] == 'R' || p->cur.start[1] == 'r') &&
             (p->cur.start[2] == 'U' || p->cur.start[2] == 'u') &&
             (p->cur.start[3] == 'E' || p->cur.start[3] == 'e')) ||
            (p->cur.len == 5 &&
             (p->cur.start[0] == 'F' || p->cur.start[0] == 'f') &&
             (p->cur.start[1] == 'A' || p->cur.start[1] == 'a') &&
             (p->cur.start[2] == 'L' || p->cur.start[2] == 'l') &&
             (p->cur.start[3] == 'S' || p->cur.start[3] == 's') &&
             (p->cur.start[4] == 'E' || p->cur.start[4] == 'e'))) {
            /* But only if not followed by '(' — Excel has no TRUE() besides
             * the function form, which is rare. Treat as bool either way
             * if next char isn't '('. */
            size_t save = p->pos;
            skipws(p);
            if (p->pos >= p->n || p->src[p->pos] != '(') {
                p->pos = save;
                p->cur.kind = TK_BOOL;
                return;
            }
            p->pos = save;
        }
        p->cur.kind = TK_NAME;
        return;
    }
    /* unknown — single char as op */
    p->cur.kind = TK_OP;
    p->cur.op[0] = c;
    p->cur.opn = 1;
    p->cur.len = 1;
    p->pos++;
}

/* --- parser -------------------------------------------------------------- */

/* Forward */
static void parse_expr(pstate *p, zval *out);

static void make_str_zval(zval *out, const char *s, size_t n) {
    ZVAL_STRINGL(out, s, (int)n);
}

/* Decide whether a NAME token is a reference (cell/range) or a true name
 * needing function-call status. We classify when we *know* there's no '('
 * following: rule of thumb is that it parses as a reference if it contains
 * '!' or matches /\$?[A-Z]+\$?[0-9]+(:.*)?/ — otherwise it's just a named
 * range, also surfaced as 'ref'. */
static int looks_like_ref_token(const char *s, size_t n) {
    size_t i;
    int has_alpha = 0, has_digit = 0;
    for (i = 0; i < n; i++) {
        char c = s[i];
        if (c == '!' || c == ':') return 1;
        if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) has_alpha = 1;
        if (c >= '0' && c <= '9') has_digit = 1;
    }
    return (has_alpha && has_digit);
}

/* atom: literal | (expr) | name | name(args) | {array} */
static void parse_atom(pstate *p, zval *out) {
    if (p->cur.kind == TK_NUMBER) {
        zval kind, val;
        char buf[64];
        size_t n = p->cur.len;
        if (n >= sizeof(buf)) n = sizeof(buf) - 1;
        memcpy(buf, p->cur.start, n);
        buf[n] = 0;
        array_init(out);
        ZVAL_STRING(&kind, "literal");
        add_assoc_zval(out, "kind", &kind);
        add_assoc_string(out, "subtype", "number");
        ZVAL_DOUBLE(&val, strtod(buf, NULL));
        add_assoc_zval(out, "value", &val);
        next_tok(p);
        return;
    }
    if (p->cur.kind == TK_STRING) {
        const char *s = p->cur.start + 1;       /* strip leading " */
        size_t      n = p->cur.len   - 2;       /* and trailing " */
        zval kind, val;
        array_init(out);
        ZVAL_STRING(&kind, "literal");
        add_assoc_zval(out, "kind", &kind);
        add_assoc_string(out, "subtype", "string");
        make_str_zval(&val, s, n);
        add_assoc_zval(out, "value", &val);
        next_tok(p);
        return;
    }
    if (p->cur.kind == TK_BOOL) {
        zval kind, val;
        array_init(out);
        ZVAL_STRING(&kind, "literal");
        add_assoc_zval(out, "kind", &kind);
        add_assoc_string(out, "subtype", "bool");
        ZVAL_BOOL(&val, p->cur.start[0] == 'T' || p->cur.start[0] == 't');
        add_assoc_zval(out, "value", &val);
        next_tok(p);
        return;
    }
    if (p->cur.kind == TK_ERROR) {
        zval kind, val;
        array_init(out);
        ZVAL_STRING(&kind, "literal");
        add_assoc_zval(out, "kind", &kind);
        add_assoc_string(out, "subtype", "error");
        make_str_zval(&val, p->cur.start, p->cur.len);
        add_assoc_zval(out, "value", &val);
        next_tok(p);
        return;
    }
    if (p->cur.kind == TK_LPAREN) {
        next_tok(p);
        parse_expr(p, out);
        if (p->cur.kind == TK_RPAREN) next_tok(p);
        return;
    }
    if (p->cur.kind == TK_LBRACE) {
        /* {1,2;3,4}: array literal — semicolons separate rows. */
        zval kind, rows, row;
        array_init(out);
        ZVAL_STRING(&kind, "array");
        add_assoc_zval(out, "kind", &kind);
        array_init(&rows);
        array_init(&row);
        next_tok(p);
        while (p->cur.kind != TK_RBRACE && p->cur.kind != TK_END) {
            zval item;
            parse_expr(p, &item);
            add_next_index_zval(&row, &item);
            if (p->cur.kind == TK_COMMA) {
                next_tok(p);
            } else if (p->cur.kind == TK_SEMI) {
                add_next_index_zval(&rows, &row);
                array_init(&row);
                next_tok(p);
            } else {
                break;
            }
        }
        add_next_index_zval(&rows, &row);
        add_assoc_zval(out, "rows", &rows);
        if (p->cur.kind == TK_RBRACE) next_tok(p);
        return;
    }
    if (p->cur.kind == TK_OP && p->cur.opn == 1 &&
        (p->cur.op[0] == '+' || p->cur.op[0] == '-')) {
        zval kind, op, operand;
        char un = p->cur.op[0];
        next_tok(p);
        parse_atom(p, &operand);  /* unary binds tighter than * / etc. */
        array_init(out);
        ZVAL_STRING(&kind, "op");
        add_assoc_zval(out, "kind", &kind);
        char buf[2] = { un, 0 };
        ZVAL_STRINGL(&op, buf, 1);
        add_assoc_zval(out, "op", &op);
        add_assoc_string(out, "arity", "unary");
        zval args; array_init(&args);
        add_next_index_zval(&args, &operand);
        add_assoc_zval(out, "args", &args);
        return;
    }
    if (p->cur.kind == TK_NAME) {
        const char *name_s = p->cur.start;
        size_t      name_n = p->cur.len;
        next_tok(p);
        if (p->cur.kind == TK_LPAREN) {
            /* function call */
            zval kind, name_z, args;
            array_init(out);
            ZVAL_STRING(&kind, "fn");
            add_assoc_zval(out, "kind", &kind);
            make_str_zval(&name_z, name_s, name_n);
            add_assoc_zval(out, "name", &name_z);
            array_init(&args);
            next_tok(p);
            if (p->cur.kind != TK_RPAREN) {
                while (1) {
                    zval arg;
                    parse_expr(p, &arg);
                    add_next_index_zval(&args, &arg);
                    if (p->cur.kind == TK_COMMA || p->cur.kind == TK_SEMI) {
                        next_tok(p);
                    } else break;
                }
            }
            if (p->cur.kind == TK_RPAREN) next_tok(p);
            add_assoc_zval(out, "args", &args);
            return;
        }
        /* Bare name: treat as ref (cell, range, or named range). */
        {
            zval kind, name_z;
            array_init(out);
            (void)looks_like_ref_token(name_s, name_n);  /* informational */
            ZVAL_STRING(&kind, "ref");
            add_assoc_zval(out, "kind", &kind);
            make_str_zval(&name_z, name_s, name_n);
            add_assoc_zval(out, "text", &name_z);
            return;
        }
    }
    /* Fallback: treat unknown as literal text. */
    {
        zval kind, val;
        array_init(out);
        ZVAL_STRING(&kind, "literal");
        add_assoc_zval(out, "kind", &kind);
        add_assoc_string(out, "subtype", "unknown");
        make_str_zval(&val, p->cur.start, p->cur.len);
        add_assoc_zval(out, "value", &val);
        if (p->cur.kind != TK_END) next_tok(p);
    }
}

/* operator precedence (higher = tighter binding) */
static int precedence(const char *op, int opn) {
    if (opn == 1 && op[0] == ':')             return 9;
    if (opn == 1 && op[0] == '%')             return 8;
    if (opn == 1 && op[0] == '^')             return 7;
    if (opn == 1 && (op[0] == '*' || op[0] == '/')) return 6;
    if (opn == 1 && (op[0] == '+' || op[0] == '-')) return 5;
    if (opn == 1 && op[0] == '&')             return 4;
    if ((opn == 1 && (op[0] == '=' || op[0] == '<' || op[0] == '>')) ||
        (opn == 2 && (op[0] == '<' || op[0] == '>')))                  return 3;
    return 0;
}

static int is_binary_op(const tok *t) {
    if (t->kind != TK_OP) return 0;
    if (t->opn == 1 && t->op[0] == '%') return 0;  /* % is postfix */
    return precedence(t->op, t->opn) > 0;
}

/* Climb operator precedences. */
static void parse_binop(pstate *p, zval *out, int min_prec) {
    parse_atom(p, out);

    /* postfix '%' */
    while (p->cur.kind == TK_OP && p->cur.opn == 1 && p->cur.op[0] == '%') {
        zval kind, op, args, wrap;
        array_init(&wrap);
        ZVAL_STRING(&kind, "op");
        add_assoc_zval(&wrap, "kind", &kind);
        ZVAL_STRING(&op, "%");
        add_assoc_zval(&wrap, "op", &op);
        add_assoc_string(&wrap, "arity", "postfix");
        array_init(&args);
        add_next_index_zval(&args, out);
        add_assoc_zval(&wrap, "args", &args);
        ZVAL_COPY_VALUE(out, &wrap);
        next_tok(p);
    }

    while (is_binary_op(&p->cur)) {
        int  prec = precedence(p->cur.op, p->cur.opn);
        char saved_op[3];
        int  saved_opn;
        zval rhs, wrap, kind, op_z, args;

        if (prec < min_prec) break;
        memcpy(saved_op, p->cur.op, 3);
        saved_opn = p->cur.opn;
        next_tok(p);
        parse_binop(p, &rhs, prec + 1);

        array_init(&wrap);
        ZVAL_STRING(&kind, "op");
        add_assoc_zval(&wrap, "kind", &kind);
        ZVAL_STRINGL(&op_z, saved_op, saved_opn);
        add_assoc_zval(&wrap, "op", &op_z);
        add_assoc_string(&wrap, "arity", "binary");
        array_init(&args);
        add_next_index_zval(&args, out);
        add_next_index_zval(&args, &rhs);
        add_assoc_zval(&wrap, "args", &args);
        ZVAL_COPY_VALUE(out, &wrap);
    }
}

static void parse_expr(pstate *p, zval *out) {
    parse_binop(p, out, 1);
}

/* Public entry: takes a formula string (with or without leading '=') and
 * fills `return_value` with the AST tree. */
void formula_ast_parse(const char *src, size_t n, zval *return_value)
{
    pstate p;
    if (src && n > 0 && src[0] == '=') { src++; n--; }
    p.src = src; p.pos = 0; p.n = n;
    next_tok(&p);
    parse_expr(&p, return_value);
}
