/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

/**
 * @file formula.h
 *
 * @brief Excel formula evaluation engine (PHP-agnostic).
 *
 * Parses an Excel formula string into a C AST and evaluates it to a value.
 * Cell references are resolved through a caller-supplied callback so the engine
 * is decoupled from worksheet internals and independently testable.
 */
#ifndef __LXLSX_FORMULA_H__
#define __LXLSX_FORMULA_H__

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

/* Excel error kinds (the worksheet error strings map onto these). */
typedef enum {
    LXLSX_FERR_NONE = 0,
    LXLSX_FERR_NULL,    /* #NULL!  */
    LXLSX_FERR_DIV0,    /* #DIV/0! */
    LXLSX_FERR_VALUE,   /* #VALUE! */
    LXLSX_FERR_REF,     /* #REF!   */
    LXLSX_FERR_NAME,    /* #NAME?  */
    LXLSX_FERR_NUM,     /* #NUM!   */
    LXLSX_FERR_NA       /* #N/A    */
} lxlsx_formula_error;

typedef enum {
    LXLSX_VAL_BLANK = 0,
    LXLSX_VAL_NUMBER,
    LXLSX_VAL_STRING,
    LXLSX_VAL_BOOL,
    LXLSX_VAL_ERROR
} lxlsx_value_kind;

/* A scalar formula value. STRING owns its buffer; free with lxlsx_value_free. */
typedef struct {
    lxlsx_value_kind kind;
    double           number;  /* NUMBER; BOOL stored as 0.0/1.0 */
    char            *string;  /* STRING (owned), else NULL */
    lxlsx_formula_error error; /* ERROR kind, else LXLSX_FERR_NONE */
} lxlsx_value;

/*
 * Resolve the value of a single cell (0-based row/col). Fill *out; set
 * out->kind = LXLSX_VAL_BLANK for empty/missing cells. The callback must not
 * allocate into out->string unless it sets kind = LXLSX_VAL_STRING (the engine
 * takes ownership and frees it).
 */
typedef void (*lxlsx_formula_resolver)(void *ctx,
                                       lxlsx_row_t row, lxlsx_col_t col,
                                       lxlsx_value *out);

/* Release any heap held by a value (the string) and reset it to BLANK. */
void lxlsx_value_free(lxlsx_value *value);

/* Map an error kind to its Excel string ("#DIV/0!", ...). Never NULL. */
const char *lxlsx_formula_error_string(lxlsx_formula_error error);

/*
 * Evaluate an Excel formula (without the leading '='). On success returns
 * LXLSX_NO_ERROR and writes the result to *out (which may itself be an
 * LXLSX_VAL_ERROR value for in-band Excel errors like #DIV/0!). Returns
 * LXLSX_ERROR_* only for engine-level failures (NULL args, OOM, parse error
 * surfaced as #NAME?/#VALUE? in *out). resolver may be NULL if the formula
 * references no cells.
 */
lxlsx_error lxlsx_formula_eval(const char *formula,
                               lxlsx_formula_resolver resolver, void *ctx,
                               lxlsx_value *out);

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_FORMULA_H__ */
