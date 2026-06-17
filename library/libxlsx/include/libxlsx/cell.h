/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

/**
 * @file cell.h
 *
 * @brief Shared worksheet cell types used by writer and reader modules.
 */
#ifndef __LXLSX_CELL_H__
#define __LXLSX_CELL_H__

#include <stddef.h>
#include <stdint.h>

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_format lxlsx_format;
typedef struct lxlsx_vml_obj lxlsx_vml_obj;

typedef struct {
    const char *ptr;
    size_t      len;
} lxlsx_str;

typedef enum cell_types {
    NUMBER_CELL = 1,
    STRING_CELL,
    INLINE_STRING_CELL,
    INLINE_RICH_STRING_CELL,
    FORMULA_CELL,
    ARRAY_FORMULA_CELL,
    DYNAMIC_ARRAY_FORMULA_CELL,
    BLANK_CELL,
    BOOLEAN_CELL,
    ERROR_CELL,
    COMMENT,
    HYPERLINK_URL,
    HYPERLINK_INTERNAL,
    HYPERLINK_EXTERNAL,
    DATETIME_CELL
} lxlsx_cell_type;

typedef enum {
    LXLSX_FORMULA_NORMAL    = 0,
    LXLSX_FORMULA_ARRAY     = 1,
    LXLSX_FORMULA_DATATABLE = 2,
    LXLSX_FORMULA_SHARED    = 3
} lxlsx_formula_kind;

typedef struct {
    lxlsx_str          formula;
    lxlsx_str          cached;
    lxlsx_formula_kind kind;
    lxlsx_str          ref;
    int                si;
    int                is_dynamic;
} lxlsx_cell_formula;

/* Shared worksheet cell representation. Writer fields are preserved for the
 * red-black tree backed write path; reader fields hold parsed cell values. */
typedef struct lxlsx_cell {
    lxlsx_row_t row_num;
    lxlsx_col_t col_num;
    enum cell_types type;
    lxlsx_format *format;
    lxlsx_vml_obj *comment;

    union {
        double number;
        int32_t string_id;
        const char *string;
    } u;

    double formula_result;
    char *user_data1;
    char *user_data2;
    char *lxlsx_sst_string;

    uint32_t style_id;
    uint32_t style_ref;
    lxlsx_str raw;

    union {
        double number;
        int64_t unix_timestamp;
        lxlsx_str string;
        int boolean;
        lxlsx_cell_formula formula;
        char error_code[8];
    } value;

    /* List pointers for tree.h. */
    RB_ENTRY (lxlsx_cell) tree_pointers;
} lxlsx_cell;

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_CELL_H__ */
