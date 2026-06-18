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

typedef struct {
    const char *formula;
    double      result;
    const char *range;
    const char *result_string;
} lxlsx_cell_writer_formula;

typedef struct {
    int32_t     id;
    const char *string;
} lxlsx_cell_writer_shared_string;

typedef struct {
    const char *url;
    const char *display;
    const char *tooltip;
} lxlsx_cell_writer_hyperlink;

typedef union {
    double                          number;
    int                             boolean;
    uint32_t                        object_id;
    const char                     *string;
    lxlsx_vml_obj                  *comment;
    lxlsx_cell_writer_formula       formula;
    lxlsx_cell_writer_shared_string shared_string;
    lxlsx_cell_writer_hyperlink     hyperlink;
} lxlsx_cell_writer_value;

typedef struct {
    lxlsx_format *format;
    lxlsx_cell_writer_value value;
    RB_ENTRY (lxlsx_cell) tree_pointers;
} lxlsx_cell_writer_data;

typedef union {
    double                    number;
    int64_t                   unix_timestamp;
    lxlsx_str                 string;
    int                       boolean;
    const lxlsx_cell_formula *formula;
    char                      error_code[8];
} lxlsx_cell_reader_value;

typedef struct {
    uint32_t                style_id;
    uint32_t                style_ref;
    lxlsx_str              raw;
    lxlsx_cell_reader_value value;
} lxlsx_cell_reader_data;

/* Shared worksheet cell representation. Writer and reader payloads are
 * mutually exclusive: writer cells are retained in worksheet RB trees while
 * reader cells are short-lived callback/iterator values. Keeping them in a
 * union avoids charging every writer cell for reader-only formula/raw fields. */
typedef struct lxlsx_cell {
    lxlsx_row_t row_num;
    lxlsx_col_t col_num;
    uint8_t type;
    union {
        lxlsx_cell_writer_data writer;
        lxlsx_cell_reader_data reader;
    } data;
} lxlsx_cell;

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_CELL_H__ */
