/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * shared_strings - A libxlsxwriter library for creating Excel XLSX
 *                  sst files.
 *
 */
#ifndef __LXLSX_SST_H__
#define __LXLSX_SST_H__

#include <string.h>
#include <stdint.h>

#include "common.h"

/* Define a tree.h RB structure for storing shared strings. */
RB_HEAD(lxlsx_sst_rb_tree, lxlsx_sst_element);

/* Define a queue.h structure for storing shared strings in insertion order. */
STAILQ_HEAD(lxlsx_sst_order_list, lxlsx_sst_element);

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXLSX_RB_GENERATE_ELEMENT(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static) \
    RB_GENERATE_INSERT(name, type, field, cmp, static)  \
    /* Add unused struct to allow adding a semicolon */ \
    struct lxlsx_rb_generate_element{int unused;}

/*
 * Elements of the SST table. They contain pointers to allow them to
 * be stored in a RB tree and also pointers to track the insertion order
 * in a separate list.
 */
struct lxlsx_sst_element {
    uint32_t index;
    char *string;
    uint8_t is_rich_string;

    STAILQ_ENTRY (lxlsx_sst_element) lxlsx_sst_order_pointers;
    RB_ENTRY (lxlsx_sst_element) lxlsx_sst_tree_pointers;
};

/*
 * Struct to represent a sst.
 */
typedef struct lxlsx_sst {
    FILE *file;

    uint32_t string_count;
    uint32_t unique_count;

    struct lxlsx_sst_order_list *order_list;
    struct lxlsx_sst_rb_tree *rb_tree;

} lxlsx_sst;

/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_sst *lxlsx_sst_new(void);
void lxlsx_sst_free(lxlsx_sst *sst);
struct lxlsx_sst_element *lxlsx_get_sst_index(lxlsx_sst *sst, const char *string,
                                      uint8_t is_rich_string);
void lxlsx_sst_assemble_xml_file(lxlsx_sst *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _sst_xml_declaration(lxlsx_sst *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_SST_H__ */
