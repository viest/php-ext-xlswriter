/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * comment - A libxlsxwriter library for creating Excel XLSX comment files.
 *
 */
#ifndef __LXW_COMMENT_H__
#define __LXW_COMMENT_H__

#include <stdint.h>

#include "common.h"
#include "worksheet.h"

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxw_author_ids, lxw_author_id);

/*
 * Struct to represent a comment object.
 */
typedef struct lxw_comment {

    FILE *file;
    struct lxw_comment_objs *comment_objs;
    struct lxw_author_ids *author_ids;
    char *comment_author;
    uint32_t author_id;

} lxw_comment;

/* Struct to an author id */
typedef struct lxw_author_id {
    uint32_t id;
    char *author;

    RB_ENTRY (lxw_author_id) tree_pointers;
} lxw_author_id;

#define LXW_RB_GENERATE_AUTHOR_IDS(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)    \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)    \
    RB_GENERATE_INSERT(name, type, field, cmp, static)     \
    RB_GENERATE_REMOVE(name, type, field, static)          \
    RB_GENERATE_FIND(name, type, field, cmp, static)       \
    RB_GENERATE_NEXT(name, type, field, static)            \
    RB_GENERATE_MINMAX(name, type, field, static)          \
    /* Add unused struct to allow adding a semicolon */    \
    struct lxw_rb_generate_author_ids{int unused;}


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_comment *lxw_comment_new(void);
void lxw_comment_free(lxw_comment *comment);
void lxw_comment_assemble_xml_file(lxw_comment *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _comment_xml_declaration(lxw_comment *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_COMMENT_H__ */
