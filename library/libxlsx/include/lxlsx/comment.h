/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * comment - A libxlsxwriter library for creating Excel XLSX comment files.
 *
 */
#ifndef __LXLSX_COMMENT_H__
#define __LXLSX_COMMENT_H__

#include <stdint.h>

#include "common.h"
#include "worksheet.h"

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxlsx_author_ids, lxlsx_author_id);

/*
 * Struct to represent a comment object.
 */
typedef struct lxlsx_comment {

    FILE *file;
    struct lxlsx_comment_objs *comment_objs;
    struct lxlsx_author_ids *author_ids;
    char *comment_author;
    uint32_t author_id;

} lxlsx_comment;

/* Struct to an author id */
typedef struct lxlsx_author_id {
    uint32_t id;
    char *author;

    RB_ENTRY (lxlsx_author_id) tree_pointers;
} lxlsx_author_id;

#define LXLSX_RB_GENERATE_AUTHOR_IDS(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)    \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)    \
    RB_GENERATE_INSERT(name, type, field, cmp, static)     \
    RB_GENERATE_REMOVE(name, type, field, static)          \
    RB_GENERATE_FIND(name, type, field, cmp, static)       \
    RB_GENERATE_NEXT(name, type, field, static)            \
    RB_GENERATE_MINMAX(name, type, field, static)          \
    /* Add unused struct to allow adding a semicolon */    \
    struct lxlsx_rb_generate_author_ids{int unused;}


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_comment *lxlsx_comment_new(void);
void lxlsx_comment_free(lxlsx_comment *comment);
void lxlsx_comment_assemble_xml_file(lxlsx_comment *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _comment_xml_declaration(lxlsx_comment *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_COMMENT_H__ */
