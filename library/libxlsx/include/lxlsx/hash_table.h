/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * hash_table - Hash table functions for libxlsxwriter.
 *
 */

#ifndef __LXLSX_HASH_TABLE_H__
#define __LXLSX_HASH_TABLE_H__

#include "common.h"

/* Macro to loop over hash table elements in insertion order. */
#define LXLSX_FOREACH_ORDERED(elem, hash_table) \
    STAILQ_FOREACH((elem), (hash_table)->order_list, lxlsx_hash_order_pointers)

/* List declarations. */
STAILQ_HEAD(lxlsx_hash_order_list, lxlsx_hash_element);
SLIST_HEAD(lxlsx_hash_bucket_list, lxlsx_hash_element);

/* LXLSX_HASH hash table struct. */
typedef struct lxlsx_hash_table {
    uint32_t num_buckets;
    uint32_t used_buckets;
    uint32_t unique_count;
    uint8_t free_key;
    uint8_t free_value;

    struct lxlsx_hash_order_list *order_list;
    struct lxlsx_hash_bucket_list **buckets;
} lxlsx_hash_table;

/*
 * LXLSX_HASH table element struct.
 *
 * The hash elements contain pointers to allow them to be stored in
 * lists in the the hash table buckets and also pointers to track the
 * insertion order in a separate list.
 */
typedef struct lxlsx_hash_element {
    void *key;
    void *value;

    STAILQ_ENTRY (lxlsx_hash_element) lxlsx_hash_order_pointers;
    SLIST_ENTRY (lxlsx_hash_element) lxlsx_hash_list_pointers;
} lxlsx_hash_element;


 /* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_hash_element *lxlsx_hash_key_exists(lxlsx_hash_table *lxlsx_hash, void *key,
                                      size_t key_len);
lxlsx_hash_element *lxlsx_insert_hash_element(lxlsx_hash_table *lxlsx_hash, void *key,
                                          void *value, size_t key_len);
lxlsx_hash_table *lxlsx_hash_new(uint32_t num_buckets, uint8_t free_key,
                             uint8_t free_value);
void lxlsx_hash_free(lxlsx_hash_table *lxlsx_hash);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_HASH_TABLE_H__ */
