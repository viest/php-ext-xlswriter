/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * hash_table - Hash table functions for libxlsxwriter.
 *
 */

#ifndef __LXW_HASH_TABLE_H__
#define __LXW_HASH_TABLE_H__

#include "common.h"

/* Macro to loop over hash table elements in insertion order. */
#define LXW_FOREACH_ORDERED(elem, hash_table) \
    STAILQ_FOREACH((elem), (hash_table)->order_list, lxw_hash_order_pointers)

/* List declarations. */
STAILQ_HEAD(lxw_hash_order_list, lxw_hash_element);
SLIST_HEAD(lxw_hash_bucket_list, lxw_hash_element);

/* LXW_HASH hash table struct. */
typedef struct lxw_hash_table {
    uint32_t num_buckets;
    uint32_t used_buckets;
    uint32_t unique_count;
    uint8_t free_key;
    uint8_t free_value;

    struct lxw_hash_order_list *order_list;
    struct lxw_hash_bucket_list **buckets;
} lxw_hash_table;

/*
 * LXW_HASH table element struct.
 *
 * The hash elements contain pointers to allow them to be stored in
 * lists in the the hash table buckets and also pointers to track the
 * insertion order in a separate list.
 */
typedef struct lxw_hash_element {
    void *key;
    void *value;

    STAILQ_ENTRY (lxw_hash_element) lxw_hash_order_pointers;
    SLIST_ENTRY (lxw_hash_element) lxw_hash_list_pointers;
} lxw_hash_element;


 /* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_hash_element *lxw_hash_key_exists(lxw_hash_table *lxw_hash, void *key,
                                      size_t key_len);
lxw_hash_element *lxw_insert_hash_element(lxw_hash_table *lxw_hash, void *key,
                                          void *value, size_t key_len);
lxw_hash_table *lxw_hash_new(uint32_t num_buckets, uint8_t free_key,
                             uint8_t free_value);
void lxw_hash_free(lxw_hash_table *lxw_hash);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_HASH_TABLE_H__ */
