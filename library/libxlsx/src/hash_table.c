/*****************************************************************************
 * hash_table - Hash table functions for libxlsxwriter.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include "libxlsx/hash_table.h"

/*
 * Calculate the hash key using the FNV function. See:
 * http://en.wikipedia.org/wiki/Fowler-Noll-Vo_hash_function
 */
STATIC size_t
_generate_hash_key(void *data, size_t data_len, size_t num_buckets)
{
    unsigned char *p = data;
    size_t hash = 2166136261U;
    size_t i;

    for (i = 0; i < data_len; i++)
        hash = (hash * 16777619) ^ p[i];

    return hash % num_buckets;
}

/*
 * Check if an element exists in the hash table and return a pointer
 * to it if it does.
 */
lxlsx_hash_element *
lxlsx_hash_key_exists(lxlsx_hash_table *lxlsx_hash, void *key, size_t key_len)
{
    size_t hash_key = _generate_hash_key(key, key_len, lxlsx_hash->num_buckets);
    struct lxlsx_hash_bucket_list *list;
    lxlsx_hash_element *element;

    if (!lxlsx_hash->buckets[hash_key]) {
        /* The key isn't in the LXLSX_HASH hash table. */
        return NULL;
    }
    else {
        /* The key is already in the table or there is a hash collision. */
        list = lxlsx_hash->buckets[hash_key];

        /* Iterate over the keys in the bucket's linked list. */
        SLIST_FOREACH(element, list, lxlsx_hash_list_pointers) {
            if (memcmp(element->key, key, key_len) == 0) {
                /* The key already exists in the table. */
                return element;
            }
        }

        /* Key doesn't exist in the list so this is a hash collision. */
        return NULL;
    }
}

/*
 * Insert or update a value in the LXLSX_HASH table based on a key
 * and return a pointer to the new or updated element.
 */
lxlsx_hash_element *
lxlsx_insert_hash_element(lxlsx_hash_table *lxlsx_hash, void *key, void *value,
                        size_t key_len)
{
    size_t hash_key = _generate_hash_key(key, key_len, lxlsx_hash->num_buckets);
    struct lxlsx_hash_bucket_list *list = NULL;
    lxlsx_hash_element *element = NULL;

    if (!lxlsx_hash->buckets[hash_key]) {
        /* The key isn't in the LXLSX_HASH hash table. */

        /* Create a linked list in the bucket to hold the lxlsx_hash keys. */
        list = calloc(1, sizeof(struct lxlsx_hash_bucket_list));
        GOTO_LABEL_ON_MEM_ERROR(list, mem_error1);

        /* Initialize the bucket linked list. */
        SLIST_INIT(list);

        /* Create an lxlsx_hash element to add to the linked list. */
        element = calloc(1, sizeof(lxlsx_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(element, mem_error1);

        /* Store the key and value. */
        element->key = key;
        element->value = value;

        /* Add the lxlsx_hash element to the bucket's linked list. */
        SLIST_INSERT_HEAD(list, element, lxlsx_hash_list_pointers);

        /* Also add it to the insertion order linked list. */
        STAILQ_INSERT_TAIL(lxlsx_hash->order_list, element,
                           lxlsx_hash_order_pointers);

        /* Store the bucket list at the hash index. */
        lxlsx_hash->buckets[hash_key] = list;

        lxlsx_hash->used_buckets++;
        lxlsx_hash->unique_count++;

        return element;
    }
    else {
        /* The key is already in the table or there is a hash collision. */
        list = lxlsx_hash->buckets[hash_key];

        /* Iterate over the keys in the bucket's linked list. */
        SLIST_FOREACH(element, list, lxlsx_hash_list_pointers) {
            if (memcmp(element->key, key, key_len) == 0) {
                /* The key already exists in the table. Update the value. */
                if (lxlsx_hash->free_value)
                    free(element->value);

                element->value = value;
                return element;
            }
        }

        /* Key doesn't exist in the list so this is a hash collision.
         * Create an lxlsx_hash element to add to the linked list. */
        element = calloc(1, sizeof(lxlsx_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(element, mem_error2);

        /* Store the key and value. */
        element->key = key;
        element->value = value;

        /* Add the lxlsx_hash element to the bucket linked list. */
        SLIST_INSERT_HEAD(list, element, lxlsx_hash_list_pointers);

        /* Also add it to the insertion order linked list. */
        STAILQ_INSERT_TAIL(lxlsx_hash->order_list, element,
                           lxlsx_hash_order_pointers);

        lxlsx_hash->unique_count++;

        return element;
    }

mem_error1:
    free(list);

mem_error2:
    free(element);
    return NULL;
}

/*
 * Create a new LXLSX_HASH hash table object.
 */
lxlsx_hash_table *
lxlsx_hash_new(uint32_t num_buckets, uint8_t free_key, uint8_t free_value)
{
    /* Create the new hash table. */
    lxlsx_hash_table *lxlsx_hash = calloc(1, sizeof(lxlsx_hash_table));
    RETURN_ON_MEM_ERROR(lxlsx_hash, NULL);

    lxlsx_hash->free_key = free_key;
    lxlsx_hash->free_value = free_value;

    /* Add the lxlsx_hash element buckets. */
    lxlsx_hash->buckets =
        calloc(num_buckets, sizeof(struct lxlsx_hash_bucket_list *));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_hash->buckets, mem_error);

    /* Add a list for tracking the insertion order. */
    lxlsx_hash->order_list = calloc(1, sizeof(struct lxlsx_hash_order_list));
    GOTO_LABEL_ON_MEM_ERROR(lxlsx_hash->order_list, mem_error);

    /* Initialize the order list. */
    STAILQ_INIT(lxlsx_hash->order_list);

    /* Store the number of buckets to calculate the load factor. */
    lxlsx_hash->num_buckets = num_buckets;

    return lxlsx_hash;

mem_error:
    lxlsx_hash_free(lxlsx_hash);
    return NULL;
}

/*
 * Free the LXLSX_HASH hash table object.
 */
void
lxlsx_hash_free(lxlsx_hash_table *lxlsx_hash)
{
    size_t i;
    lxlsx_hash_element *element;
    lxlsx_hash_element *element_temp;

    if (!lxlsx_hash)
        return;

    /* Free the lxlsx_hash_elements and data using the ordered linked list. */
    if (lxlsx_hash->order_list) {
        STAILQ_FOREACH_SAFE(element, lxlsx_hash->order_list,
                            lxlsx_hash_order_pointers, element_temp) {
            if (lxlsx_hash->free_key)
                free(element->key);
            if (lxlsx_hash->free_value)
                free(element->value);
            free(element);
        }
    }

    /* Free the buckets from the hash table. */
    for (i = 0; i < lxlsx_hash->num_buckets; i++) {
        free(lxlsx_hash->buckets[i]);
    }

    free(lxlsx_hash->order_list);
    free(lxlsx_hash->buckets);
    free(lxlsx_hash);
}
