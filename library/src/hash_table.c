/*****************************************************************************
 * hash_table - Hash table functions for libxlsxwriter.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include "xlsxwriter/hash_table.h"

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
lxw_hash_element *
lxw_hash_key_exists(lxw_hash_table *lxw_hash, void *key, size_t key_len)
{
    size_t hash_key = _generate_hash_key(key, key_len, lxw_hash->num_buckets);
    struct lxw_hash_bucket_list *list;
    lxw_hash_element *element;

    if (!lxw_hash->buckets[hash_key]) {
        /* The key isn't in the LXW_HASH hash table. */
        return NULL;
    }
    else {
        /* The key is already in the table or there is a hash collision. */
        list = lxw_hash->buckets[hash_key];

        /* Iterate over the keys in the bucket's linked list. */
        SLIST_FOREACH(element, list, lxw_hash_list_pointers) {
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
 * Insert or update a value in the LXW_HASH table based on a key
 * and return a pointer to the new or updated element.
 */
lxw_hash_element *
lxw_insert_hash_element(lxw_hash_table *lxw_hash, void *key, void *value,
                        size_t key_len)
{
    size_t hash_key = _generate_hash_key(key, key_len, lxw_hash->num_buckets);
    struct lxw_hash_bucket_list *list = NULL;
    lxw_hash_element *element = NULL;

    if (!lxw_hash->buckets[hash_key]) {
        /* The key isn't in the LXW_HASH hash table. */

        /* Create a linked list in the bucket to hold the lxw_hash keys. */
        list = calloc(1, sizeof(struct lxw_hash_bucket_list));
        GOTO_LABEL_ON_MEM_ERROR(list, mem_error1);

        /* Initialize the bucket linked list. */
        SLIST_INIT(list);

        /* Create an lxw_hash element to add to the linked list. */
        element = calloc(1, sizeof(lxw_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(element, mem_error1);

        /* Store the key and value. */
        element->key = key;
        element->value = value;

        /* Add the lxw_hash element to the bucket's linked list. */
        SLIST_INSERT_HEAD(list, element, lxw_hash_list_pointers);

        /* Also add it to the insertion order linked list. */
        STAILQ_INSERT_TAIL(lxw_hash->order_list, element,
                           lxw_hash_order_pointers);

        /* Store the bucket list at the hash index. */
        lxw_hash->buckets[hash_key] = list;

        lxw_hash->used_buckets++;
        lxw_hash->unique_count++;

        return element;
    }
    else {
        /* The key is already in the table or there is a hash collision. */
        list = lxw_hash->buckets[hash_key];

        /* Iterate over the keys in the bucket's linked list. */
        SLIST_FOREACH(element, list, lxw_hash_list_pointers) {
            if (memcmp(element->key, key, key_len) == 0) {
                /* The key already exists in the table. Update the value. */
                if (lxw_hash->free_value)
                    free(element->value);

                element->value = value;
                return element;
            }
        }

        /* Key doesn't exist in the list so this is a hash collision.
         * Create an lxw_hash element to add to the linked list. */
        element = calloc(1, sizeof(lxw_hash_element));
        GOTO_LABEL_ON_MEM_ERROR(element, mem_error2);

        /* Store the key and value. */
        element->key = key;
        element->value = value;

        /* Add the lxw_hash element to the bucket linked list. */
        SLIST_INSERT_HEAD(list, element, lxw_hash_list_pointers);

        /* Also add it to the insertion order linked list. */
        STAILQ_INSERT_TAIL(lxw_hash->order_list, element,
                           lxw_hash_order_pointers);

        lxw_hash->unique_count++;

        return element;
    }

mem_error1:
    free(list);

mem_error2:
    free(element);
    return NULL;
}

/*
 * Create a new LXW_HASH hash table object.
 */
lxw_hash_table *
lxw_hash_new(uint32_t num_buckets, uint8_t free_key, uint8_t free_value)
{
    /* Create the new hash table. */
    lxw_hash_table *lxw_hash = calloc(1, sizeof(lxw_hash_table));
    RETURN_ON_MEM_ERROR(lxw_hash, NULL);

    lxw_hash->free_key = free_key;
    lxw_hash->free_value = free_value;

    /* Add the lxw_hash element buckets. */
    lxw_hash->buckets =
        calloc(num_buckets, sizeof(struct lxw_hash_bucket_list *));
    GOTO_LABEL_ON_MEM_ERROR(lxw_hash->buckets, mem_error);

    /* Add a list for tracking the insertion order. */
    lxw_hash->order_list = calloc(1, sizeof(struct lxw_hash_order_list));
    GOTO_LABEL_ON_MEM_ERROR(lxw_hash->order_list, mem_error);

    /* Initialize the order list. */
    STAILQ_INIT(lxw_hash->order_list);

    /* Store the number of buckets to calculate the load factor. */
    lxw_hash->num_buckets = num_buckets;

    return lxw_hash;

mem_error:
    lxw_hash_free(lxw_hash);
    return NULL;
}

/*
 * Free the LXW_HASH hash table object.
 */
void
lxw_hash_free(lxw_hash_table *lxw_hash)
{
    size_t i;
    lxw_hash_element *element;
    lxw_hash_element *element_temp;

    if (!lxw_hash)
        return;

    /* Free the lxw_hash_elements and data using the ordered linked list. */
    if (lxw_hash->order_list) {
        STAILQ_FOREACH_SAFE(element, lxw_hash->order_list,
                            lxw_hash_order_pointers, element_temp) {
            if (lxw_hash->free_key)
                free(element->key);
            if (lxw_hash->free_value)
                free(element->value);
            free(element);
        }
    }

    /* Free the buckets from the hash table. */
    for (i = 0; i < lxw_hash->num_buckets; i++) {
        free(lxw_hash->buckets[i]);
    }

    free(lxw_hash->order_list);
    free(lxw_hash->buckets);
    free(lxw_hash);
}
