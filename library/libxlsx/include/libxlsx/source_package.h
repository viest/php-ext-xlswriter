/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

/**
 * @file source_package.h
 *
 * @brief Snapshot-based XLSX source package helpers used by edit support.
 */
#ifndef __LXLSX_SOURCE_PACKAGE_H__
#define __LXLSX_SOURCE_PACKAGE_H__

#include <stddef.h>
#include <stdint.h>

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_source_package lxlsx_source_package;

typedef struct {
    const char *name;
    size_t      name_len;
    uint16_t    flags;
    uint16_t    compression_method;
    uint32_t    crc32;
    uint64_t    compressed_size;
    uint64_t    uncompressed_size;
    uint32_t    external_attr;
} lxlsx_source_package_entry_info;

typedef struct {
    size_t               entry_index;
    const unsigned char *data;
    size_t               size;
} lxlsx_source_package_replacement;

/* A brand-new part to append to the package (e.g. a new worksheet). */
typedef struct {
    const char          *name;   /* zip member name, e.g. xl/worksheets/sheet3.xml */
    const unsigned char *data;
    size_t               size;
} lxlsx_source_package_addition;

lxlsx_error lxlsx_source_package_open(const char *path,
                                      lxlsx_source_package **out);
lxlsx_error lxlsx_source_package_open_memory(const void *data, size_t len,
                                             lxlsx_source_package **out);
void        lxlsx_source_package_close(lxlsx_source_package *package);

const unsigned char *lxlsx_source_package_data(const lxlsx_source_package *package,
                                               size_t *len);
size_t lxlsx_source_package_entry_count(const lxlsx_source_package *package);
const lxlsx_source_package_entry_info *
lxlsx_source_package_entry_info_at(const lxlsx_source_package *package,
                                   size_t index);
int lxlsx_source_package_find_first(const lxlsx_source_package *package,
                                    const char *name);

lxlsx_error lxlsx_source_package_read_entry(const lxlsx_source_package *package,
                                            size_t index,
                                            unsigned char **out,
                                            size_t *out_len);
void lxlsx_source_package_free_buffer(void *buffer);

/* Return the untouched local-file record slice for an entry in the original
 * XLSX package. This is intended for tests and diagnostics that need to verify
 * byte-for-byte preservation of entries not replaced by edit output. The
 * returned pointer is owned by the package and remains valid until close. */
lxlsx_error lxlsx_source_package_entry_raw_local_record(
    const lxlsx_source_package *package,
    size_t index,
    const unsigned char **out,
    size_t *out_len);

lxlsx_error lxlsx_source_package_save_copy(const lxlsx_source_package *package,
                                           const char *path);
lxlsx_error lxlsx_source_package_save_with_replacements(
    const lxlsx_source_package *package,
    const char *path,
    const lxlsx_source_package_replacement *replacements,
    size_t replacement_count);

/* Like save_with_replacements but also appends brand-new parts. */
lxlsx_error lxlsx_source_package_save_with_changes(
    const lxlsx_source_package *package,
    const char *path,
    const lxlsx_source_package_replacement *replacements,
    size_t replacement_count,
    const lxlsx_source_package_addition *additions,
    size_t addition_count);

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_SOURCE_PACKAGE_H__ */
