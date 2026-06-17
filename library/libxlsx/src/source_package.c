/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <limits.h>

#include <zlib.h>

#include "libxlsx/source_package.h"

#define ZIP_LOCAL_FILE_SIG   0x04034b50u
#define ZIP_CENTRAL_FILE_SIG 0x02014b50u
#define ZIP_EOCD_SIG         0x06054b50u
#define ZIP64_LOCATOR_SIG    0x07064b50u

#define ZIP_METHOD_STORE     0
#define ZIP_METHOD_DEFLATE   8

typedef struct {
    lxlsx_source_package_entry_info info;

    uint16_t version_made_by;
    uint16_t version_needed;
    uint16_t mod_time;
    uint16_t mod_date;
    uint16_t disk_start;
    uint16_t internal_attr;

    size_t central_offset;
    size_t central_size;
    size_t central_name_offset;
    size_t central_extra_offset;
    size_t central_extra_len;
    size_t central_comment_offset;
    size_t central_comment_len;

    size_t local_offset;
    size_t local_record_size;
    size_t data_offset;
} lxlsx_source_entry;

struct lxlsx_source_package {
    unsigned char *data;
    size_t         size;
    lxlsx_source_entry *entries;
    size_t         entry_count;
    size_t         central_offset;
    size_t         central_size;
    size_t         eocd_offset;
    size_t         global_comment_offset;
    size_t         global_comment_len;
};

static uint16_t read_le16(const unsigned char *p)
{
    return (uint16_t)((uint16_t)p[0] | ((uint16_t)p[1] << 8));
}

static uint32_t read_le32(const unsigned char *p)
{
    return (uint32_t)p[0]
        | ((uint32_t)p[1] << 8)
        | ((uint32_t)p[2] << 16)
        | ((uint32_t)p[3] << 24);
}

static void put_le16(unsigned char *p, uint16_t v)
{
    p[0] = (unsigned char)(v & 0xff);
    p[1] = (unsigned char)((v >> 8) & 0xff);
}

static void put_le32(unsigned char *p, uint32_t v)
{
    p[0] = (unsigned char)(v & 0xff);
    p[1] = (unsigned char)((v >> 8) & 0xff);
    p[2] = (unsigned char)((v >> 16) & 0xff);
    p[3] = (unsigned char)((v >> 24) & 0xff);
}

static int has_zip64_extra(const unsigned char *extra, size_t len)
{
    size_t pos = 0;
    while (pos + 4 <= len) {
        uint16_t header_id = read_le16(extra + pos);
        uint16_t data_size = read_le16(extra + pos + 2);
        pos += 4;
        if (pos + data_size > len)
            return 1;
        if (header_id == 0x0001)
            return 1;
        pos += data_size;
    }
    return pos != len;
}

static lxlsx_error write_all(FILE *fp, const void *buf, size_t len)
{
    if (len == 0)
        return LXLSX_NO_ERROR;
    return fwrite(buf, 1, len, fp) == len
        ? LXLSX_NO_ERROR
        : LXLSX_ERROR_ZIP_FILE_OPERATION;
}

static lxlsx_error write_le16(FILE *fp, uint16_t v)
{
    unsigned char buf[2];
    put_le16(buf, v);
    return write_all(fp, buf, sizeof(buf));
}

static lxlsx_error write_le32(FILE *fp, uint32_t v)
{
    unsigned char buf[4];
    put_le32(buf, v);
    return write_all(fp, buf, sizeof(buf));
}

static lxlsx_error current_offset(FILE *fp, uint32_t *out)
{
    long pos = ftell(fp);
    if (pos < 0)
        return LXLSX_ERROR_ZIP_FILE_OPERATION;
    if ((unsigned long)pos > 0xffffffffUL)
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    *out = (uint32_t)pos;
    return LXLSX_NO_ERROR;
}

static void free_entries(lxlsx_source_package *package)
{
    size_t i;
    if (!package || !package->entries)
        return;
    for (i = 0; i < package->entry_count; i++)
        free((void *)package->entries[i].info.name);
    free(package->entries);
    package->entries = NULL;
    package->entry_count = 0;
}

static lxlsx_error find_eocd(const unsigned char *data, size_t len, size_t *out)
{
    size_t min_pos;
    size_t pos;

    if (!data || len < 22)
        return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

    min_pos = len > (22 + 0xffff) ? len - (22 + 0xffff) : 0;
    pos = len - 22;
    for (;;) {
        if (read_le32(data + pos) == ZIP_EOCD_SIG) {
            uint16_t comment_len = read_le16(data + pos + 20);
            if (pos + 22 + comment_len == len) {
                *out = pos;
                return LXLSX_NO_ERROR;
            }
        }
        if (pos == min_pos)
            break;
        pos--;
    }

    return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
}

static lxlsx_error parse_local_headers(lxlsx_source_package *package)
{
    size_t i;
    size_t j;

    for (i = 0; i < package->entry_count; i++) {
        lxlsx_source_entry *e = &package->entries[i];
        size_t local_offset = e->local_offset;
        size_t name_len;
        size_t extra_len;
        size_t data_offset;
        size_t next_offset = package->central_offset;

        if (local_offset + 30 > package->size)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        if (read_le32(package->data + local_offset) != ZIP_LOCAL_FILE_SIG)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

        name_len = read_le16(package->data + local_offset + 26);
        extra_len = read_le16(package->data + local_offset + 28);
        data_offset = local_offset + 30 + name_len + extra_len;
        if (data_offset > package->size)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        if (data_offset + (size_t)e->info.compressed_size > package->size)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

        for (j = 0; j < package->entry_count; j++) {
            size_t candidate = package->entries[j].local_offset;
            if (candidate > local_offset && candidate < next_offset)
                next_offset = candidate;
        }
        if (next_offset < data_offset + (size_t)e->info.compressed_size)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

        e->data_offset = data_offset;
        e->local_record_size = next_offset - local_offset;
    }

    return LXLSX_NO_ERROR;
}

static lxlsx_error parse_central_directory(lxlsx_source_package *package)
{
    const unsigned char *data = package->data;
    size_t pos;
    size_t i;
    lxlsx_error err;

    if (package->central_offset + package->central_size > package->size)
        return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

    pos = package->central_offset;
    for (i = 0; i < package->entry_count; i++) {
        lxlsx_source_entry *e = &package->entries[i];
        uint16_t name_len;
        uint16_t extra_len;
        uint16_t comment_len;
        uint32_t compressed_size;
        uint32_t uncompressed_size;
        uint32_t local_offset;
        char *name;

        if (pos + 46 > package->central_offset + package->central_size)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        if (read_le32(data + pos) != ZIP_CENTRAL_FILE_SIG)
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

        name_len = read_le16(data + pos + 28);
        extra_len = read_le16(data + pos + 30);
        comment_len = read_le16(data + pos + 32);
        if (pos + 46 + name_len + extra_len + comment_len
            > package->central_offset + package->central_size) {
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        }

        compressed_size = read_le32(data + pos + 20);
        uncompressed_size = read_le32(data + pos + 24);
        local_offset = read_le32(data + pos + 42);
        if (compressed_size == 0xffffffffu ||
            uncompressed_size == 0xffffffffu ||
            local_offset == 0xffffffffu) {
            return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
        }
        if (has_zip64_extra(data + pos + 46 + name_len, extra_len))
            return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

        name = (char *)malloc((size_t)name_len + 1);
        if (!name)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
        memcpy(name, data + pos + 46, name_len);
        name[name_len] = 0;

        e->info.name = name;
        e->info.name_len = name_len;
        e->info.flags = read_le16(data + pos + 8);
        e->info.compression_method = read_le16(data + pos + 10);
        e->info.crc32 = read_le32(data + pos + 16);
        e->info.compressed_size = compressed_size;
        e->info.uncompressed_size = uncompressed_size;
        e->info.external_attr = read_le32(data + pos + 38);
        e->version_made_by = read_le16(data + pos + 4);
        e->version_needed = read_le16(data + pos + 6);
        e->mod_time = read_le16(data + pos + 12);
        e->mod_date = read_le16(data + pos + 14);
        e->disk_start = read_le16(data + pos + 34);
        e->internal_attr = read_le16(data + pos + 36);
        e->central_offset = pos;
        e->central_size = 46 + name_len + extra_len + comment_len;
        e->central_name_offset = pos + 46;
        e->central_extra_offset = pos + 46 + name_len;
        e->central_extra_len = extra_len;
        e->central_comment_offset = pos + 46 + name_len + extra_len;
        e->central_comment_len = comment_len;
        e->local_offset = local_offset;

        if (e->disk_start != 0)
            return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
        if (e->info.flags & 0x0001)
            return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

        pos += e->central_size;
    }

    if (pos != package->central_offset + package->central_size)
        return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;

    err = parse_local_headers(package);
    if (err != LXLSX_NO_ERROR)
        return err;

    return LXLSX_NO_ERROR;
}

static lxlsx_error parse_package(lxlsx_source_package *package)
{
    const unsigned char *eocd;
    uint16_t disk_no;
    uint16_t central_disk;
    uint16_t entries_disk;
    uint16_t entries_total;
    uint32_t central_size;
    uint32_t central_offset;
    lxlsx_error err;

    err = find_eocd(package->data, package->size, &package->eocd_offset);
    if (err != LXLSX_NO_ERROR)
        return err;

    if (package->eocd_offset >= 20 &&
        read_le32(package->data + package->eocd_offset - 20) == ZIP64_LOCATOR_SIG) {
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    }

    eocd = package->data + package->eocd_offset;
    disk_no = read_le16(eocd + 4);
    central_disk = read_le16(eocd + 6);
    entries_disk = read_le16(eocd + 8);
    entries_total = read_le16(eocd + 10);
    central_size = read_le32(eocd + 12);
    central_offset = read_le32(eocd + 16);

    if (disk_no != 0 || central_disk != 0 || entries_disk != entries_total)
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    if (entries_total == 0xffff || central_size == 0xffffffffu ||
        central_offset == 0xffffffffu) {
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    }

    package->central_offset = central_offset;
    package->central_size = central_size;
    package->entry_count = entries_total;
    package->global_comment_offset = package->eocd_offset + 22;
    package->global_comment_len = read_le16(eocd + 20);

    if (package->entry_count > 0) {
        package->entries = (lxlsx_source_entry *)calloc(package->entry_count,
                                                       sizeof(*package->entries));
        if (!package->entries)
            return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    return parse_central_directory(package);
}

static lxlsx_error slurp_file(const char *path, unsigned char **out, size_t *out_len)
{
    FILE *fp;
    long size;
    unsigned char *buf;
    size_t got;

    if (!path || !out || !out_len)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;

    fp = fopen(path, "rb");
    if (!fp)
        return LXLSX_ERROR_CREATING_XLSX_FILE;

    if (fseek(fp, 0, SEEK_END) != 0) {
        fclose(fp);
        return LXLSX_ERROR_ZIP_FILE_OPERATION;
    }
    size = ftell(fp);
    if (size < 0) {
        fclose(fp);
        return LXLSX_ERROR_ZIP_FILE_OPERATION;
    }
    if (fseek(fp, 0, SEEK_SET) != 0) {
        fclose(fp);
        return LXLSX_ERROR_ZIP_FILE_OPERATION;
    }

    buf = (unsigned char *)malloc((size_t)size);
    if (!buf && size != 0) {
        fclose(fp);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }
    got = fread(buf, 1, (size_t)size, fp);
    fclose(fp);
    if (got != (size_t)size) {
        free(buf);
        return LXLSX_ERROR_ZIP_FILE_OPERATION;
    }

    *out = buf;
    *out_len = got;
    return LXLSX_NO_ERROR;
}

static lxlsx_error source_package_open_memory_owned(unsigned char *data,
                                                    size_t len,
                                                    lxlsx_source_package **out)
{
    lxlsx_source_package *package;
    lxlsx_error err;

    package = (lxlsx_source_package *)calloc(1, sizeof(*package));
    if (!package) {
        free(data);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    package->data = data;
    package->size = len;

    err = parse_package(package);
    if (err != LXLSX_NO_ERROR) {
        lxlsx_source_package_close(package);
        return err;
    }

    *out = package;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_source_package_open(const char *path,
                                      lxlsx_source_package **out)
{
    unsigned char *buf = NULL;
    size_t len = 0;
    lxlsx_error err;

    if (!out)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    *out = NULL;

    err = slurp_file(path, &buf, &len);
    if (err != LXLSX_NO_ERROR)
        return err;

    return source_package_open_memory_owned(buf, len, out);
}

lxlsx_error lxlsx_source_package_open_memory(const void *data, size_t len,
                                             lxlsx_source_package **out)
{
    unsigned char *copy;

    if (!data || !out)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    *out = NULL;

    copy = (unsigned char *)malloc(len);
    if (!copy && len != 0)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    if (len != 0)
        memcpy(copy, data, len);

    return source_package_open_memory_owned(copy, len, out);
}

void lxlsx_source_package_close(lxlsx_source_package *package)
{
    if (!package)
        return;
    free_entries(package);
    free(package->data);
    free(package);
}

const unsigned char *lxlsx_source_package_data(const lxlsx_source_package *package,
                                               size_t *len)
{
    if (len)
        *len = package ? package->size : 0;
    return package ? package->data : NULL;
}

size_t lxlsx_source_package_entry_count(const lxlsx_source_package *package)
{
    return package ? package->entry_count : 0;
}

const lxlsx_source_package_entry_info *
lxlsx_source_package_entry_info_at(const lxlsx_source_package *package,
                                   size_t index)
{
    if (!package || index >= package->entry_count)
        return NULL;
    return &package->entries[index].info;
}

int lxlsx_source_package_find_first(const lxlsx_source_package *package,
                                    const char *name)
{
    size_t i;
    if (!package || !name)
        return -1;
    for (i = 0; i < package->entry_count; i++) {
        if (strcmp(package->entries[i].info.name, name) == 0)
            return (int)i;
    }
    return -1;
}

lxlsx_error lxlsx_source_package_read_entry(const lxlsx_source_package *package,
                                            size_t index,
                                            unsigned char **out,
                                            size_t *out_len)
{
    const lxlsx_source_entry *e;
    unsigned char *buf;

    if (!package || !out || !out_len)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    *out = NULL;
    *out_len = 0;
    if (index >= package->entry_count)
        return LXLSX_ERROR_PARAMETER_VALIDATION;

    e = &package->entries[index];
    if (e->info.uncompressed_size > (uint64_t)((size_t)-1) - 1)
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    if (e->info.compressed_size > UINT_MAX ||
        e->info.uncompressed_size > UINT_MAX) {
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    }

    buf = (unsigned char *)malloc((size_t)e->info.uncompressed_size + 1);
    if (!buf)
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;

    if (e->info.compression_method == ZIP_METHOD_STORE) {
        if (e->info.compressed_size != e->info.uncompressed_size) {
            free(buf);
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        }
        memcpy(buf, package->data + e->data_offset, (size_t)e->info.uncompressed_size);
    } else if (e->info.compression_method == ZIP_METHOD_DEFLATE) {
        z_stream zs;
        int zrc;
        memset(&zs, 0, sizeof(zs));
        zs.next_in = (Bytef *)(package->data + e->data_offset);
        zs.avail_in = (uInt)e->info.compressed_size;
        zs.next_out = buf;
        zs.avail_out = (uInt)e->info.uncompressed_size;
        zrc = inflateInit2(&zs, -MAX_WBITS);
        if (zrc != Z_OK) {
            free(buf);
            return LXLSX_ERROR_ZIP_FILE_OPERATION;
        }
        zrc = inflate(&zs, Z_FINISH);
        inflateEnd(&zs);
        if (zrc != Z_STREAM_END || zs.total_out != e->info.uncompressed_size) {
            free(buf);
            return LXLSX_ERROR_ZIP_BAD_ZIP_FILE;
        }
    } else {
        free(buf);
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;
    }

    buf[e->info.uncompressed_size] = 0;
    *out = buf;
    *out_len = (size_t)e->info.uncompressed_size;
    return LXLSX_NO_ERROR;
}

void lxlsx_source_package_free_buffer(void *buffer)
{
    free(buffer);
}

lxlsx_error lxlsx_source_package_entry_raw_local_record(
    const lxlsx_source_package *package,
    size_t index,
    const unsigned char **out,
    size_t *out_len)
{
    const lxlsx_source_entry *e;
    if (!package || !out || !out_len)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    *out = NULL;
    *out_len = 0;
    if (index >= package->entry_count)
        return LXLSX_ERROR_PARAMETER_VALIDATION;
    e = &package->entries[index];
    *out = package->data + e->local_offset;
    *out_len = e->local_record_size;
    return LXLSX_NO_ERROR;
}

lxlsx_error lxlsx_source_package_save_copy(const lxlsx_source_package *package,
                                           const char *path)
{
    FILE *fp;
    lxlsx_error err;

    if (!package || !path)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    fp = fopen(path, "wb");
    if (!fp)
        return LXLSX_ERROR_CREATING_XLSX_FILE;
    err = write_all(fp, package->data, package->size);
    if (fclose(fp) != 0 && err == LXLSX_NO_ERROR)
        err = LXLSX_ERROR_ZIP_FILE_OPERATION;
    return err;
}

static const lxlsx_source_package_replacement *find_replacement(
    const lxlsx_source_package_replacement *replacements,
    size_t replacement_count,
    size_t entry_index)
{
    size_t i;
    const lxlsx_source_package_replacement *found = NULL;
    for (i = 0; i < replacement_count; i++) {
        if (replacements[i].entry_index == entry_index)
            found = &replacements[i];
    }
    return found;
}

static lxlsx_error write_replacement_local(
    FILE *fp,
    const lxlsx_source_entry *entry,
    const lxlsx_source_package_replacement *replacement,
    uint32_t *new_offset,
    uint32_t *crc_out)
{
    uint32_t offset;
    uint32_t crc;
    lxlsx_error err;

    if (!replacement->data && replacement->size != 0)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (replacement->size > 0xffffffffUL)
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    err = current_offset(fp, &offset);
    if (err != LXLSX_NO_ERROR)
        return err;
    *new_offset = offset;

    crc = crc32(0L, Z_NULL, 0);
    crc = crc32(crc, replacement->data, (uInt)replacement->size);
    *crc_out = crc;

    if ((err = write_le32(fp, ZIP_LOCAL_FILE_SIG)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 20)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, ZIP_METHOD_STORE)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->mod_time)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->mod_date)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, crc)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, (uint32_t)replacement->size)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, (uint32_t)replacement->size)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, (uint16_t)entry->info.name_len)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) return err;
    if ((err = write_all(fp, entry->info.name, entry->info.name_len)) != LXLSX_NO_ERROR) return err;
    return write_all(fp, replacement->data, replacement->size);
}

static lxlsx_error write_replacement_central(
    FILE *fp,
    const lxlsx_source_package *package,
    const lxlsx_source_entry *entry,
    const lxlsx_source_package_replacement *replacement,
    uint32_t new_offset,
    uint32_t crc)
{
    lxlsx_error err;

    if (replacement->size > 0xffffffffUL)
        return LXLSX_ERROR_FEATURE_NOT_SUPPORTED;

    if ((err = write_le32(fp, ZIP_CENTRAL_FILE_SIG)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->version_made_by)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 20)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, ZIP_METHOD_STORE)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->mod_time)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->mod_date)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, crc)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, (uint32_t)replacement->size)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, (uint32_t)replacement->size)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, (uint16_t)entry->info.name_len)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, (uint16_t)entry->central_comment_len)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le16(fp, entry->internal_attr)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, entry->info.external_attr)) != LXLSX_NO_ERROR) return err;
    if ((err = write_le32(fp, new_offset)) != LXLSX_NO_ERROR) return err;
    if ((err = write_all(fp, entry->info.name, entry->info.name_len)) != LXLSX_NO_ERROR) return err;
    return write_all(fp, package->data + entry->central_comment_offset,
                     entry->central_comment_len);
}

static lxlsx_error write_preserved_central(
    FILE *fp,
    const lxlsx_source_package *package,
    const lxlsx_source_entry *entry,
    uint32_t new_offset)
{
    lxlsx_error err;
    unsigned char off[4];

    err = write_all(fp, package->data + entry->central_offset, 42);
    if (err != LXLSX_NO_ERROR)
        return err;
    put_le32(off, new_offset);
    err = write_all(fp, off, sizeof(off));
    if (err != LXLSX_NO_ERROR)
        return err;
    return write_all(fp, package->data + entry->central_offset + 46,
                     entry->central_size - 46);
}

lxlsx_error lxlsx_source_package_save_with_replacements(
    const lxlsx_source_package *package,
    const char *path,
    const lxlsx_source_package_replacement *replacements,
    size_t replacement_count)
{
    FILE *fp;
    uint32_t *new_offsets;
    uint32_t *new_crcs;
    uint32_t central_offset;
    uint32_t central_size;
    uint32_t end_offset;
    size_t i;
    lxlsx_error err = LXLSX_NO_ERROR;

    if (!package || !path)
        return LXLSX_ERROR_NULL_PARAMETER_IGNORED;
    if (!replacements || replacement_count == 0)
        return lxlsx_source_package_save_copy(package, path);

    new_offsets = (uint32_t *)calloc(package->entry_count, sizeof(*new_offsets));
    new_crcs = (uint32_t *)calloc(package->entry_count, sizeof(*new_crcs));
    if (!new_offsets || !new_crcs) {
        free(new_offsets);
        free(new_crcs);
        return LXLSX_ERROR_MEMORY_MALLOC_FAILED;
    }

    fp = fopen(path, "wb");
    if (!fp) {
        free(new_offsets);
        free(new_crcs);
        return LXLSX_ERROR_CREATING_XLSX_FILE;
    }

    for (i = 0; i < package->entry_count && err == LXLSX_NO_ERROR; i++) {
        const lxlsx_source_entry *entry = &package->entries[i];
        const lxlsx_source_package_replacement *replacement =
            find_replacement(replacements, replacement_count, i);

        if (replacement) {
            err = write_replacement_local(fp, entry, replacement,
                                          &new_offsets[i], &new_crcs[i]);
        } else {
            err = current_offset(fp, &new_offsets[i]);
            if (err == LXLSX_NO_ERROR) {
                err = write_all(fp, package->data + entry->local_offset,
                                entry->local_record_size);
            }
            new_crcs[i] = entry->info.crc32;
        }
    }

    if (err == LXLSX_NO_ERROR)
        err = current_offset(fp, &central_offset);

    for (i = 0; i < package->entry_count && err == LXLSX_NO_ERROR; i++) {
        const lxlsx_source_entry *entry = &package->entries[i];
        const lxlsx_source_package_replacement *replacement =
            find_replacement(replacements, replacement_count, i);
        if (replacement) {
            err = write_replacement_central(fp, package, entry, replacement,
                                            new_offsets[i], new_crcs[i]);
        } else {
            err = write_preserved_central(fp, package, entry, new_offsets[i]);
        }
    }

    if (err == LXLSX_NO_ERROR)
        err = current_offset(fp, &end_offset);
    if (err == LXLSX_NO_ERROR) {
        central_size = end_offset - central_offset;
        if ((err = write_le32(fp, ZIP_EOCD_SIG)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le16(fp, 0)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le16(fp, (uint16_t)package->entry_count)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le16(fp, (uint16_t)package->entry_count)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le32(fp, central_size)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le32(fp, central_offset)) != LXLSX_NO_ERROR) goto done;
        if ((err = write_le16(fp, (uint16_t)package->global_comment_len)) != LXLSX_NO_ERROR) goto done;
        err = write_all(fp, package->data + package->global_comment_offset,
                        package->global_comment_len);
    }

done:
    if (fclose(fp) != 0 && err == LXLSX_NO_ERROR)
        err = LXLSX_ERROR_ZIP_FILE_OPERATION;
    free(new_offsets);
    free(new_crcs);
    if (err != LXLSX_NO_ERROR)
        remove(path);
    return err;
}
