#ifndef LXLSX_XLSX_UTIL_H
#define LXLSX_XLSX_UTIL_H

#include <stddef.h>

#include "libxlsx/common.h"
#include "libxlsx/worksheet.h"
#include "zip_io.h"

typedef struct {
    char *id;
    char *type;
    char *target;
} lxlsx_reader_rel;

typedef struct {
    lxlsx_reader_rel *items;
    size_t count;
    size_t cap;
} lxlsx_reader_rel_map;

char *lxlsx_reader_strdup(const char *s);
char *lxlsx_reader_strndup(const char *s, size_t len);

int  lxlsx_reader_buf_append(char **buf, size_t *len, size_t *cap,
                             const char *src, size_t n);
void lxlsx_reader_buf_reset(char *buf, size_t *len);
void lxlsx_reader_copy_attr(char *dst, size_t cap, const char *src);

char *lxlsx_reader_zip_normalize_path(const char *path);
char *lxlsx_reader_zip_base_path(const char *path);
char *lxlsx_reader_zip_rels_path(const char *path);
char *lxlsx_reader_zip_join_path(const char *base, const char *target);
char *lxlsx_reader_zip_resolve_path(const char *source_path, const char *target);
int   lxlsx_reader_ascii_case_eq(const char *a, const char *b);

void lxlsx_reader_parse_a1_ref(const char *ref, size_t *out_row, size_t *out_col);
void lxlsx_reader_parse_a1_range(const char *ref, lxlsx_reader_range *out);

void lxlsx_reader_rel_map_init(lxlsx_reader_rel_map *map);
void lxlsx_reader_rel_map_free(lxlsx_reader_rel_map *map);
lxlsx_reader_error lxlsx_reader_rel_map_push(lxlsx_reader_rel_map *map,
                                             const char *id,
                                             const char *type,
                                             const char *target);
const char *lxlsx_reader_rel_map_target_for_id(const lxlsx_reader_rel_map *map,
                                               const char *id);
const char *lxlsx_reader_rel_map_first_target_by_type_suffix(
    const lxlsx_reader_rel_map *map, const char *type_suffix);
lxlsx_reader_error lxlsx_reader_load_rels(lxlsx_reader_zip *zip,
                                          const char *source_path,
                                          lxlsx_reader_rel_map *out,
                                          int missing_ok);

int lxlsx_reader_zip_read_entry_all(lxlsx_reader_zip *zip, const char *path,
                                    void **out_data, size_t *out_len);

#endif
