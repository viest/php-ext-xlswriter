#include <ctype.h>
#include <stdlib.h>
#include <string.h>

#include "xlsx_util.h"
#include "xml_pump.h"

char *lxlsx_reader_strndup(const char *s, size_t len)
{
    char *out;
    if (!s) return NULL;
    out = (char *)malloc(len + 1);
    if (!out) return NULL;
    memcpy(out, s, len);
    out[len] = 0;
    return out;
}

char *lxlsx_reader_strdup(const char *s)
{
    return s ? lxlsx_reader_strndup(s, strlen(s)) : NULL;
}

int lxlsx_reader_buf_append(char **buf, size_t *len, size_t *cap,
                            const char *src, size_t n)
{
    size_t need;
    size_t next_cap;
    char *next;

    if (!buf || !len || !cap || (!src && n > 0)) return -1;
    if (n > ((size_t)-1) - *len - 1) return -1;
    need = *len + n + 1;

    if (need > *cap) {
        next_cap = *cap ? *cap : 64;
        while (next_cap < need) {
            if (next_cap > ((size_t)-1) / 2) return -1;
            next_cap *= 2;
        }
        next = (char *)realloc(*buf, next_cap);
        if (!next) return -1;
        *buf = next;
        *cap = next_cap;
    }

    if (n > 0) memcpy(*buf + *len, src, n);
    *len += n;
    (*buf)[*len] = 0;
    return 0;
}

void lxlsx_reader_buf_reset(char *buf, size_t *len)
{
    if (len) *len = 0;
    if (buf) buf[0] = 0;
}

void lxlsx_reader_copy_attr(char *dst, size_t cap, const char *src)
{
    size_t n;
    if (!dst || cap == 0) return;
    if (!src) {
        dst[0] = 0;
        return;
    }
    n = strlen(src);
    if (n >= cap) n = cap - 1;
    memcpy(dst, src, n);
    dst[n] = 0;
}

char *lxlsx_reader_zip_normalize_path(const char *path)
{
    char *copy;
    char **segments;
    char *p;
    size_t count = 0;
    size_t i;
    size_t len;
    size_t out_len = 0;
    char *out;
    char *write;

    if (!path) return NULL;
    while (*path == '/') path++;

    len = strlen(path);
    copy = lxlsx_reader_strndup(path, len);
    if (!copy) return NULL;

    segments = (char **)calloc(len + 1, sizeof(*segments));
    if (!segments) {
        free(copy);
        return NULL;
    }

    p = copy;
    while (*p) {
        char *seg;
        while (*p == '/') p++;
        if (!*p) break;
        seg = p;
        while (*p && *p != '/') p++;
        if (*p == '/') *p++ = 0;

        if (strcmp(seg, ".") == 0 || seg[0] == 0) {
            continue;
        } else if (strcmp(seg, "..") == 0) {
            if (count > 0) count--;
        } else {
            segments[count++] = seg;
        }
    }

    for (i = 0; i < count; i++)
        out_len += strlen(segments[i]) + (i ? 1 : 0);

    out = (char *)malloc(out_len + 1);
    if (!out) {
        free(segments);
        free(copy);
        return NULL;
    }

    write = out;
    for (i = 0; i < count; i++) {
        size_t seg_len = strlen(segments[i]);
        if (i) *write++ = '/';
        memcpy(write, segments[i], seg_len);
        write += seg_len;
    }
    *write = 0;

    free(segments);
    free(copy);
    return out;
}

char *lxlsx_reader_zip_base_path(const char *path)
{
    const char *slash = path ? strrchr(path, '/') : NULL;
    size_t n;
    if (!slash) return lxlsx_reader_strdup("");
    n = (size_t)(slash - path) + 1;
    return lxlsx_reader_strndup(path, n);
}

char *lxlsx_reader_zip_rels_path(const char *path)
{
    const char *slash;
    const char *file;
    size_t prefix_len;
    size_t file_len;
    char *out;

    if (!path) return NULL;
    slash = strrchr(path, '/');
    prefix_len = slash ? (size_t)(slash - path) + 1 : 0;
    file = slash ? slash + 1 : path;
    file_len = strlen(file);

    out = (char *)malloc(prefix_len + 6 + file_len + 5 + 1);
    if (!out) return NULL;
    memcpy(out, path, prefix_len);
    memcpy(out + prefix_len, "_rels/", 6);
    memcpy(out + prefix_len + 6, file, file_len);
    memcpy(out + prefix_len + 6 + file_len, ".rels", 5);
    out[prefix_len + 6 + file_len + 5] = 0;
    return out;
}

char *lxlsx_reader_zip_join_path(const char *base, const char *target)
{
    size_t base_len;
    size_t target_len;
    int need_slash;
    char *joined;

    if (!target) return NULL;
    if (target[0] == '/') return lxlsx_reader_zip_normalize_path(target + 1);
    if (!base || !*base) return lxlsx_reader_zip_normalize_path(target);

    base_len = strlen(base);
    target_len = strlen(target);
    need_slash = base_len > 0 && base[base_len - 1] != '/';

    joined = (char *)malloc(base_len + (need_slash ? 1 : 0) + target_len + 1);
    if (!joined) return NULL;
    memcpy(joined, base, base_len);
    if (need_slash) joined[base_len++] = '/';
    memcpy(joined + base_len, target, target_len);
    joined[base_len + target_len] = 0;

    {
        char *normalized = lxlsx_reader_zip_normalize_path(joined);
        free(joined);
        return normalized;
    }
}

char *lxlsx_reader_zip_resolve_path(const char *source_path, const char *target)
{
    char *base;
    char *out;

    if (!source_path || !target) return NULL;
    base = lxlsx_reader_zip_base_path(source_path);
    if (!base) return NULL;
    out = lxlsx_reader_zip_join_path(base, target);
    free(base);
    return out;
}

int lxlsx_reader_ascii_case_eq(const char *a, const char *b)
{
    if (!a || !b) return 0;
    while (*a && *b) {
        int x = tolower((unsigned char)*a++);
        int y = tolower((unsigned char)*b++);
        if (x != y) return 0;
    }
    return *a == 0 && *b == 0;
}

void lxlsx_reader_parse_a1_ref(const char *ref, size_t *out_row, size_t *out_col)
{
    size_t row = 0;
    size_t col = 0;

    if (out_row) *out_row = 0;
    if (out_col) *out_col = 0;
    if (!ref) return;

    while (*ref == '$') ref++;
    while ((*ref >= 'A' && *ref <= 'Z') || (*ref >= 'a' && *ref <= 'z')) {
        col = col * 26 + (size_t)(toupper((unsigned char)*ref) - 'A' + 1);
        ref++;
    }
    while (*ref == '$') ref++;
    while (*ref >= '0' && *ref <= '9') {
        row = row * 10 + (size_t)(*ref - '0');
        ref++;
    }

    if (out_row) *out_row = row;
    if (out_col) *out_col = col;
}

void lxlsx_reader_parse_a1_range(const char *ref, lxlsx_reader_range *out)
{
    const char *colon;
    size_t r1 = 0, c1 = 0, r2 = 0, c2 = 0;

    if (out) out->first_row = out->first_col = out->last_row = out->last_col = 0;
    if (!ref || !out) return;

    colon = strchr(ref, ':');
    if (colon) {
        char head[64];
        size_t n = (size_t)(colon - ref);
        if (n >= sizeof(head)) n = sizeof(head) - 1;
        memcpy(head, ref, n);
        head[n] = 0;
        lxlsx_reader_parse_a1_ref(head, &r1, &c1);
        lxlsx_reader_parse_a1_ref(colon + 1, &r2, &c2);
    } else {
        lxlsx_reader_parse_a1_ref(ref, &r1, &c1);
        r2 = r1;
        c2 = c1;
    }

    out->first_row = r1;
    out->first_col = c1;
    out->last_row = r2;
    out->last_col = c2;
}

void lxlsx_reader_rel_map_init(lxlsx_reader_rel_map *map)
{
    if (!map) return;
    map->items = NULL;
    map->count = 0;
    map->cap = 0;
}

void lxlsx_reader_rel_map_free(lxlsx_reader_rel_map *map)
{
    size_t i;
    if (!map) return;
    for (i = 0; i < map->count; i++) {
        free(map->items[i].id);
        free(map->items[i].type);
        free(map->items[i].target);
    }
    free(map->items);
    lxlsx_reader_rel_map_init(map);
}

lxlsx_reader_error lxlsx_reader_rel_map_push(lxlsx_reader_rel_map *map,
                                             const char *id,
                                             const char *type,
                                             const char *target)
{
    lxlsx_reader_rel *rel;
    if (!map) return LXLSX_READER_ERROR_NULL_PARAMETER;
    if (map->count >= map->cap) {
        size_t next_cap = map->cap ? map->cap * 2 : 4;
        lxlsx_reader_rel *next =
            (lxlsx_reader_rel *)realloc(map->items, next_cap * sizeof(*next));
        if (!next) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
        map->items = next;
        map->cap = next_cap;
    }

    rel = &map->items[map->count];
    memset(rel, 0, sizeof(*rel));
    rel->id = lxlsx_reader_strdup(id);
    rel->type = lxlsx_reader_strdup(type);
    rel->target = lxlsx_reader_strdup(target);
    if ((id && !rel->id) || (type && !rel->type) || (target && !rel->target)) {
        free(rel->id);
        free(rel->type);
        free(rel->target);
        memset(rel, 0, sizeof(*rel));
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }

    map->count++;
    return LXLSX_READER_NO_ERROR;
}

const char *lxlsx_reader_rel_map_target_for_id(const lxlsx_reader_rel_map *map,
                                               const char *id)
{
    size_t i;
    if (!map || !id) return NULL;
    for (i = 0; i < map->count; i++) {
        if (map->items[i].id && strcmp(map->items[i].id, id) == 0)
            return map->items[i].target;
    }
    return NULL;
}

const char *lxlsx_reader_rel_map_first_target_by_type_suffix(
    const lxlsx_reader_rel_map *map, const char *type_suffix)
{
    size_t i;
    size_t suffix_len;
    if (!map) return NULL;
    if (!type_suffix) {
        return map->count > 0 ? map->items[0].target : NULL;
    }

    suffix_len = strlen(type_suffix);
    for (i = 0; i < map->count; i++) {
        const char *type = map->items[i].type;
        size_t type_len;
        if (!type || !map->items[i].target) continue;
        type_len = strlen(type);
        if (type_len >= suffix_len &&
            strcmp(type + type_len - suffix_len, type_suffix) == 0) {
            return map->items[i].target;
        }
    }
    return NULL;
}

typedef struct {
    lxlsx_reader_rel_map *map;
    lxlsx_reader_error err;
} rel_parse_ctx;

static void rel_on_start(void *ud, const char *name, const char **attrs)
{
    rel_parse_ctx *ctx = (rel_parse_ctx *)ud;
    const char *id;
    const char *type;
    const char *target;

    if (ctx->err != LXLSX_READER_NO_ERROR) return;
    if (!lxlsx_reader_xml_name_eq(name, "Relationship")) return;

    id = lxlsx_reader_xml_attr(attrs, "Id");
    type = lxlsx_reader_xml_attr(attrs, "Type");
    target = lxlsx_reader_xml_attr(attrs, "Target");
    if (!id || !target) return;

    ctx->err = lxlsx_reader_rel_map_push(ctx->map, id, type, target);
}

lxlsx_reader_error lxlsx_reader_load_rels(lxlsx_reader_zip *zip,
                                          const char *source_path,
                                          lxlsx_reader_rel_map *out,
                                          int missing_ok)
{
    char *rels_path;
    lxlsx_reader_zip_file *zf;
    lxlsx_reader_xml_pump *pump;
    rel_parse_ctx ctx;
    lxlsx_reader_error rc;

    if (!out) return LXLSX_READER_ERROR_NULL_PARAMETER;
    lxlsx_reader_rel_map_init(out);
    if (!zip || !source_path) {
        return missing_ok ? LXLSX_READER_NO_ERROR : LXLSX_READER_ERROR_NULL_PARAMETER;
    }

    rels_path = lxlsx_reader_zip_rels_path(source_path);
    if (!rels_path) return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;

    zf = lxlsx_reader_zip_open_entry(zip, rels_path);
    free(rels_path);
    if (!zf) return missing_ok ? LXLSX_READER_NO_ERROR : LXLSX_READER_ERROR_ZIP_ENTRY_NOT_FOUND;

    pump = lxlsx_reader_xml_pump_create_zip_file(zf);
    if (!pump) {
        lxlsx_reader_zip_close_entry(zf);
        return LXLSX_READER_ERROR_MEMORY_MALLOC_FAILED;
    }

    ctx.map = out;
    ctx.err = LXLSX_READER_NO_ERROR;
    lxlsx_reader_xml_pump_set_handlers(pump, rel_on_start, NULL, NULL, &ctx);
    rc = lxlsx_reader_xml_pump_run(pump);
    if (rc == LXLSX_READER_NO_ERROR)
        rc = ctx.err;

    lxlsx_reader_xml_pump_destroy(pump);
    lxlsx_reader_zip_close_entry(zf);
    if (rc != LXLSX_READER_NO_ERROR)
        lxlsx_reader_rel_map_free(out);
    return rc;
}

int lxlsx_reader_zip_read_entry_all(lxlsx_reader_zip *zip, const char *path,
                                    void **out_data, size_t *out_len)
{
    lxlsx_reader_zip_file *zf;
    char *buf;
    size_t cap = 8192;
    size_t len = 0;

    if (!out_data || !out_len) return -1;
    *out_data = NULL;
    *out_len = 0;
    if (!zip || !path) return -1;

    zf = lxlsx_reader_zip_open_entry(zip, path);
    if (!zf) return -1;

    buf = (char *)malloc(cap + 1);
    if (!buf) {
        lxlsx_reader_zip_close_entry(zf);
        return -1;
    }

    for (;;) {
        ssize_t n;
        if (cap - len < 4096) {
            char *next;
            if (cap > ((size_t)-1) / 2) {
                free(buf);
                lxlsx_reader_zip_close_entry(zf);
                return -1;
            }
            cap *= 2;
            next = (char *)realloc(buf, cap + 1);
            if (!next) {
                free(buf);
                lxlsx_reader_zip_close_entry(zf);
                return -1;
            }
            buf = next;
        }
        n = lxlsx_reader_zip_read(zf, buf + len, cap - len);
        if (n < 0) {
            free(buf);
            lxlsx_reader_zip_close_entry(zf);
            return -1;
        }
        if (n == 0) break;
        len += (size_t)n;
    }

    lxlsx_reader_zip_close_entry(zf);
    buf[len] = 0;
    *out_data = buf;
    *out_len = len;
    return 0;
}
