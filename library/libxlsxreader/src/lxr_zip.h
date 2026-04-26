#ifndef LXR_ZIP_H
#define LXR_ZIP_H

#include <stddef.h>
#include <sys/types.h>

#include "xlsxreader/common.h"

typedef struct lxr_zip      lxr_zip;
typedef struct lxr_zip_file lxr_zip_file;

lxr_zip *lxr_zip_open_path  (const char *path);
lxr_zip *lxr_zip_open_fd    (int fd);
lxr_zip *lxr_zip_open_memory(const void *data, size_t len);
void     lxr_zip_close      (lxr_zip *zip);

int           lxr_zip_entry_exists(lxr_zip *zip, const char *name);
lxr_zip_file *lxr_zip_open_entry  (lxr_zip *zip, const char *name);
ssize_t       lxr_zip_read        (lxr_zip_file *zf, void *buf, size_t n);
void          lxr_zip_close_entry (lxr_zip_file *zf);

typedef int (*lxr_zip_iter_fn)(const char *name, void *userdata);

lxr_error lxr_zip_iterate_entries(lxr_zip *zip,
                                  lxr_zip_iter_fn fn,
                                  void *userdata);

#endif
