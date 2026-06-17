#ifndef LXLSX_READER_ZIP_H
#define LXLSX_READER_ZIP_H

#include <stddef.h>

#include "lxlsx_reader_compat.h"
#include "lxlsx/reader/common.h"

typedef struct lxlsx_reader_zip      lxlsx_reader_zip;
typedef struct lxlsx_reader_zip_file lxlsx_reader_zip_file;

lxlsx_reader_zip *lxlsx_reader_zip_open_path  (const char *path);
lxlsx_reader_zip *lxlsx_reader_zip_open_fd    (int fd);
lxlsx_reader_zip *lxlsx_reader_zip_open_memory(const void *data, size_t len);
void     lxlsx_reader_zip_close      (lxlsx_reader_zip *zip);

int           lxlsx_reader_zip_entry_exists(lxlsx_reader_zip *zip, const char *name);
lxlsx_reader_zip_file *lxlsx_reader_zip_open_entry  (lxlsx_reader_zip *zip, const char *name);
ssize_t       lxlsx_reader_zip_read        (lxlsx_reader_zip_file *zf, void *buf, size_t n);
void          lxlsx_reader_zip_close_entry (lxlsx_reader_zip_file *zf);

typedef int (*lxlsx_reader_zip_iter_fn)(const char *name, void *userdata);

lxlsx_reader_error lxlsx_reader_zip_iterate_entries(lxlsx_reader_zip *zip,
                                  lxlsx_reader_zip_iter_fn fn,
                                  void *userdata);

#endif
