#ifndef LXLSX_SST_H
#define LXLSX_SST_H

#include <stddef.h>
#include <stdint.h>

#include "libxlsx/common.h"
#include "zip_io.h"

typedef struct lxlsx_reader_sst lxlsx_reader_sst;

lxlsx_reader_error    lxlsx_reader_sst_open (lxlsx_reader_zip *zip, const char *path,
                           lxlsx_reader_sst_mode mode, lxlsx_reader_sst **out);
const char * lxlsx_reader_sst_get  (lxlsx_reader_sst *sst, uint32_t index);
size_t       lxlsx_reader_sst_count(const lxlsx_reader_sst *sst);
size_t       lxlsx_reader_sst_loaded_count(const lxlsx_reader_sst *sst);
void         lxlsx_reader_sst_close(lxlsx_reader_sst *sst);

/* Owned by the SST — pointers here remain valid until lxlsx_reader_sst_close. */
typedef struct lxlsx_reader_sst_run {
    char  *text;     /* owned */
    char  *font_name;
    double font_size;
    int    bold;
    int    italic;
    int    strike;
    int    underline;
    char  *color;    /* "FF0000" hex */
} lxlsx_reader_sst_run;

/* For STREAMING SST mode the runs of a given index are populated as a
 * side-effect of the same lazy resume that loads the string. Returns the
 * run array head and count for index, or NULL/0 on absent index. */
const lxlsx_reader_sst_run *lxlsx_reader_sst_get_runs(lxlsx_reader_sst *sst, uint32_t index, size_t *out_count);

#endif
