#ifndef LXR_SST_H
#define LXR_SST_H

#include <stddef.h>
#include <stdint.h>

#include "xlsxreader/common.h"
#include "lxr_zip.h"

typedef struct lxr_sst lxr_sst;

lxr_error    lxr_sst_open (lxr_zip *zip, const char *path,
                           lxr_sst_mode mode, lxr_sst **out);
const char * lxr_sst_get  (lxr_sst *sst, uint32_t index);
size_t       lxr_sst_count(const lxr_sst *sst);
size_t       lxr_sst_loaded_count(const lxr_sst *sst);
void         lxr_sst_close(lxr_sst *sst);

/* Owned by the SST — pointers here remain valid until lxr_sst_close. */
typedef struct lxr_sst_run {
    char  *text;     /* owned */
    char  *font_name;
    double font_size;
    int    bold;
    int    italic;
    int    strike;
    int    underline;
    char  *color;    /* "FF0000" hex */
} lxr_sst_run;

/* For STREAMING SST mode the runs of a given index are populated as a
 * side-effect of the same lazy resume that loads the string. Returns the
 * run array head and count for index, or NULL/0 on absent index. */
const lxr_sst_run *lxr_sst_get_runs(lxr_sst *sst, uint32_t index, size_t *out_count);

#endif
