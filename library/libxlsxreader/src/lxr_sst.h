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

#endif
