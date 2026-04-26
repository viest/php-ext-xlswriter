#ifndef LXR_STYLES_PRIV_H
#define LXR_STYLES_PRIV_H

#include <stdint.h>

#include "xlsxreader/styles.h"
#include "lxr_zip.h"

lxr_error lxr_styles_load (lxr_zip *zip, const char *path, lxr_styles **out);
void      lxr_styles_free (lxr_styles *st);
size_t    lxr_styles_count(const lxr_styles *st);

#endif
