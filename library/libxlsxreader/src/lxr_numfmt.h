#ifndef LXR_NUMFMT_H
#define LXR_NUMFMT_H

#include <stdint.h>

#include "xlsxreader/styles.h"

const char       *lxr_numfmt_builtin_format  (uint16_t num_fmt_id);
lxr_fmt_category  lxr_numfmt_builtin_category(uint16_t num_fmt_id);
lxr_fmt_category  lxr_numfmt_classify        (const char *fmt_string);

#endif
