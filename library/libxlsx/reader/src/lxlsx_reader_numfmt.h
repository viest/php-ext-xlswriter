#ifndef LXLSX_READER_NUMFMT_H
#define LXLSX_READER_NUMFMT_H

#include <stdint.h>

#include "lxlsx/reader/styles.h"

const char       *lxlsx_reader_numfmt_builtin_format  (uint16_t num_fmt_id);
lxlsx_reader_fmt_category  lxlsx_reader_numfmt_builtin_category(uint16_t num_fmt_id);
lxlsx_reader_fmt_category  lxlsx_reader_numfmt_classify        (const char *fmt_string);

#endif
