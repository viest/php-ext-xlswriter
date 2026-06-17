#ifndef LXLSX_READER_STYLES_PRIV_H
#define LXLSX_READER_STYLES_PRIV_H

#include <stdint.h>

#include "lxlsx/reader/styles.h"
#include "lxlsx_reader_zip.h"

lxlsx_reader_error lxlsx_reader_styles_load (lxlsx_reader_zip *zip, const char *path, lxlsx_reader_styles **out);
void      lxlsx_reader_styles_free (lxlsx_reader_styles *st);
size_t    lxlsx_reader_styles_count(const lxlsx_reader_styles *st);

#endif
