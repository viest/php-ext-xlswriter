#ifndef LXLSX_STYLES_PRIVATE_H
#define LXLSX_STYLES_PRIVATE_H

#include <stdint.h>

#include "libxlsx/styles.h"
#include "zip_io.h"

lxlsx_reader_error lxlsx_reader_styles_load (lxlsx_reader_zip *zip, const char *path, lxlsx_reader_styles **out);
void      lxlsx_reader_styles_free (lxlsx_reader_styles *st);
size_t    lxlsx_reader_styles_count(const lxlsx_reader_styles *st);

#endif
