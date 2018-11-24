/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * packager - A libxlsxwriter library for creating Excel XLSX packager files.
 *
 */
#ifndef __LXW_PACKAGER_H__
#define __LXW_PACKAGER_H__

#include <stdint.h>

#ifdef USE_SYSTEM_MINIZIP
#include "minizip/zip.h"
#else
#include "third_party/zip.h"
#endif

#include "common.h"
#include "workbook.h"
#include "worksheet.h"
#include "shared_strings.h"
#include "app.h"
#include "core.h"
#include "custom.h"
#include "theme.h"
#include "styles.h"
#include "format.h"
#include "content_types.h"
#include "relationships.h"

#define LXW_ZIP_BUFFER_SIZE (16384)

/* If zlib returns Z_ERRNO then errno is set and we can trap that. Otherwise
 * return a default libxlsxwriter error. */
#define RETURN_ON_ZIP_ERROR(err, default_err)   \
    if (err == Z_ERRNO)                         \
        return LXW_ERROR_ZIP_FILE_OPERATION;    \
    else                                        \
        return default_err;

/*
 * Struct to represent a packager.
 */
typedef struct lxw_packager {

    FILE *file;
    lxw_workbook *workbook;

    size_t buffer_size;
    zipFile zipfile;
    zip_fileinfo zipfile_info;
    char *filename;
    char *buffer;
    char *tmpdir;

    uint16_t chart_count;
    uint16_t drawing_count;

} lxw_packager;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_packager *lxw_packager_new(const char *filename, char *tmpdir);
void lxw_packager_free(lxw_packager *packager);
lxw_error lxw_create_package(lxw_packager *self);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_PACKAGER_H__ */
