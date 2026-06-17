/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * packager - A libxlsxwriter library for creating Excel XLSX packager files.
 *
 */
#ifndef __LXW_PACKAGER_H__
#define __LXW_PACKAGER_H__

#include <stdint.h>

#ifdef USE_SYSTEM_MINIZIP
#ifdef __GNUC__
#pragma GCC system_header
#endif
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
#include "table.h"
#include "theme.h"
#include "styles.h"
#include "format.h"
#include "content_types.h"
#include "relationships.h"
#include "vml.h"
#include "comment.h"
#include "metadata.h"
#include "rich_value.h"
#include "rich_value_rel.h"
#include "rich_value_types.h"
#include "rich_value_structure.h"

#define LXW_ZIP_BUFFER_SIZE (16384)

/* If zip returns a ZIP_XXX error then errno is set and we can trap that in
 * workbook.c. Otherwise return a default libxlsxwriter error. */
#define RETURN_ON_ZIP_ERROR(err, default_err)       \
    do {                                            \
        if (err == ZIP_ERRNO)                       \
            return LXW_ERROR_ZIP_FILE_OPERATION;    \
        else if (err == ZIP_PARAMERROR)             \
            return LXW_ERROR_ZIP_PARAMETER_ERROR;   \
        else if (err == ZIP_BADZIPFILE)             \
            return LXW_ERROR_ZIP_BAD_ZIP_FILE;      \
        else if (err == ZIP_INTERNALERROR)          \
            return LXW_ERROR_ZIP_INTERNAL_ERROR;    \
        else                                        \
            return default_err;                     \
    } while (0)

/*
 * Struct to represent a packager.
 */
typedef struct lxw_packager {

    FILE *file;
    lxw_workbook *workbook;

    size_t buffer_size;
    size_t output_buffer_size;
    zipFile zipfile;
    zip_fileinfo zipfile_info;
    const char *filename;
    const char *buffer;
    char *output_buffer;
    const char *tmpdir;
    uint8_t use_zip64;

} lxw_packager;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_packager *lxw_packager_new(const char *filename, const char *tmpdir,
                               uint8_t use_zip64);
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
