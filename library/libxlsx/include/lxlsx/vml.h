/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * vml - A libxlsxwriter library for creating Excel XLSX vml files.
 *
 */
#ifndef __LXW_VML_H__
#define __LXW_VML_H__

#include <stdint.h>

#include "common.h"
#include "worksheet.h"

/*
 * Struct to represent a vml object.
 */
typedef struct lxw_vml {

    FILE *file;
    uint8_t type;
    struct lxw_comment_objs *button_objs;
    struct lxw_comment_objs *comment_objs;
    struct lxw_comment_objs *image_objs;
    char *vml_data_id_str;
    uint32_t vml_shape_id;
    uint8_t comment_display_default;

} lxw_vml;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_vml *lxw_vml_new(void);
void lxw_vml_free(lxw_vml *vml);
void lxw_vml_assemble_xml_file(lxw_vml *self);

/* Declarations required for unit testing. */
#ifdef TESTING

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_VML_H__ */
