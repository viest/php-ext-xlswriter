/*
 * libxlsxwriter
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 * drawing - A libxlsxwriter library for creating Excel XLSX drawing files.
 *
 */
#ifndef __LXLSX_DRAWING_H__
#define __LXLSX_DRAWING_H__

#include <stdint.h>
#include <string.h>

#include "common.h"

STAILQ_HEAD(lxlsx_drawing_objects, lxlsx_drawing_object);

enum lxlsx_drawing_types {
    LXLSX_DRAWING_NONE = 0,
    LXLSX_DRAWING_IMAGE,
    LXLSX_DRAWING_CHART,
    LXLSX_DRAWING_SHAPE
};

enum image_types {
    LXLSX_IMAGE_UNKNOWN = 0,
    LXLSX_IMAGE_PNG,
    LXLSX_IMAGE_JPEG,
    LXLSX_IMAGE_BMP,
    LXLSX_IMAGE_GIF
};

/* Coordinates used in a drawing object. */
typedef struct lxlsx_drawing_coords {
    uint32_t col;
    uint32_t row;
    double col_offset;
    double row_offset;
} lxlsx_drawing_coords;

/* Object to represent the properties of a drawing. */
typedef struct lxlsx_drawing_object {
    uint8_t type;
    uint8_t anchor;
    struct lxlsx_drawing_coords from;
    struct lxlsx_drawing_coords to;
    uint64_t col_absolute;
    uint64_t row_absolute;
    uint32_t width;
    uint32_t height;
    uint8_t shape;
    uint32_t rel_index;
    uint32_t url_rel_index;
    char *description;
    char *tip;
    uint8_t decorative;

    STAILQ_ENTRY (lxlsx_drawing_object) list_pointers;

} lxlsx_drawing_object;

/*
 * Struct to represent a collection of drawings.
 */
typedef struct lxlsx_drawing {

    FILE *file;

    uint8_t embedded;
    uint8_t orientation;

    struct lxlsx_drawing_objects *lxlsx_drawing_objects;

} lxlsx_drawing;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxlsx_drawing *lxlsx_drawing_new(void);
void lxlsx_drawing_free(lxlsx_drawing *drawing);
void lxlsx_drawing_assemble_xml_file(lxlsx_drawing *self);
void lxlsx_free_drawing_object(struct lxlsx_drawing_object *lxlsx_drawing_object);
void lxlsx_add_drawing_object(lxlsx_drawing *drawing,
                            lxlsx_drawing_object *lxlsx_drawing_object);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _drawing_xml_declaration(lxlsx_drawing *self);
#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXLSX_DRAWING_H__ */
