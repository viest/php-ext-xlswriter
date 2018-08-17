/*
 * libxlsxwriter
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * drawing - A libxlsxwriter library for creating Excel XLSX drawing files.
 *
 */
#ifndef __LXW_DRAWING_H__
#define __LXW_DRAWING_H__

#include <stdint.h>

#include "common.h"

STAILQ_HEAD(lxw_drawing_objects, lxw_drawing_object);

enum lxw_drawing_types {
    LXW_DRAWING_NONE = 0,
    LXW_DRAWING_IMAGE,
    LXW_DRAWING_CHART,
    LXW_DRAWING_SHAPE
};

enum lxw_anchor_types {
    LXW_ANCHOR_TYPE_NONE = 0,
    LXW_ANCHOR_TYPE_IMAGE,
    LXW_ANCHOR_TYPE_CHART
};

enum lxw_anchor_edit_types {
    LXW_ANCHOR_EDIT_AS_NONE = 0,
    LXW_ANCHOR_EDIT_AS_RELATIVE,
    LXW_ANCHOR_EDIT_AS_ONE_CELL,
    LXW_ANCHOR_EDIT_AS_ABSOLUTE
};

enum image_types {
    LXW_IMAGE_UNKNOWN = 0,
    LXW_IMAGE_PNG,
    LXW_IMAGE_JPEG,
    LXW_IMAGE_BMP
};

/* Coordinates used in a drawing object. */
typedef struct lxw_drawing_coords {
    uint32_t col;
    uint32_t row;
    double col_offset;
    double row_offset;
} lxw_drawing_coords;

/* Object to represent the properties of a drawing. */
typedef struct lxw_drawing_object {
    uint8_t anchor_type;
    uint8_t edit_as;
    struct lxw_drawing_coords from;
    struct lxw_drawing_coords to;
    uint32_t col_absolute;
    uint32_t row_absolute;
    uint32_t width;
    uint32_t height;
    uint8_t shape;
    char *description;
    char *url;
    char *tip;

    STAILQ_ENTRY (lxw_drawing_object) list_pointers;

} lxw_drawing_object;

/*
 * Struct to represent a collection of drawings.
 */
typedef struct lxw_drawing {

    FILE *file;

    uint8_t embedded;

    struct lxw_drawing_objects *drawing_objects;

} lxw_drawing;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_drawing *lxw_drawing_new();
void lxw_drawing_free(lxw_drawing *drawing);
void lxw_drawing_assemble_xml_file(lxw_drawing *self);
void lxw_free_drawing_object(struct lxw_drawing_object *drawing_object);
void lxw_add_drawing_object(lxw_drawing *drawing,
                            lxw_drawing_object *drawing_object);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _drawing_xml_declaration(lxw_drawing *self);
#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_DRAWING_H__ */
