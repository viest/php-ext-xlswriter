#ifndef LXR_DRAWING_H
#define LXR_DRAWING_H

#include "common.h"
#include "worksheet.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct {
    size_t      from_row;     /* 0-based */
    size_t      from_col;
    size_t      to_row;
    size_t      to_col;
    const char *mime_type;    /* "image/png" / "image/jpeg" / ... */
    const void *data;         /* image bytes; valid only during the callback */
    size_t      data_len;
    const char *name;         /* original entry name within the xlsx */
} lxr_image;

typedef int (*lxr_image_cb)(const lxr_image *img, void *userdata);

lxr_error lxr_worksheet_iterate_images(lxr_worksheet *ws,
                                       lxr_image_cb   cb,
                                       void          *userdata);

#ifdef __cplusplus
}
#endif

#endif
