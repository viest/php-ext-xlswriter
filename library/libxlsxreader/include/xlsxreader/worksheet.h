#ifndef LXR_WORKSHEET_H
#define LXR_WORKSHEET_H

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxr_worksheet lxr_worksheet;

void lxr_worksheet_close(lxr_worksheet *ws);

lxr_error lxr_worksheet_next_row (lxr_worksheet *ws);
lxr_error lxr_worksheet_next_cell(lxr_worksheet *ws, lxr_cell *out);

size_t   lxr_worksheet_current_row    (const lxr_worksheet *ws);
size_t   lxr_worksheet_max_column_seen(const lxr_worksheet *ws);
uint32_t lxr_worksheet_flags          (const lxr_worksheet *ws);

typedef int (*lxr_cell_cb)   (const lxr_cell *cell, void *userdata);
typedef int (*lxr_row_end_cb)(size_t row, size_t max_col, void *userdata);

lxr_error lxr_worksheet_process(lxr_worksheet *ws,
                                lxr_cell_cb    cell_cb,
                                lxr_row_end_cb row_cb,
                                void          *userdata);

lxr_error lxr_worksheet_skip_rows(lxr_worksheet *ws, size_t n);

#ifdef __cplusplus
}
#endif

#endif
