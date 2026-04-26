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

/* ----- Phase 1 metadata accessors ---------------------------------------- */

/* Merged cells. Index is 0-based. */
size_t lxr_worksheet_merged_count(const lxr_worksheet *ws);
int    lxr_worksheet_merged_get  (const lxr_worksheet *ws, size_t idx,
                                  lxr_range *out);

/* Returns 1 when (row, col) (1-based) lies inside a merged range and is NOT
 * the merge's master cell; 0 otherwise. Useful for honoring
 * LXR_SKIP_MERGED_FOLLOW in row-shaped readers. */
int lxr_worksheet_in_merge_follow(const lxr_worksheet *ws, size_t row, size_t col);

/* Hyperlinks. The url/tooltip/display/location pointers (if non-NULL) are
 * owned by the worksheet and remain valid until lxr_worksheet_close().
 * Internal links (e.g. "Sheet2!A1") have location set and url == NULL. */
typedef struct {
    lxr_range range;
    const char *url;       /* external URL, or NULL */
    const char *location;  /* internal anchor, or NULL */
    const char *display;   /* display text, or NULL */
    const char *tooltip;   /* tooltip text, or NULL */
} lxr_hyperlink;

size_t      lxr_worksheet_hyperlink_count (const lxr_worksheet *ws);
int         lxr_worksheet_hyperlink_get   (const lxr_worksheet *ws, size_t idx,
                                           lxr_hyperlink *out);
const char *lxr_worksheet_hyperlink_url   (const lxr_worksheet *ws,
                                           size_t row, size_t col);

/* Sheet protection. Booleans use 1/0. password_hash is empty when absent.
 * Per ECMA-376 the attribute defaults differ per field; we report only what
 * was *present* in the XML, plus a boolean is_present for the element. */
typedef struct {
    int  is_present;
    char password_hash[20];
    int  sheet;
    int  content;
    int  objects;
    int  scenarios;
    int  format_cells;
    int  format_columns;
    int  format_rows;
    int  insert_columns;
    int  insert_rows;
    int  insert_hyperlinks;
    int  delete_columns;
    int  delete_rows;
    int  select_locked_cells;
    int  sort;
    int  auto_filter;
    int  pivot_tables;
    int  select_unlocked_cells;
} lxr_protection;

int lxr_worksheet_protection(const lxr_worksheet *ws, lxr_protection *out);

/* Row / column metadata. row/col indexes are 1-based. */
typedef struct {
    int    has_height;
    double height;
    int    hidden;
    int    outline_level;
    int    collapsed;
    int    custom_height;
} lxr_row_options;

typedef struct {
    int    has_width;
    double width;
    int    hidden;
    int    outline_level;
    int    collapsed;
} lxr_col_options;

int lxr_worksheet_row_options(const lxr_worksheet *ws, size_t row,
                              lxr_row_options *out);
int lxr_worksheet_col_options(const lxr_worksheet *ws, size_t col,
                              lxr_col_options *out);

int    lxr_worksheet_default_row_height(const lxr_worksheet *ws, double *out);
int    lxr_worksheet_default_col_width (const lxr_worksheet *ws, double *out);

#ifdef __cplusplus
}
#endif

#endif
