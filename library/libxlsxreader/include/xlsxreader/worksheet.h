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

/* ---- Data validation ---------------------------------------------------- */

/* All string pointers are owned by the worksheet; valid until close. */
typedef struct {
    const char *type;          /* "whole","decimal","list","date","time",
                                  "textLength","custom" — or NULL */
    const char *operator_;     /* "between","notBetween","equal", ... — or NULL */
    const char *error_style;   /* "stop","warning","information" — or NULL */
    const char *formula1;
    const char *formula2;
    const char *prompt;
    const char *prompt_title;
    const char *error;
    const char *error_title;
    const char *sqref;         /* original sqref (may be space-separated list) */
    int allow_blank;
    int show_drop_down;        /* per OOXML this is INVERTED — 1 means "hide" */
    int show_input_message;
    int show_error_message;
} lxr_data_validation;

size_t lxr_worksheet_data_validation_count(const lxr_worksheet *ws);
int    lxr_worksheet_data_validation_get  (const lxr_worksheet *ws, size_t idx,
                                           lxr_data_validation *out);

/* ---- AutoFilter --------------------------------------------------------- */

typedef enum {
    LXR_FILTER_NONE = 0,
    LXR_FILTER_LIST,           /* discrete value match (<filters>) */
    LXR_FILTER_CUSTOM,         /* operator-based (<customFilters>) */
    LXR_FILTER_TOP10,          /* <top10>: top/bottom N (or %) */
    LXR_FILTER_DYNAMIC         /* <dynamicFilter>: aboveAverage etc. */
} lxr_filter_kind;

typedef struct {
    int             col_id;          /* 0-based column offset within the autoFilter range */
    lxr_filter_kind kind;
    /* For LXR_FILTER_LIST: list of value strings. */
    const char    **values;          /* NULL-terminated; NULL when not applicable */
    /* For LXR_FILTER_CUSTOM: at most two predicates joined by AND/OR. */
    int             custom_and;      /* 1 = AND, 0 = OR */
    const char     *custom_op_1;     /* "equal","greaterThan", ... */
    const char     *custom_val_1;
    const char     *custom_op_2;
    const char     *custom_val_2;
    /* For LXR_FILTER_TOP10. */
    int             top;             /* 1=top, 0=bottom */
    int             percent;         /* 1=percent */
    double          top_value;       /* N */
} lxr_filter_column;

typedef struct {
    const char              *range;       /* original ref attr (e.g. "A1:Z100"), or NULL */
    const lxr_filter_column *columns;
    size_t                   columns_count;
} lxr_autofilter;

int lxr_worksheet_autofilter(const lxr_worksheet *ws, lxr_autofilter *out);

/* ---- Page setup --------------------------------------------------------- */

typedef struct {
    /* Margins (inches). has_* indicate explicit element presence. */
    int    has_margins;
    double margin_left, margin_right, margin_top, margin_bottom;
    double margin_header, margin_footer;

    /* pageSetup attrs. */
    int    has_setup;
    int    paper_size;          /* 0 if absent */
    int    fit_to_width;
    int    fit_to_height;
    int    scale;               /* 0 if absent (interpreted as default 100) */
    int    orientation_landscape;  /* 1 = landscape, 0 = portrait/default */
    int    horizontal_dpi;
    int    vertical_dpi;
    int    first_page_number;
    int    use_first_page_number;

    /* printOptions. */
    int    print_horizontal_centered;
    int    print_vertical_centered;
    int    print_grid_lines;
    int    print_headings;

    /* Header / footer text (strings; pointers owned by worksheet). */
    const char *odd_header;
    const char *odd_footer;
    const char *even_header;
    const char *even_footer;
    const char *first_header;
    const char *first_footer;
    int   different_odd_even;
    int   different_first;
    int   scale_with_doc;
    int   align_with_margins;
} lxr_page_setup;

int lxr_worksheet_page_setup(const lxr_worksheet *ws, lxr_page_setup *out);

/* ---- Conditional formats (§8.2.3) -------------------------------------- */

typedef struct {
    const char *type;       /* "cellIs","expression","colorScale","dataBar",
                                "iconSet","top10","aboveAverage","duplicateValues",
                                "uniqueValues","containsText","containsBlanks",
                                "notContainsBlanks","containsErrors","notContainsErrors",
                                "timePeriod" — or NULL */
    const char *operator_;  /* "greaterThan","between", ... — or NULL */
    int         priority;
    int         stop_if_true;
    int         dxf_id;     /* -1 if absent */
    int         percent;    /* top10's percent attr */
    int         bottom;     /* top10's bottom attr */
    double      rank;       /* top10's rank attr */
    const char *text;       /* containsText/notContainsText/etc. */
    const char *time_period; /* timePeriod's attr */
    /* up to 2 inline formulas (formula1, formula2) */
    const char *formula1;
    const char *formula2;
} lxr_cf_rule;

typedef struct {
    const char        *sqref;     /* the conditionalFormatting sqref */
    const lxr_cf_rule *rules;
    size_t             rules_count;
} lxr_cf_block;

size_t lxr_worksheet_cf_block_count(const lxr_worksheet *ws);
int    lxr_worksheet_cf_block_get  (const lxr_worksheet *ws, size_t idx,
                                    lxr_cf_block *out);

/* ---- Comments (§8.2.1) -------------------------------------------------- */

typedef struct {
    size_t      row;          /* 1-based */
    size_t      col;          /* 1-based */
    const char *text;
    const char *author;
    int         visible;      /* legacy VML's "visible" attr */
    int         threaded;
    const char *parent_id;    /* threaded reply parent GUID, or NULL */
} lxr_comment_info;

typedef int (*lxr_comment_cb)(const lxr_comment_info *info, void *userdata);

/* Iterates legacy + threaded comments. cb returning non-zero stops. Returns
 * LXR_NO_ERROR on success; ZIP_ENTRY_NOT_FOUND when the sheet has none. */
lxr_error lxr_worksheet_iterate_comments(lxr_worksheet *ws,
                                         lxr_comment_cb cb, void *userdata);

/* ---- Chart metadata (§8.2.4) ------------------------------------------- */

typedef struct {
    const char *name;        /* series name, may be NULL */
    const char *categories;  /* category range / formula */
    const char *values;      /* values range */
} lxr_chart_series_info;

typedef struct {
    const char *type;        /* "bar"/"line"/"pie"/"area"/"scatter"/"radar"/"doughnut" or NULL */
    const char *title;
    /* 1-based inclusive anchor; 0 if not from a twoCellAnchor. */
    size_t from_row, from_col, to_row, to_col;
    const lxr_chart_series_info *series;
    size_t                       series_count;
} lxr_chart_meta;

typedef int (*lxr_chart_cb)(const lxr_chart_meta *info, void *userdata);

lxr_error lxr_worksheet_iterate_charts(lxr_worksheet *ws,
                                       lxr_chart_cb cb, void *userdata);

/* ---- Rich-text runs (§8.2.2) ------------------------------------------- */

/* For LXR_CELL_STRING (SST) and LXR_CELL_INLINE_STRING cells, fills `out`
 * with up to `cap` runs and returns the actual count. When `out` is NULL
 * just returns the run count. Callers can sniff the run count first, then
 * allocate appropriately. The pointers inside each run remain valid until
 * the next streaming step that overwrites the inline buffer (for inline
 * strings) or until the workbook is closed (for SST strings). */
size_t lxr_cell_string_runs(const lxr_worksheet *ws, const lxr_cell *c,
                            lxr_string_run *out, size_t cap);

#ifdef __cplusplus
}
#endif

#endif
