#ifndef LXLSX_READER_WORKSHEET_H
#define LXLSX_READER_WORKSHEET_H

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_worksheet lxlsx_reader_worksheet;

void lxlsx_reader_worksheet_close(lxlsx_reader_worksheet *ws);

lxlsx_reader_error lxlsx_reader_worksheet_next_row (lxlsx_reader_worksheet *ws);
lxlsx_reader_error lxlsx_reader_worksheet_next_cell(lxlsx_reader_worksheet *ws, lxlsx_cell *out);

size_t   lxlsx_reader_worksheet_current_row    (const lxlsx_reader_worksheet *ws);
size_t   lxlsx_reader_worksheet_max_column_seen(const lxlsx_reader_worksheet *ws);
uint32_t lxlsx_reader_worksheet_flags          (const lxlsx_reader_worksheet *ws);

typedef int (*lxlsx_reader_cell_cb)   (const lxlsx_cell *cell, void *userdata);
typedef int (*lxlsx_reader_row_end_cb)(size_t row, size_t max_col, void *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_process(lxlsx_reader_worksheet *ws,
                                lxlsx_reader_cell_cb    cell_cb,
                                lxlsx_reader_row_end_cb row_cb,
                                void          *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_skip_rows(lxlsx_reader_worksheet *ws, size_t n);

/* ----- Phase 1 metadata accessors ---------------------------------------- */

/* Merged cells. Index is 0-based. */
size_t lxlsx_reader_worksheet_merged_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_merged_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                  lxlsx_reader_range *out);

/* Returns 1 when (row, col) (1-based) lies inside a merged range and is NOT
 * the merge's master cell; 0 otherwise. Useful for honoring
 * LXLSX_READER_SKIP_MERGED_FOLLOW in row-shaped readers. */
int lxlsx_reader_worksheet_in_merge_follow(lxlsx_reader_worksheet *ws, size_t row, size_t col);

/* Hyperlinks. The url/tooltip/display/location pointers (if non-NULL) are
 * owned by the worksheet and remain valid until lxlsx_reader_worksheet_close().
 * Internal links (e.g. "Sheet2!A1") have location set and url == NULL. */
typedef struct {
    lxlsx_reader_range range;
    const char *url;       /* external URL, or NULL */
    const char *location;  /* internal anchor, or NULL */
    const char *display;   /* display text, or NULL */
    const char *tooltip;   /* tooltip text, or NULL */
} lxlsx_reader_hyperlink;

size_t      lxlsx_reader_worksheet_hyperlink_count (const lxlsx_reader_worksheet *ws);
int         lxlsx_reader_worksheet_hyperlink_get   (const lxlsx_reader_worksheet *ws, size_t idx,
                                           lxlsx_reader_hyperlink *out);
const char *lxlsx_reader_worksheet_hyperlink_url   (const lxlsx_reader_worksheet *ws,
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
} lxlsx_reader_protection;

int lxlsx_reader_worksheet_protection(const lxlsx_reader_worksheet *ws, lxlsx_reader_protection *out);

/* Row / column metadata. row/col indexes are 1-based. */
typedef struct {
    int    has_height;
    double height;
    int    hidden;
    int    outline_level;
    int    collapsed;
    int    custom_height;
} lxlsx_reader_row_options;

typedef struct {
    int    has_width;
    double width;
    int    hidden;
    int    outline_level;
    int    collapsed;
} lxlsx_reader_col_options;

int lxlsx_reader_worksheet_row_options(const lxlsx_reader_worksheet *ws, size_t row,
                              lxlsx_reader_row_options *out);
int lxlsx_reader_worksheet_col_options(const lxlsx_reader_worksheet *ws, size_t col,
                              lxlsx_reader_col_options *out);

int    lxlsx_reader_worksheet_default_row_height(const lxlsx_reader_worksheet *ws, double *out);
int    lxlsx_reader_worksheet_default_col_width (const lxlsx_reader_worksheet *ws, double *out);

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
} lxlsx_reader_data_validation;

size_t lxlsx_reader_worksheet_data_validation_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_data_validation_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                           lxlsx_reader_data_validation *out);

/* ---- AutoFilter --------------------------------------------------------- */

typedef enum {
    LXLSX_READER_FILTER_NONE = 0,
    LXLSX_READER_FILTER_LIST,           /* discrete value match (<filters>) */
    LXLSX_READER_FILTER_CUSTOM,         /* operator-based (<customFilters>) */
    LXLSX_READER_FILTER_TOP10,          /* <top10>: top/bottom N (or %) */
    LXLSX_READER_FILTER_DYNAMIC         /* <dynamicFilter>: aboveAverage etc. */
} lxlsx_reader_filter_kind;

typedef struct {
    int             col_id;          /* 0-based column offset within the autoFilter range */
    lxlsx_reader_filter_kind kind;
    /* For LXLSX_READER_FILTER_LIST: list of value strings. */
    const char    **values;          /* NULL-terminated; NULL when not applicable */
    /* For LXLSX_READER_FILTER_CUSTOM: at most two predicates joined by AND/OR. */
    int             custom_and;      /* 1 = AND, 0 = OR */
    const char     *custom_op_1;     /* "equal","greaterThan", ... */
    const char     *custom_val_1;
    const char     *custom_op_2;
    const char     *custom_val_2;
    /* For LXLSX_READER_FILTER_TOP10. */
    int             top;             /* 1=top, 0=bottom */
    int             percent;         /* 1=percent */
    double          top_value;       /* N */
} lxlsx_reader_filter_column;

typedef struct {
    const char              *range;       /* original ref attr (e.g. "A1:Z100"), or NULL */
    const lxlsx_reader_filter_column *columns;
    size_t                   columns_count;
} lxlsx_reader_autofilter;

int lxlsx_reader_worksheet_autofilter(const lxlsx_reader_worksheet *ws, lxlsx_reader_autofilter *out);

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
} lxlsx_reader_page_setup;

int lxlsx_reader_worksheet_page_setup(const lxlsx_reader_worksheet *ws, lxlsx_reader_page_setup *out);

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
} lxlsx_reader_cf_rule;

typedef struct {
    const char        *sqref;     /* the conditionalFormatting sqref */
    const lxlsx_reader_cf_rule *rules;
    size_t             rules_count;
} lxlsx_reader_cf_block;

size_t lxlsx_reader_worksheet_cf_block_count(const lxlsx_reader_worksheet *ws);
int    lxlsx_reader_worksheet_cf_block_get  (const lxlsx_reader_worksheet *ws, size_t idx,
                                    lxlsx_reader_cf_block *out);

/* ---- Comments (§8.2.1) -------------------------------------------------- */

typedef struct {
    size_t      row;          /* 1-based */
    size_t      col;          /* 1-based */
    const char *text;
    const char *author;
    int         visible;      /* legacy VML's "visible" attr */
    int         threaded;
    const char *parent_id;    /* threaded reply parent GUID, or NULL */
} lxlsx_reader_comment_info;

typedef int (*lxlsx_reader_comment_cb)(const lxlsx_reader_comment_info *info, void *userdata);

/* Iterates legacy + threaded comments. cb returning non-zero stops. Returns
 * LXLSX_READER_NO_ERROR on success; ZIP_ENTRY_NOT_FOUND when the sheet has none. */
lxlsx_reader_error lxlsx_reader_worksheet_iterate_comments(lxlsx_reader_worksheet *ws,
                                         lxlsx_reader_comment_cb cb, void *userdata);

/* ---- Chart metadata (§8.2.4) ------------------------------------------- */

typedef struct {
    const char *name;        /* series name, may be NULL */
    const char *categories;  /* category range / formula */
    const char *values;      /* values range */
} lxlsx_reader_chart_series_info;

typedef struct {
    const char *type;        /* "bar"/"line"/"pie"/"area"/"scatter"/"radar"/"doughnut" or NULL */
    const char *title;
    /* 1-based inclusive anchor; 0 if not from a twoCellAnchor. */
    size_t from_row, from_col, to_row, to_col;
    const lxlsx_reader_chart_series_info *series;
    size_t                       series_count;
} lxlsx_reader_chart_meta;

typedef int (*lxlsx_reader_chart_cb)(const lxlsx_reader_chart_meta *info, void *userdata);

lxlsx_reader_error lxlsx_reader_worksheet_iterate_charts(lxlsx_reader_worksheet *ws,
                                       lxlsx_reader_chart_cb cb, void *userdata);

/* ---- Rich-text runs (§8.2.2) ------------------------------------------- */

/* For STRING_CELL (SST) and INLINE_STRING_CELL cells, fills `out`
 * with up to `cap` runs and returns the actual count. When `out` is NULL
 * just returns the run count. Callers can sniff the run count first, then
 * allocate appropriately. The pointers inside each run remain valid until
 * the next streaming step that overwrites the inline buffer (for inline
 * strings) or until the workbook is closed (for SST strings). */
size_t lxlsx_reader_cell_string_runs(const lxlsx_reader_worksheet *ws, const lxlsx_cell *c,
                            lxlsx_reader_string_run *out, size_t cap);

#ifdef __cplusplus
}
#endif

#endif
