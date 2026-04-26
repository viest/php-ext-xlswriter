#ifndef LXR_INTERNAL_H
#define LXR_INTERNAL_H

#include <stdint.h>

#include "xlsxreader/common.h"
#include "xlsxreader/workbook.h"
#include "xlsxreader/worksheet.h"
#include "xlsxreader/styles.h"

#include "lxr_zip.h"
#include "lxr_sst.h"
#include "lxr_styles_priv.h"
#include "lxr_xml_pump.h"

typedef struct {
    char *name;       /* sheet display name */
    char *rel_id;     /* relationship id, e.g. "rId1" */
    char *target;     /* zip entry path, e.g. "xl/worksheets/sheet1.xml" */
    int   visibility; /* lxr_sheet_visibility */
} lxr_sheet_info;

/* Defined name entry, owned by workbook. */
typedef struct {
    char *name;
    char *formula;
    int   scope_sheet_index;
    int   hidden;
} lxr_defined_name_entry;

struct lxr_workbook {
    lxr_zip          *zip;
    int               uses_1904;

    lxr_sheet_info   *sheets;
    size_t            sheet_count;
    size_t            sheet_cap;

    char             *workbook_path;
    char             *base_path;
    char             *sst_path;
    char             *styles_path;

    lxr_sst          *sst;
    lxr_styles       *styles;

    lxr_open_options  opts;

    /* defined names */
    lxr_defined_name_entry *defined_names;
    size_t                  defined_name_count;
    size_t                  defined_name_cap;
};

/* Per-worksheet metadata cached at sheet open (mergeCells, hyperlinks, sheet
 * protection, row/col options, default sizes). Loaded eagerly via a
 * dedicated parsing pass that ignores <c> children inside <sheetData>. */
typedef struct {
    /* Defaults from <sheetFormatPr>. */
    int    has_default_row_height;
    double default_row_height;
    int    has_default_col_width;
    double default_col_width;

    /* Merged cells. */
    lxr_range *merges;
    size_t     merges_count;
    size_t     merges_cap;

    /* Hyperlinks (owned). url_storage / etc. point into a pool of strings. */
    struct lxr_hyperlink_owned {
        lxr_range range;
        char     *url;
        char     *location;
        char     *display;
        char     *tooltip;
    } *hyperlinks;
    size_t hyperlinks_count;
    size_t hyperlinks_cap;

    /* Sheet protection (defined in xlsxreader/worksheet.h). */
    int  prot_present;
    char prot_hash[20];
    int  prot_sheet, prot_content, prot_objects, prot_scenarios;
    int  prot_format_cells, prot_format_columns, prot_format_rows;
    int  prot_insert_columns, prot_insert_rows, prot_insert_hyperlinks;
    int  prot_delete_columns, prot_delete_rows;
    int  prot_select_locked_cells, prot_sort, prot_auto_filter;
    int  prot_pivot_tables, prot_select_unlocked_cells;

    /* Row metadata: sparse, indexed by row number 1-based. */
    struct lxr_row_meta {
        size_t row;
        int    has_height;
        double height;
        int    hidden;
        int    outline_level;
        int    collapsed;
        int    custom_height;
    } *rows;
    size_t rows_count;
    size_t rows_cap;

    /* Column metadata: each entry covers [min, max]. */
    struct lxr_col_meta {
        size_t min;
        size_t max;
        int    has_width;
        double width;
        int    hidden;
        int    outline_level;
        int    collapsed;
    } *cols;
    size_t cols_count;
    size_t cols_cap;
} lxr_worksheet_meta;

/* Worksheet FSM states (per plans/reader.md §8). */
typedef enum {
    LXR_WS_INIT = 0,
    LXR_WS_IN_WORKSHEET,
    LXR_WS_IN_SHEETDATA,
    LXR_WS_IN_ROW,
    LXR_WS_IN_CELL,
    LXR_WS_IN_VALUE,
    LXR_WS_IN_FORMULA,
    LXR_WS_IN_INLINE_STR,
    LXR_WS_IN_INLINE_STR_T,
    LXR_WS_SKIP
} lxr_ws_state;

/* What the pull-mode caller is currently waiting for. */
typedef enum {
    LXR_WS_PULL_NONE = 0,    /* no pull active (callback mode or idle) */
    LXR_WS_PULL_ROW,         /* waiting for next <row> */
    LXR_WS_PULL_CELL         /* waiting for next </c> or </row> */
} lxr_ws_pull_mode;

struct lxr_worksheet {
    lxr_workbook *wb;            /* not owned */
    lxr_zip_file *zf;            /* owned */
    lxr_xml_pump *pump;          /* owned */
    char         *target_path;   /* owned: e.g. "xl/worksheets/sheet1.xml" */

    uint32_t      flags;

    /* FSM state */
    lxr_ws_state  state;
    lxr_ws_state  state_before_skip;
    int           skip_depth;
    char         *skip_tag;

    /* Per-row state */
    size_t        row_nr;
    size_t        max_col_seen;
    int           row_in_progress;
    int           row_hidden;
    int           pending_row_start;   /* row started, awaiting consumer ack */
    int           pending_row_end;     /* row ended, awaiting consumer */

    /* Cell being assembled */
    char          cell_t[16];
    char          cell_ref[32];
    uint32_t      cell_style_id;
    size_t        cell_row;
    size_t        cell_col;

    char         *cell_value;
    size_t        cell_value_len;
    size_t        cell_value_cap;

    char         *cell_formula;
    size_t        cell_formula_len;
    size_t        cell_formula_cap;

    char         *cell_inline;
    size_t        cell_inline_len;
    size_t        cell_inline_cap;

    int           cell_has_formula;
    int           cell_has_inline;

    int           pending_cell;        /* a complete cell awaits consumer */

    /* Pull-mode dispatch */
    lxr_ws_pull_mode pull_mode;

    /* Push-mode (lxr_worksheet_process) */
    lxr_cell_cb     user_cell_cb;
    lxr_row_end_cb  user_row_cb;
    void           *user_data;
    int             callback_stop;

    /* Skip-rows budget */
    size_t          skip_rows_remaining;

    int             eof;

    /* Phase 1 metadata cache (eager, populated at open time). */
    lxr_worksheet_meta meta;
};

/* worksheet.c */
lxr_error lxr_worksheet_open_internal(lxr_workbook *wb,
                                      const char *target_path,
                                      uint32_t flags,
                                      lxr_worksheet **out);

/* worksheet_meta.c */
lxr_error lxr_worksheet_meta_load(lxr_worksheet *ws);
void      lxr_worksheet_meta_free(lxr_worksheet_meta *m);

/* numeric helpers */
int64_t lxr_excel_serial_to_unix(double serial, int uses_1904);

#endif
