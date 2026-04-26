#ifndef LXR_WORKBOOK_H
#define LXR_WORKBOOK_H

#include "common.h"
#include "worksheet.h"
#include "styles.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxr_workbook lxr_workbook;

typedef struct {
    lxr_sst_mode sst_mode;
} lxr_open_options;

lxr_error lxr_workbook_open       (const char *filename, lxr_workbook **out);
lxr_error lxr_workbook_open_fd    (int fd,                lxr_workbook **out);
lxr_error lxr_workbook_open_memory(const void *data, size_t len, lxr_workbook **out);
lxr_error lxr_workbook_open_ex    (const char *filename,
                                   const lxr_open_options *opts,
                                   lxr_workbook **out);
void      lxr_workbook_close      (lxr_workbook *wb);

size_t            lxr_workbook_sheet_count(const lxr_workbook *wb);
const char       *lxr_workbook_sheet_name (const lxr_workbook *wb, size_t index);

typedef enum {
    LXR_SHEET_VISIBLE      = 0,
    LXR_SHEET_HIDDEN       = 1,
    LXR_SHEET_VERY_HIDDEN  = 2
} lxr_sheet_visibility;

lxr_sheet_visibility lxr_workbook_sheet_visibility(const lxr_workbook *wb, size_t index);

/* Defined names (a.k.a. named ranges). scope_sheet_index is -1 for
 * workbook-scope; otherwise the 0-based sheet index it's bound to. */
typedef struct {
    const char *name;
    const char *formula;
    int         scope_sheet_index;
    int         hidden;
} lxr_defined_name;

size_t lxr_workbook_defined_name_count(const lxr_workbook *wb);
int    lxr_workbook_defined_name_get  (const lxr_workbook *wb, size_t idx,
                                       lxr_defined_name *out);

lxr_error lxr_workbook_get_worksheet_by_name (lxr_workbook   *wb,
                                              const char     *name,
                                              uint32_t        flags,
                                              lxr_worksheet **out);
lxr_error lxr_workbook_get_worksheet_by_index(lxr_workbook   *wb,
                                              size_t          index,
                                              uint32_t        flags,
                                              lxr_worksheet **out);

int               lxr_workbook_uses_1904_dates(const lxr_workbook *wb);
const lxr_styles *lxr_workbook_get_styles     (const lxr_workbook *wb);

#ifdef __cplusplus
}
#endif

#endif
