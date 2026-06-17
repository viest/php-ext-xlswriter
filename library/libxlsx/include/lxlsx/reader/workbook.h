#ifndef LXLSX_READER_WORKBOOK_H
#define LXLSX_READER_WORKBOOK_H

#include "common.h"
#include "worksheet.h"
#include "styles.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_workbook lxlsx_reader_workbook;

typedef struct {
    lxlsx_reader_sst_mode sst_mode;
} lxlsx_reader_open_options;

lxlsx_reader_error lxlsx_reader_workbook_open       (const char *filename, lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_fd    (int fd,                lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_memory(const void *data, size_t len, lxlsx_reader_workbook **out);
lxlsx_reader_error lxlsx_reader_workbook_open_ex    (const char *filename,
                                   const lxlsx_reader_open_options *opts,
                                   lxlsx_reader_workbook **out);
void      lxlsx_reader_workbook_close      (lxlsx_reader_workbook *wb);

size_t            lxlsx_reader_workbook_sheet_count(const lxlsx_reader_workbook *wb);
const char       *lxlsx_reader_workbook_sheet_name (const lxlsx_reader_workbook *wb, size_t index);

typedef enum {
    LXLSX_READER_SHEET_VISIBLE      = 0,
    LXLSX_READER_SHEET_HIDDEN       = 1,
    LXLSX_READER_SHEET_VERY_HIDDEN  = 2
} lxlsx_reader_sheet_visibility;

lxlsx_reader_sheet_visibility lxlsx_reader_workbook_sheet_visibility(const lxlsx_reader_workbook *wb, size_t index);

/* Defined names (a.k.a. named ranges). scope_sheet_index is -1 for
 * workbook-scope; otherwise the 0-based sheet index it's bound to. */
typedef struct {
    const char *name;
    const char *formula;
    int         scope_sheet_index;
    int         hidden;
} lxlsx_reader_defined_name;

size_t lxlsx_reader_workbook_defined_name_count(const lxlsx_reader_workbook *wb);
int    lxlsx_reader_workbook_defined_name_get  (const lxlsx_reader_workbook *wb, size_t idx,
                                       lxlsx_reader_defined_name *out);

lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_name (lxlsx_reader_workbook   *wb,
                                              const char     *name,
                                              uint32_t        flags,
                                              lxlsx_reader_worksheet **out);
lxlsx_reader_error lxlsx_reader_workbook_get_worksheet_by_index(lxlsx_reader_workbook   *wb,
                                              size_t          index,
                                              uint32_t        flags,
                                              lxlsx_reader_worksheet **out);

int               lxlsx_reader_workbook_uses_1904_dates(const lxlsx_reader_workbook *wb);
const lxlsx_reader_styles *lxlsx_reader_workbook_get_styles     (const lxlsx_reader_workbook *wb);

#ifdef __cplusplus
}
#endif

#endif
