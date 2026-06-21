/*
 * libxlsx
 *
 * SPDX-License-Identifier: BSD-2-Clause
 */

/**
 * @file edit.h
 *
 * @brief Snapshot-based XLSX edit session API.
 */
#ifndef __LXLSX_EDIT_H__
#define __LXLSX_EDIT_H__

#include "common.h"
#include "format.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_edit_session lxlsx_edit_session;

lxlsx_edit_session *lxlsx_edit_open(const char *path);
lxlsx_error         lxlsx_edit_save_as(lxlsx_edit_session *session,
                                       const char *path);
void                lxlsx_edit_close(lxlsx_edit_session *session);
size_t              lxlsx_edit_sheet_count(lxlsx_edit_session *session);
const char         *lxlsx_edit_sheet_name(lxlsx_edit_session *session,
                                          size_t index);

/*
 * If sheet_name is NULL, edits target the first worksheet in workbook order.
 */
lxlsx_error lxlsx_edit_set_number(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  double value);

lxlsx_error lxlsx_edit_set_string(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  const char *value);

lxlsx_error lxlsx_edit_set_boolean(lxlsx_edit_session *session,
                                   const char *sheet_name,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col,
                                   int value);

lxlsx_error lxlsx_edit_set_formula(lxlsx_edit_session *session,
                                   const char *sheet_name,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col,
                                   const char *formula,
                                   const char *cached_result);

/*
 * Apply a number-format code to a cell already targeted by a value edit in this
 * session (call after the value setter for the same cell). A new numFmt + xf is
 * injected into styles.xml on save and the cell's s= is repointed to it.
 */
lxlsx_error lxlsx_edit_set_number_format(lxlsx_edit_session *session,
                                         const char *sheet_name,
                                         lxlsx_row_t row,
                                         lxlsx_col_t col,
                                         const char *format_code);

/*
 * Apply a cell format (font/fill/alignment/number-format) to a cell already
 * targeted by a value edit in this session. The format is snapshotted; the
 * style records are injected into styles.xml on save and the cell's s= is
 * repointed. Borders are not yet applied in edit mode.
 */
lxlsx_error lxlsx_edit_set_format(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  const lxlsx_format *format);

/*
 * Add a merged cell range (0-based, inclusive) to a sheet. The range is
 * injected into the sheet's <mergeCells> on save; existing merges are kept.
 */
lxlsx_error lxlsx_edit_set_merge(lxlsx_edit_session *session,
                                 const char *sheet_name,
                                 lxlsx_row_t first_row,
                                 lxlsx_col_t first_col,
                                 lxlsx_row_t last_row,
                                 lxlsx_col_t last_col);

/* Set the width of a 0-based column range; injected into <cols> on save. */
lxlsx_error lxlsx_edit_set_column(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_col_t first_col,
                                  lxlsx_col_t last_col,
                                  double width);

/* Set the height of an existing 0-based row; patched onto its <row> on save. */
lxlsx_error lxlsx_edit_set_row_height(lxlsx_edit_session *session,
                                      const char *sheet_name,
                                      lxlsx_row_t row,
                                      double height);

/*
 * Append a brand-new worksheet to the workbook. `xml` is the complete
 * worksheet part content. On save it becomes a new xl/worksheets/sheetN.xml and
 * is registered in workbook.xml, workbook.xml.rels and [Content_Types].xml.
 */
lxlsx_error lxlsx_edit_add_sheet(lxlsx_edit_session *session,
                                 const char *name,
                                 const char *xml,
                                 size_t xml_len);

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_EDIT_H__ */
