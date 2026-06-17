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

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_edit_session lxlsx_edit_session;

lxlsx_edit_session *lxlsx_edit_open(const char *path);
lxlsx_error         lxlsx_edit_save_as(lxlsx_edit_session *session,
                                       const char *path);
void                lxlsx_edit_close(lxlsx_edit_session *session);

lxlsx_error lxlsx_edit_set_number(lxlsx_edit_session *session,
                                  const char *sheet_name,
                                  lxlsx_row_t row,
                                  lxlsx_col_t col,
                                  double value);

lxlsx_error lxlsx_edit_set_formula(lxlsx_edit_session *session,
                                   const char *sheet_name,
                                   lxlsx_row_t row,
                                   lxlsx_col_t col,
                                   const char *formula,
                                   const char *cached_result);

#ifdef __cplusplus
}
#endif

#endif /* __LXLSX_EDIT_H__ */
