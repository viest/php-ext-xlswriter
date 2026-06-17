/*
 * libxlsxwriter
 *
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * rich_value - A libxlsxwriter library for creating Excel XLSX rich_value files.
 *
 */
#ifndef __LXW_RICH_VALUE_H__
#define __LXW_RICH_VALUE_H__

#include <stdint.h>

#include "common.h"
#include "workbook.h"

/*
 * Struct to represent a rich_value object.
 */
typedef struct lxw_rich_value {

    FILE *file;
    lxw_workbook *workbook;

} lxw_rich_value;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

lxw_rich_value *lxw_rich_value_new(void);
void lxw_rich_value_free(lxw_rich_value *rich_value);
void lxw_rich_value_assemble_xml_file(lxw_rich_value *self);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _rich_value_xml_declaration(lxw_rich_value *self);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_RICH_VALUE_H__ */
