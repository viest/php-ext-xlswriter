#ifndef LXLSX_READER_STYLES_H
#define LXLSX_READER_STYLES_H

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxlsx_reader_styles lxlsx_reader_styles;

typedef enum {
    LXLSX_READER_FMT_CATEGORY_GENERAL = 0,
    LXLSX_READER_FMT_CATEGORY_NUMBER,
    LXLSX_READER_FMT_CATEGORY_PERCENT,
    LXLSX_READER_FMT_CATEGORY_DATE,
    LXLSX_READER_FMT_CATEGORY_TIME,
    LXLSX_READER_FMT_CATEGORY_DATETIME,
    LXLSX_READER_FMT_CATEGORY_CURRENCY,
    LXLSX_READER_FMT_CATEGORY_TEXT,
    LXLSX_READER_FMT_CATEGORY_CUSTOM
} lxlsx_reader_fmt_category;

/* 8-byte AARRGGBB hex color, plus terminator. Empty string = unset/auto. */
#define LXLSX_READER_COLOR_LEN 9

typedef enum {
    LXLSX_READER_UNDERLINE_NONE   = 0,
    LXLSX_READER_UNDERLINE_SINGLE = 1,
    LXLSX_READER_UNDERLINE_DOUBLE = 2,
    LXLSX_READER_UNDERLINE_SINGLE_ACCOUNTING = 3,
    LXLSX_READER_UNDERLINE_DOUBLE_ACCOUNTING = 4
} lxlsx_reader_underline;

typedef struct {
    const char   *name;                 /* font family name (NULL if unset) */
    double        size;                 /* point size; 0 if unset           */
    char          color[LXLSX_READER_COLOR_LEN]; /* AARRGGBB hex; ""=unset/theme     */
    int           bold;                 /* 0/1                              */
    int           italic;
    int           strike;
    lxlsx_reader_underline underline;
} lxlsx_reader_font;

typedef struct {
    const char *pattern_type;           /* "solid"/"none"/"darkGrid"/...    */
    char        fg_color[LXLSX_READER_COLOR_LEN];
    char        bg_color[LXLSX_READER_COLOR_LEN];
} lxlsx_reader_fill;

typedef struct {
    const char *style;                  /* "thin"/"medium"/"none"/...       */
    char        color[LXLSX_READER_COLOR_LEN];
} lxlsx_reader_border_side;

typedef struct {
    lxlsx_reader_border_side left;
    lxlsx_reader_border_side right;
    lxlsx_reader_border_side top;
    lxlsx_reader_border_side bottom;
} lxlsx_reader_border;

typedef struct {
    const char *horizontal;             /* "general"/"left"/"center"/...    */
    const char *vertical;               /* "top"/"center"/"bottom"/...      */
    int         wrap_text;              /* 0/1 */
    int         indent;
    int         rotation;
} lxlsx_reader_alignment;

typedef struct {
    uint16_t         num_fmt_id;
    lxlsx_reader_fmt_category category;
    const char      *format_string;

    uint32_t         font_id;
    uint32_t         fill_id;
    uint32_t         border_id;

    int              has_alignment;
    lxlsx_reader_alignment    alignment;
    int              locked;            /* protection: locked (default 1)  */
    int              hidden;            /* protection: hidden (default 0)  */
} lxlsx_reader_xf;

const lxlsx_reader_xf     *lxlsx_reader_styles_get_xf    (const lxlsx_reader_styles *st, uint32_t style_id);
const lxlsx_reader_font   *lxlsx_reader_styles_get_font  (const lxlsx_reader_styles *st, uint32_t font_id);
const lxlsx_reader_fill   *lxlsx_reader_styles_get_fill  (const lxlsx_reader_styles *st, uint32_t fill_id);
const lxlsx_reader_border *lxlsx_reader_styles_get_border(const lxlsx_reader_styles *st, uint32_t border_id);

#ifdef __cplusplus
}
#endif

#endif
