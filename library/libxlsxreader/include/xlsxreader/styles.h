#ifndef LXR_STYLES_H
#define LXR_STYLES_H

#include "common.h"

#ifdef __cplusplus
extern "C" {
#endif

typedef struct lxr_styles lxr_styles;

typedef enum {
    LXR_FMT_CATEGORY_GENERAL = 0,
    LXR_FMT_CATEGORY_NUMBER,
    LXR_FMT_CATEGORY_PERCENT,
    LXR_FMT_CATEGORY_DATE,
    LXR_FMT_CATEGORY_TIME,
    LXR_FMT_CATEGORY_DATETIME,
    LXR_FMT_CATEGORY_CURRENCY,
    LXR_FMT_CATEGORY_TEXT,
    LXR_FMT_CATEGORY_CUSTOM
} lxr_fmt_category;

/* 8-byte AARRGGBB hex color, plus terminator. Empty string = unset/auto. */
#define LXR_COLOR_LEN 9

typedef enum {
    LXR_UNDERLINE_NONE   = 0,
    LXR_UNDERLINE_SINGLE = 1,
    LXR_UNDERLINE_DOUBLE = 2,
    LXR_UNDERLINE_SINGLE_ACCOUNTING = 3,
    LXR_UNDERLINE_DOUBLE_ACCOUNTING = 4
} lxr_underline;

typedef struct {
    const char   *name;                 /* font family name (NULL if unset) */
    double        size;                 /* point size; 0 if unset           */
    char          color[LXR_COLOR_LEN]; /* AARRGGBB hex; ""=unset/theme     */
    int           bold;                 /* 0/1                              */
    int           italic;
    int           strike;
    lxr_underline underline;
} lxr_font;

typedef struct {
    const char *pattern_type;           /* "solid"/"none"/"darkGrid"/...    */
    char        fg_color[LXR_COLOR_LEN];
    char        bg_color[LXR_COLOR_LEN];
} lxr_fill;

typedef struct {
    const char *style;                  /* "thin"/"medium"/"none"/...       */
    char        color[LXR_COLOR_LEN];
} lxr_border_side;

typedef struct {
    lxr_border_side left;
    lxr_border_side right;
    lxr_border_side top;
    lxr_border_side bottom;
} lxr_border;

typedef struct {
    const char *horizontal;             /* "general"/"left"/"center"/...    */
    const char *vertical;               /* "top"/"center"/"bottom"/...      */
    int         wrap_text;              /* 0/1 */
    int         indent;
    int         rotation;
} lxr_alignment;

typedef struct {
    uint16_t         num_fmt_id;
    lxr_fmt_category category;
    const char      *format_string;

    uint32_t         font_id;
    uint32_t         fill_id;
    uint32_t         border_id;

    int              has_alignment;
    lxr_alignment    alignment;
    int              locked;            /* protection: locked (default 1)  */
    int              hidden;            /* protection: hidden (default 0)  */
} lxr_xf;

const lxr_xf     *lxr_styles_get_xf    (const lxr_styles *st, uint32_t style_id);
const lxr_font   *lxr_styles_get_font  (const lxr_styles *st, uint32_t font_id);
const lxr_fill   *lxr_styles_get_fill  (const lxr_styles *st, uint32_t fill_id);
const lxr_border *lxr_styles_get_border(const lxr_styles *st, uint32_t border_id);

#ifdef __cplusplus
}
#endif

#endif
