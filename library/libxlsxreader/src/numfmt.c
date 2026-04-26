#include <ctype.h>
#include <string.h>

#include "lxr_numfmt.h"

/* OOXML 18.8.30 built-in numFmtIds. Locale-dependent slots
 * (5..8, 23..36, 41..44) are intentionally NULL — they fall through to the
 * GENERAL category and are treated as numbers. */
static const struct {
    uint16_t        id;
    const char      *fmt;
    lxr_fmt_category cat;
} BUILTIN_NUMFMTS[] = {
    {  0, "General",           LXR_FMT_CATEGORY_GENERAL  },
    {  1, "0",                 LXR_FMT_CATEGORY_NUMBER   },
    {  2, "0.00",              LXR_FMT_CATEGORY_NUMBER   },
    {  3, "#,##0",             LXR_FMT_CATEGORY_NUMBER   },
    {  4, "#,##0.00",          LXR_FMT_CATEGORY_NUMBER   },
    {  9, "0%",                LXR_FMT_CATEGORY_PERCENT  },
    { 10, "0.00%",             LXR_FMT_CATEGORY_PERCENT  },
    { 11, "0.00E+00",          LXR_FMT_CATEGORY_NUMBER   },
    { 12, "# ?/?",             LXR_FMT_CATEGORY_NUMBER   },
    { 13, "# ?\?/?\?",         LXR_FMT_CATEGORY_NUMBER   },
    { 14, "m/d/yyyy",          LXR_FMT_CATEGORY_DATE     },
    { 15, "d-mmm-yy",          LXR_FMT_CATEGORY_DATE     },
    { 16, "d-mmm",             LXR_FMT_CATEGORY_DATE     },
    { 17, "mmm-yy",            LXR_FMT_CATEGORY_DATE     },
    { 18, "h:mm AM/PM",        LXR_FMT_CATEGORY_TIME     },
    { 19, "h:mm:ss AM/PM",     LXR_FMT_CATEGORY_TIME     },
    { 20, "h:mm",              LXR_FMT_CATEGORY_TIME     },
    { 21, "h:mm:ss",           LXR_FMT_CATEGORY_TIME     },
    { 22, "m/d/yyyy h:mm",     LXR_FMT_CATEGORY_DATETIME },
    { 37, "#,##0 ;(#,##0)",                LXR_FMT_CATEGORY_NUMBER },
    { 38, "#,##0 ;[Red](#,##0)",           LXR_FMT_CATEGORY_NUMBER },
    { 39, "#,##0.00;(#,##0.00)",           LXR_FMT_CATEGORY_NUMBER },
    { 40, "#,##0.00;[Red](#,##0.00)",      LXR_FMT_CATEGORY_NUMBER },
    { 45, "mm:ss",             LXR_FMT_CATEGORY_TIME     },
    { 46, "[h]:mm:ss",         LXR_FMT_CATEGORY_TIME     },
    { 47, "mmss.0",            LXR_FMT_CATEGORY_TIME     },
    { 48, "##0.0E+0",          LXR_FMT_CATEGORY_NUMBER   },
    { 49, "@",                 LXR_FMT_CATEGORY_TEXT     },
};

const char *lxr_numfmt_builtin_format(uint16_t id)
{
    size_t i;
    for (i = 0; i < sizeof(BUILTIN_NUMFMTS) / sizeof(BUILTIN_NUMFMTS[0]); i++) {
        if (BUILTIN_NUMFMTS[i].id == id) return BUILTIN_NUMFMTS[i].fmt;
    }
    return NULL;
}

lxr_fmt_category lxr_numfmt_builtin_category(uint16_t id)
{
    size_t i;
    for (i = 0; i < sizeof(BUILTIN_NUMFMTS) / sizeof(BUILTIN_NUMFMTS[0]); i++) {
        if (BUILTIN_NUMFMTS[i].id == id) return BUILTIN_NUMFMTS[i].cat;
    }
    return LXR_FMT_CATEGORY_GENERAL;
}

/* Classify a custom format string into a category. Heuristic — strict from
 * date/time markers down to general. Skips characters inside quoted literals
 * ("..."), bracketed colour/condition tags ([Red], [>=0]), and escaped
 * characters (\x). */
lxr_fmt_category lxr_numfmt_classify(const char *fmt)
{
    int has_y = 0, has_m_letter = 0, has_d = 0;
    int has_h = 0, has_s = 0;
    int has_percent = 0, has_currency = 0, has_at = 0;
    int has_digit = 0;
    const char *p;

    if (!fmt) return LXR_FMT_CATEGORY_GENERAL;

    for (p = fmt; *p; p++) {
        char c = *p;
        if (c == '"') {
            p++;
            while (*p && *p != '"') p++;
            if (!*p) break;
            continue;
        }
        if (c == '[') {
            while (*p && *p != ']') p++;
            if (!*p) break;
            continue;
        }
        if (c == '\\' && *(p + 1)) {
            p++;
            continue;
        }
        switch (c) {
        case 'y': case 'Y': has_y = 1; break;
        case 'd': case 'D': has_d = 1; break;
        case 'm': case 'M': has_m_letter = 1; break;
        case 'h': case 'H': has_h = 1; break;
        case 's': case 'S': has_s = 1; break;
        case '%':           has_percent  = 1; break;
        case '@':           has_at       = 1; break;
        case '$': case '\xa3': /* £ */
                            has_currency = 1; break;
        case '0': case '#': case '?': has_digit = 1; break;
        default: break;
        }
    }

    /* Multi-byte currency symbols (UTF-8 encoded ¥, €) — coarse detection:
     * any 0xC2..0xF4 byte not matched above is treated as a hint of currency. */
    for (p = fmt; *p; p++) {
        unsigned char b = (unsigned char)*p;
        if (b >= 0xC2 && b <= 0xF4) { has_currency = 1; break; }
    }

    if (has_y || has_d) {
        if (has_h || has_s) return LXR_FMT_CATEGORY_DATETIME;
        return LXR_FMT_CATEGORY_DATE;
    }
    if (has_h || has_s) return LXR_FMT_CATEGORY_TIME;
    if (has_m_letter && (has_h || has_s)) return LXR_FMT_CATEGORY_TIME;

    if (has_percent)  return LXR_FMT_CATEGORY_PERCENT;
    if (has_currency) return LXR_FMT_CATEGORY_CURRENCY;
    if (has_at && !has_digit) return LXR_FMT_CATEGORY_TEXT;
    if (has_digit)    return LXR_FMT_CATEGORY_NUMBER;

    return LXR_FMT_CATEGORY_GENERAL;
}
