#ifndef INCLUDED_XLSXIO_PRIVATE_H
#define INCLUDED_XLSXIO_PRIVATE_H

#if defined(_MSC_VER) || (defined(__MINGW32__) && !defined(__MINGW64__))
#define strcasecmp _stricmp
#endif
#ifdef _WIN32
#define wcscasecmp _wcsicmp
#endif


#define XLSXIOCHAR XML_Char

#if !defined(XML_UNICODE_WCHAR_T) && !defined(XML_UNICODE)

//UTF-8 version
#define X(s) s
#ifdef _WIN32
#define XML_Char_icmp stricmp
#else
#define XML_Char_icmp strcasecmp
#endif
#define XML_Char_len strlen
#define XML_Char_dup strdup
#define XML_Char_cpy strcpy
#define XML_Char_poscpy(d,p,s,l) memcpy(d + p, s, l)
#define XML_Char_malloc(n) ((char*)malloc(n))
#define XML_Char_realloc(m,n) ((char*)realloc((m), (n)))
#define XML_Char_tol(s) strtol((s), NULL, 10)
#define XML_Char_tod(s) strtod((s), NULL)
#define XML_Char_strtol strtol
#define XML_Char_sscanf sscanf
#define XML_Char_printf printf

#else

//UTF-16 version
#include <wchar.h>
#define X(s) L##s
#define XML_Char_icmp wcscasecmp
#define XML_Char_len wcslen
#define XML_Char_dup wcsdup
#define XML_Char_cpy wcscpy
#define XML_Char_poscpy(d,p,s,l) wmemcpy(d + p, s, l)
#define XML_Char_malloc(n) ((XML_Char*)malloc((n) * sizeof(XML_Char)))
#define XML_Char_realloc(m,n) ((XML_Char*)realloc((m), (n) * sizeof(XML_Char)))
#define XML_Char_tol(s) wcstol((s), NULL, 10)
#define XML_Char_tod(s) wcstod((s), NULL)
#define XML_Char_strtol wcstol
#define XML_Char_sscanf swscanf
#define XML_Char_printf wprintf

#endif


#endif
