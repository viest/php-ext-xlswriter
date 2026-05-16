#ifndef LXR_XML_PUMP_H
#define LXR_XML_PUMP_H

#include <stddef.h>

#include "lxr_compat.h"
#include "xlsxreader/common.h"
#include "lxr_zip.h"

typedef struct lxr_xml_pump lxr_xml_pump;

typedef void (*lxr_xml_start_fn)(void *ud, const char *name, const char **attrs);
typedef void (*lxr_xml_end_fn)  (void *ud, const char *name);
typedef void (*lxr_xml_text_fn) (void *ud, const char *text, int len);

/* Reader callback abstraction. Returns bytes read, 0 on EOF, -1 on error. */
typedef ssize_t (*lxr_xml_read_fn)(void *userdata, void *buf, size_t n);

lxr_xml_pump *lxr_xml_pump_create_callback(lxr_xml_read_fn read_fn,
                                           void           *read_ud);
lxr_xml_pump *lxr_xml_pump_create_zip_file(lxr_zip_file *zf);
lxr_xml_pump *lxr_xml_pump_create_buffer  (const char *data, size_t len);

void lxr_xml_pump_destroy(lxr_xml_pump *p);

void lxr_xml_pump_set_handlers(lxr_xml_pump    *p,
                               lxr_xml_start_fn s,
                               lxr_xml_end_fn   e,
                               lxr_xml_text_fn  t,
                               void            *userdata);

/* Parse to completion (or until handler suspends). */
lxr_error lxr_xml_pump_run(lxr_xml_pump *p);

/* Resume after suspend. */
lxr_error lxr_xml_pump_resume(lxr_xml_pump *p);

/* Suspend at the current event boundary; call from inside a handler. */
void lxr_xml_pump_suspend(lxr_xml_pump *p);

/* True after run/resume returns because of suspend (vs. EOF / error). */
int  lxr_xml_pump_is_suspended(const lxr_xml_pump *p);

/* True after EOF reached. */
int  lxr_xml_pump_is_eof(const lxr_xml_pump *p);

/* Helpers. */
const char *lxr_xml_attr   (const char **attrs, const char *name);
int         lxr_xml_name_eq(const char *xml_name, const char *want);

#endif
