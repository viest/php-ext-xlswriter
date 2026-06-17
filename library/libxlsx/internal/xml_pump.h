#ifndef LXLSX_XML_PUMP_H
#define LXLSX_XML_PUMP_H

#include <stddef.h>

#include "platform.h"
#include "libxlsx/common.h"
#include "zip_io.h"

typedef struct lxlsx_reader_xml_pump lxlsx_reader_xml_pump;

typedef void (*lxlsx_reader_xml_start_fn)(void *ud, const char *name, const char **attrs);
typedef void (*lxlsx_reader_xml_end_fn)  (void *ud, const char *name);
typedef void (*lxlsx_reader_xml_text_fn) (void *ud, const char *text, int len);

/* Reader callback abstraction. Returns bytes read, 0 on EOF, -1 on error. */
typedef ssize_t (*lxlsx_reader_xml_read_fn)(void *userdata, void *buf, size_t n);

lxlsx_reader_xml_pump *lxlsx_reader_xml_pump_create_callback(lxlsx_reader_xml_read_fn read_fn,
                                           void           *read_ud);
lxlsx_reader_xml_pump *lxlsx_reader_xml_pump_create_zip_file(lxlsx_reader_zip_file *zf);
lxlsx_reader_xml_pump *lxlsx_reader_xml_pump_create_buffer  (const char *data, size_t len);

void lxlsx_reader_xml_pump_destroy(lxlsx_reader_xml_pump *p);

void lxlsx_reader_xml_pump_set_handlers(lxlsx_reader_xml_pump    *p,
                               lxlsx_reader_xml_start_fn s,
                               lxlsx_reader_xml_end_fn   e,
                               lxlsx_reader_xml_text_fn  t,
                               void            *userdata);

/* Parse to completion (or until handler suspends). */
lxlsx_reader_error lxlsx_reader_xml_pump_run(lxlsx_reader_xml_pump *p);

/* Resume after suspend. */
lxlsx_reader_error lxlsx_reader_xml_pump_resume(lxlsx_reader_xml_pump *p);

/* Suspend at the current event boundary; call from inside a handler. */
void lxlsx_reader_xml_pump_suspend(lxlsx_reader_xml_pump *p);

/* True after run/resume returns because of suspend (vs. EOF / error). */
int  lxlsx_reader_xml_pump_is_suspended(const lxlsx_reader_xml_pump *p);

/* True after EOF reached. */
int  lxlsx_reader_xml_pump_is_eof(const lxlsx_reader_xml_pump *p);

/* Helpers. */
const char *lxlsx_reader_xml_attr   (const char **attrs, const char *name);
int         lxlsx_reader_xml_name_eq(const char *xml_name, const char *want);

#endif
