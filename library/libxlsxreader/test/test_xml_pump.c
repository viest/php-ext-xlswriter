#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "lxr_xml_pump.h"
#include "lxr_zip.h"

void setUp(void) {}
void tearDown(void) {}

/* ------------------------------------------------------------------------- */
/* helpers: counting handler                                                 */
/* ------------------------------------------------------------------------- */

typedef struct {
    int n_start;
    int n_end;
    int n_text;
    int saw_root;
    int saw_namespaced;
    char last_name[64];
    char text_buf[256];
    size_t text_len;
} counter_state;

static void on_start(void *ud, const char *name, const char **attrs)
{
    counter_state *s = (counter_state *)ud;
    s->n_start++;
    strncpy(s->last_name, name, sizeof(s->last_name) - 1);
    s->last_name[sizeof(s->last_name) - 1] = 0;
    if (lxr_xml_name_eq(name, "root")) s->saw_root = 1;
    if (lxr_xml_name_eq(name, "child")) s->saw_namespaced = 1;
    (void)attrs;
}

static void on_end(void *ud, const char *name)
{
    counter_state *s = (counter_state *)ud;
    s->n_end++;
    (void)name;
}

static void on_text(void *ud, const char *text, int len)
{
    counter_state *s = (counter_state *)ud;
    s->n_text++;
    if (s->text_len + (size_t)len < sizeof(s->text_buf) - 1) {
        memcpy(s->text_buf + s->text_len, text, (size_t)len);
        s->text_len += (size_t)len;
        s->text_buf[s->text_len] = 0;
    }
}

/* ------------------------------------------------------------------------- */
/* buffer-source basic parsing                                               */
/* ------------------------------------------------------------------------- */

static void test_parse_simple_buffer(void)
{
    const char *xml = "<root><a>hi</a><b>bye</b></root>";
    counter_state st = {0};
    lxr_xml_pump *p = lxr_xml_pump_create_buffer(xml, strlen(xml));
    TEST_ASSERT_NOT_NULL(p);
    lxr_xml_pump_set_handlers(p, on_start, on_end, on_text, &st);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_run(p));
    TEST_ASSERT_TRUE(lxr_xml_pump_is_eof(p));
    TEST_ASSERT_EQUAL_INT(3, st.n_start);
    TEST_ASSERT_EQUAL_INT(3, st.n_end);
    TEST_ASSERT_TRUE(st.saw_root);
    TEST_ASSERT_NOT_NULL(strstr(st.text_buf, "hi"));
    TEST_ASSERT_NOT_NULL(strstr(st.text_buf, "bye"));
    lxr_xml_pump_destroy(p);
}

static void test_parse_namespace_stripped(void)
{
    const char *xml =
        "<root xmlns:x=\"urn:x\"><x:child/></root>";
    counter_state st = {0};
    lxr_xml_pump *p = lxr_xml_pump_create_buffer(xml, strlen(xml));
    TEST_ASSERT_NOT_NULL(p);
    lxr_xml_pump_set_handlers(p, on_start, on_end, on_text, &st);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_run(p));
    TEST_ASSERT_TRUE(st.saw_namespaced);
    lxr_xml_pump_destroy(p);
}

static void test_parse_error_returns_xml_parse(void)
{
    const char *xml = "<root>unclosed";
    counter_state st = {0};
    lxr_xml_pump *p = lxr_xml_pump_create_buffer(xml, strlen(xml));
    TEST_ASSERT_NOT_NULL(p);
    lxr_xml_pump_set_handlers(p, on_start, on_end, on_text, &st);
    TEST_ASSERT_EQUAL_INT(LXR_ERROR_XML_PARSE, lxr_xml_pump_run(p));
    lxr_xml_pump_destroy(p);
}

/* ------------------------------------------------------------------------- */
/* attr helper                                                               */
/* ------------------------------------------------------------------------- */

typedef struct {
    char target_attr[32];
    char captured_value[32];
    int  found;
} attr_state;

static void capture_start(void *ud, const char *name, const char **attrs)
{
    attr_state *s = (attr_state *)ud;
    if (lxr_xml_name_eq(name, "cell")) {
        const char *v = lxr_xml_attr(attrs, s->target_attr);
        if (v) {
            strncpy(s->captured_value, v, sizeof(s->captured_value) - 1);
            s->captured_value[sizeof(s->captured_value) - 1] = 0;
            s->found = 1;
        }
    }
}

static void test_attr_lookup(void)
{
    const char *xml = "<sheet><cell r=\"A1\" t=\"s\" s=\"3\"/></sheet>";
    attr_state st;
    lxr_xml_pump *p;

    memset(&st, 0, sizeof(st));
    strcpy(st.target_attr, "t");
    p = lxr_xml_pump_create_buffer(xml, strlen(xml));
    TEST_ASSERT_NOT_NULL(p);
    lxr_xml_pump_set_handlers(p, capture_start, NULL, NULL, &st);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_run(p));
    TEST_ASSERT_TRUE(st.found);
    TEST_ASSERT_EQUAL_STRING("s", st.captured_value);
    lxr_xml_pump_destroy(p);
}

/* ------------------------------------------------------------------------- */
/* suspend / resume                                                          */
/* ------------------------------------------------------------------------- */

typedef struct {
    int row_starts;
    int suspend_after_row;
    lxr_xml_pump *pump;
} suspend_state;

static void suspend_start(void *ud, const char *name, const char **attrs)
{
    suspend_state *s = (suspend_state *)ud;
    if (lxr_xml_name_eq(name, "row")) {
        s->row_starts++;
        if (s->suspend_after_row) {
            lxr_xml_pump_suspend(s->pump);
        }
    }
    (void)attrs;
}

static void test_suspend_and_resume(void)
{
    const char *xml =
        "<sheetData>"
        "<row r=\"1\"><c/></row>"
        "<row r=\"2\"><c/></row>"
        "<row r=\"3\"><c/></row>"
        "</sheetData>";
    suspend_state st = {0, 1, NULL};
    lxr_xml_pump *p = lxr_xml_pump_create_buffer(xml, strlen(xml));
    int iterations = 0;

    TEST_ASSERT_NOT_NULL(p);
    st.pump = p;
    lxr_xml_pump_set_handlers(p, suspend_start, NULL, NULL, &st);

    /* Initial run: suspends after the first <row> */
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_run(p));
    TEST_ASSERT_TRUE(lxr_xml_pump_is_suspended(p));
    TEST_ASSERT_EQUAL_INT(1, st.row_starts);

    while (lxr_xml_pump_is_suspended(p) && !lxr_xml_pump_is_eof(p)) {
        TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_resume(p));
        if (++iterations > 10) TEST_FAIL_MESSAGE("infinite loop");
    }

    TEST_ASSERT_EQUAL_INT(3, st.row_starts);
    TEST_ASSERT_TRUE(lxr_xml_pump_is_eof(p));
    lxr_xml_pump_destroy(p);
}

/* ------------------------------------------------------------------------- */
/* zip-source: drive on a real xlsx workbook.xml                             */
/* ------------------------------------------------------------------------- */

static void test_pump_on_zip_file(void)
{
    counter_state st = {0};
    lxr_zip *z = lxr_zip_open_path(LXR_TEST_HIDDEN_ROW_XLSX);
    lxr_zip_file *zf;
    lxr_xml_pump *p;
    TEST_ASSERT_NOT_NULL(z);
    zf = lxr_zip_open_entry(z, "xl/workbook.xml");
    TEST_ASSERT_NOT_NULL(zf);

    p = lxr_xml_pump_create_zip_file(zf);
    TEST_ASSERT_NOT_NULL(p);
    lxr_xml_pump_set_handlers(p, on_start, on_end, on_text, &st);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_xml_pump_run(p));
    TEST_ASSERT_TRUE(lxr_xml_pump_is_eof(p));
    TEST_ASSERT_GREATER_THAN(0, st.n_start);
    TEST_ASSERT_EQUAL_INT(st.n_start, st.n_end);

    lxr_xml_pump_destroy(p);
    lxr_zip_close_entry(zf);
    lxr_zip_close(z);
}

/* ------------------------------------------------------------------------- */
/* helper helpers                                                            */
/* ------------------------------------------------------------------------- */

static void test_name_eq_basic(void)
{
    TEST_ASSERT_TRUE (lxr_xml_name_eq("c", "c"));
    TEST_ASSERT_TRUE (lxr_xml_name_eq("x:c", "c"));
    TEST_ASSERT_FALSE(lxr_xml_name_eq("xc", "c"));
    TEST_ASSERT_FALSE(lxr_xml_name_eq("c", "x:c"));
    TEST_ASSERT_FALSE(lxr_xml_name_eq(NULL, "c"));
    TEST_ASSERT_FALSE(lxr_xml_name_eq("c", NULL));
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_parse_simple_buffer);
    RUN_TEST(test_parse_namespace_stripped);
    RUN_TEST(test_parse_error_returns_xml_parse);
    RUN_TEST(test_attr_lookup);
    RUN_TEST(test_suspend_and_resume);
    RUN_TEST(test_pump_on_zip_file);
    RUN_TEST(test_name_eq_basic);
    return UNITY_END();
}
