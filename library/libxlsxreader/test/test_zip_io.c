#include <fcntl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>

#include <unity.h>

#include "lxr_test_paths.h"
#include "lxr_zip.h"

void setUp(void) {}
void tearDown(void) {}

static const char *XLSX = LXR_TEST_HIDDEN_ROW_XLSX;

/* ------------------------------------------------------------------------- */
/* path-based open                                                           */
/* ------------------------------------------------------------------------- */

static void test_open_path_succeeds(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    TEST_ASSERT_NOT_NULL_MESSAGE(z, "open_path failed for hidden_row.xlsx");
    lxr_zip_close(z);
}

static void test_open_path_returns_null_on_missing(void)
{
    lxr_zip *z = lxr_zip_open_path("/nonexistent/path/to/file.xlsx");
    TEST_ASSERT_NULL(z);
}

static void test_open_path_returns_null_on_null_arg(void)
{
    TEST_ASSERT_NULL(lxr_zip_open_path(NULL));
}

/* ------------------------------------------------------------------------- */
/* entry existence                                                           */
/* ------------------------------------------------------------------------- */

static void test_entry_exists_known_files(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_TRUE(lxr_zip_entry_exists(z, "[Content_Types].xml"));
    TEST_ASSERT_TRUE(lxr_zip_entry_exists(z, "xl/workbook.xml"));
    TEST_ASSERT_TRUE(lxr_zip_entry_exists(z, "xl/worksheets/sheet1.xml"));
    TEST_ASSERT_FALSE(lxr_zip_entry_exists(z, "xl/worksheets/sheet999.xml"));
    lxr_zip_close(z);
}

/* ------------------------------------------------------------------------- */
/* read entry                                                                */
/* ------------------------------------------------------------------------- */

static void test_read_workbook_xml(void)
{
    char buf[2048];
    ssize_t n;
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_zip_file *zf;

    TEST_ASSERT_NOT_NULL(z);
    zf = lxr_zip_open_entry(z, "xl/workbook.xml");
    TEST_ASSERT_NOT_NULL(zf);

    n = lxr_zip_read(zf, buf, sizeof(buf) - 1);
    TEST_ASSERT_GREATER_THAN(0, n);
    buf[n] = 0;
    TEST_ASSERT_NOT_NULL(strstr(buf, "<workbook"));
    TEST_ASSERT_NOT_NULL(strstr(buf, "<sheet"));

    lxr_zip_close_entry(zf);
    lxr_zip_close(z);
}

static void test_read_drains_to_eof(void)
{
    char buf[256];
    ssize_t n;
    size_t total = 0;
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_zip_file *zf;

    TEST_ASSERT_NOT_NULL(z);
    zf = lxr_zip_open_entry(z, "xl/sharedStrings.xml");
    TEST_ASSERT_NOT_NULL(zf);

    while ((n = lxr_zip_read(zf, buf, sizeof(buf))) > 0) total += (size_t)n;
    TEST_ASSERT_EQUAL_INT(0, n);
    TEST_ASSERT_GREATER_THAN(0, total);

    lxr_zip_close_entry(zf);
    lxr_zip_close(z);
}

static void test_open_missing_entry_returns_null(void)
{
    lxr_zip *z = lxr_zip_open_path(XLSX);
    lxr_zip_file *zf = lxr_zip_open_entry(z, "xl/no_such_thing.xml");
    TEST_ASSERT_NULL(zf);
    lxr_zip_close(z);
}

/* ------------------------------------------------------------------------- */
/* iterate                                                                   */
/* ------------------------------------------------------------------------- */

typedef struct {
    int found_workbook;
    int found_content_types;
    int total;
} iter_state;

static int collect_iter(const char *name, void *ud)
{
    iter_state *s = (iter_state *)ud;
    s->total++;
    if (strcmp(name, "xl/workbook.xml")        == 0) s->found_workbook = 1;
    if (strcmp(name, "[Content_Types].xml")    == 0) s->found_content_types = 1;
    return 0;
}

static void test_iterate_entries(void)
{
    iter_state st = {0, 0, 0};
    lxr_zip *z = lxr_zip_open_path(XLSX);
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_EQUAL_INT(LXR_NO_ERROR, lxr_zip_iterate_entries(z, collect_iter, &st));
    TEST_ASSERT_TRUE(st.found_workbook);
    TEST_ASSERT_TRUE(st.found_content_types);
    TEST_ASSERT_GREATER_THAN(5, st.total);
    lxr_zip_close(z);
}

/* ------------------------------------------------------------------------- */
/* memory and fd sources                                                     */
/* ------------------------------------------------------------------------- */

static void *slurp_file(const char *path, size_t *len)
{
    FILE *fp = fopen(path, "rb");
    void *buf;
    size_t got;
    long fsize;
    if (!fp) return NULL;
    fseek(fp, 0, SEEK_END);
    fsize = ftell(fp);
    fseek(fp, 0, SEEK_SET);
    buf = malloc((size_t)fsize);
    got = fread(buf, 1, (size_t)fsize, fp);
    fclose(fp);
    *len = got;
    return buf;
}

static void test_open_memory(void)
{
    size_t len;
    void *data = slurp_file(XLSX, &len);
    lxr_zip *z;
    TEST_ASSERT_NOT_NULL(data);
    z = lxr_zip_open_memory(data, len);
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_TRUE(lxr_zip_entry_exists(z, "xl/workbook.xml"));
    lxr_zip_close(z);
    free(data);
}

static void test_open_fd(void)
{
    int fd = open(XLSX, O_RDONLY);
    lxr_zip *z;
    TEST_ASSERT_GREATER_OR_EQUAL_INT(0, fd);
    z = lxr_zip_open_fd(fd);
    TEST_ASSERT_NOT_NULL(z);
    TEST_ASSERT_TRUE(lxr_zip_entry_exists(z, "xl/workbook.xml"));
    lxr_zip_close(z);
    close(fd);
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_open_path_succeeds);
    RUN_TEST(test_open_path_returns_null_on_missing);
    RUN_TEST(test_open_path_returns_null_on_null_arg);
    RUN_TEST(test_entry_exists_known_files);
    RUN_TEST(test_read_workbook_xml);
    RUN_TEST(test_read_drains_to_eof);
    RUN_TEST(test_open_missing_entry_returns_null);
    RUN_TEST(test_iterate_entries);
    RUN_TEST(test_open_memory);
    RUN_TEST(test_open_fd);
    return UNITY_END();
}
