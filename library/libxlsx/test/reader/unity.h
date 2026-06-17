#ifndef LXLSX_TEST_UNITY_H
#define LXLSX_TEST_UNITY_H

#include <setjmp.h>
#include <stddef.h>
#include <stdio.h>
#include <string.h>

static int lxlsx_test_unity_tests_run;
static int lxlsx_test_unity_failures;
static const char *lxlsx_test_unity_current_test;
static jmp_buf lxlsx_test_unity_jump;

static void lxlsx_test_unity_fail(const char *file, int line,
                                  const char *expr, const char *message)
{
    fprintf(stderr, "%s:%d: %s failed", file, line,
            lxlsx_test_unity_current_test ? lxlsx_test_unity_current_test : "test");
    if (expr && *expr) {
        fprintf(stderr, ": %s", expr);
    }
    if (message && *message) {
        fprintf(stderr, " (%s)", message);
    }
    fputc('\n', stderr);
    lxlsx_test_unity_failures++;
    longjmp(lxlsx_test_unity_jump, 1);
}

static int lxlsx_test_unity_begin(void)
{
    lxlsx_test_unity_tests_run = 0;
    lxlsx_test_unity_failures = 0;
    lxlsx_test_unity_current_test = NULL;
    return 0;
}

static int lxlsx_test_unity_end(void)
{
    if (lxlsx_test_unity_failures == 0) {
        printf("%d tests, 0 failures\n", lxlsx_test_unity_tests_run);
    } else {
        printf("%d tests, %d failures\n",
               lxlsx_test_unity_tests_run, lxlsx_test_unity_failures);
    }
    return lxlsx_test_unity_failures ? 1 : 0;
}

#define UNITY_BEGIN() lxlsx_test_unity_begin()
#define UNITY_END() lxlsx_test_unity_end()

#define RUN_TEST(func) do { \
    lxlsx_test_unity_current_test = #func; \
    lxlsx_test_unity_tests_run++; \
    if (setjmp(lxlsx_test_unity_jump) == 0) { \
        setUp(); \
        func(); \
        tearDown(); \
    } else { \
        tearDown(); \
    } \
} while (0)

#define TEST_FAIL_MESSAGE(message) \
    lxlsx_test_unity_fail(__FILE__, __LINE__, "explicit failure", (message))

#define TEST_ASSERT_TRUE(condition) do { \
    if (!(condition)) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #condition, NULL); \
    } \
} while (0)

#define TEST_ASSERT_FALSE(condition) do { \
    if ((condition)) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, "!(" #condition ")", NULL); \
    } \
} while (0)

#define TEST_ASSERT_NULL(actual) do { \
    if ((actual) != NULL) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual " == NULL", NULL); \
    } \
} while (0)

#define TEST_ASSERT_NOT_NULL(actual) do { \
    if ((actual) == NULL) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual " != NULL", NULL); \
    } \
} while (0)

#define TEST_ASSERT_NOT_NULL_MESSAGE(actual, message) do { \
    if ((actual) == NULL) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual " != NULL", (message)); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_INT(expected, actual) do { \
    long long lxlsx_test_e = (long long)(expected); \
    long long lxlsx_test_a = (long long)(actual); \
    if (lxlsx_test_e != lxlsx_test_a) { \
        char lxlsx_test_msg[128]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "expected %lld, got %lld", lxlsx_test_e, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_INT_MESSAGE(expected, actual, message) do { \
    long long lxlsx_test_e = (long long)(expected); \
    long long lxlsx_test_a = (long long)(actual); \
    if (lxlsx_test_e != lxlsx_test_a) { \
        char lxlsx_test_msg[256]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "%s: expected %lld, got %lld", (message), lxlsx_test_e, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_size_t(expected, actual) do { \
    size_t lxlsx_test_e = (size_t)(expected); \
    size_t lxlsx_test_a = (size_t)(actual); \
    if (lxlsx_test_e != lxlsx_test_a) { \
        char lxlsx_test_msg[128]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "expected %zu, got %zu", lxlsx_test_e, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_GREATER_THAN(threshold, actual) do { \
    long long lxlsx_test_t = (long long)(threshold); \
    long long lxlsx_test_a = (long long)(actual); \
    if (!(lxlsx_test_a > lxlsx_test_t)) { \
        char lxlsx_test_msg[128]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "expected > %lld, got %lld", lxlsx_test_t, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_GREATER_OR_EQUAL_INT(threshold, actual) do { \
    long long lxlsx_test_t = (long long)(threshold); \
    long long lxlsx_test_a = (long long)(actual); \
    if (!(lxlsx_test_a >= lxlsx_test_t)) { \
        char lxlsx_test_msg[128]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "expected >= %lld, got %lld", lxlsx_test_t, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_DOUBLE(expected, actual) do { \
    double lxlsx_test_e = (double)(expected); \
    double lxlsx_test_a = (double)(actual); \
    double lxlsx_test_d = lxlsx_test_e - lxlsx_test_a; \
    if (lxlsx_test_d < 0) lxlsx_test_d = -lxlsx_test_d; \
    if (lxlsx_test_d > 0.000000001) { \
        char lxlsx_test_msg[160]; \
        snprintf(lxlsx_test_msg, sizeof(lxlsx_test_msg), \
                 "expected %.12g, got %.12g", lxlsx_test_e, lxlsx_test_a); \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, lxlsx_test_msg); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_STRING(expected, actual) do { \
    const char *lxlsx_test_e = (expected); \
    const char *lxlsx_test_a = (actual); \
    if ((lxlsx_test_e == NULL) != (lxlsx_test_a == NULL) || \
        (lxlsx_test_e && strcmp(lxlsx_test_e, lxlsx_test_a) != 0)) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, "strings differ"); \
    } \
} while (0)

#define TEST_ASSERT_EQUAL_STRING_LEN(expected, actual, len) do { \
    const char *lxlsx_test_e = (expected); \
    const char *lxlsx_test_a = (actual); \
    size_t lxlsx_test_n = (size_t)(len); \
    if (!lxlsx_test_e || !lxlsx_test_a || strlen(lxlsx_test_e) != lxlsx_test_n || \
        strncmp(lxlsx_test_e, lxlsx_test_a, lxlsx_test_n) != 0) { \
        lxlsx_test_unity_fail(__FILE__, __LINE__, #actual, "string slices differ"); \
    } \
} while (0)

#endif
