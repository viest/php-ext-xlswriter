#include <math.h>
#include <stdlib.h>
#include <string.h>

#include <unity.h>

#include "libxlsx/formula.h"

void setUp(void) {}
void tearDown(void) {}

/* Mock grid: A1=10, A2=20, A3=30, B1="hello", C1=TRUE, rest blank. */
static void resolver(void *ctx, lxlsx_row_t row, lxlsx_col_t col, lxlsx_value *out)
{
    (void) ctx;
    out->kind = LXLSX_VAL_BLANK; out->number = 0; out->string = NULL; out->error = LXLSX_FERR_NONE;
    if (col == 0 && row == 0) { out->kind = LXLSX_VAL_NUMBER; out->number = 10; }
    else if (col == 0 && row == 1) { out->kind = LXLSX_VAL_NUMBER; out->number = 20; }
    else if (col == 0 && row == 2) { out->kind = LXLSX_VAL_NUMBER; out->number = 30; }
    else if (col == 1 && row == 0) { out->kind = LXLSX_VAL_STRING; out->string = strdup("hello"); }
    else if (col == 2 && row == 0) { out->kind = LXLSX_VAL_BOOL; out->number = 1; }
}

static lxlsx_value evalf(const char *f)
{
    lxlsx_value v;
    TEST_ASSERT_EQUAL_INT(LXLSX_NO_ERROR, lxlsx_formula_eval(f, resolver, NULL, &v));
    return v;
}

static void assert_num(const char *f, double expect)
{
    lxlsx_value v = evalf(f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(LXLSX_VAL_NUMBER, v.kind, f);
    TEST_ASSERT_TRUE(fabs(expect - v.number) < 1e-9);
    lxlsx_value_free(&v);
}
static void assert_bool(const char *f, int expect)
{
    lxlsx_value v = evalf(f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(LXLSX_VAL_BOOL, v.kind, f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(expect, v.number != 0.0, f);
    lxlsx_value_free(&v);
}
static void assert_str(const char *f, const char *expect)
{
    lxlsx_value v = evalf(f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(LXLSX_VAL_STRING, v.kind, f);
    TEST_ASSERT_EQUAL_STRING(expect, v.string);
    lxlsx_value_free(&v);
}
static void assert_err(const char *f, lxlsx_formula_error expect)
{
    lxlsx_value v = evalf(f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(LXLSX_VAL_ERROR, v.kind, f);
    TEST_ASSERT_EQUAL_INT_MESSAGE(expect, v.error, f);
    lxlsx_value_free(&v);
}

static void test_arithmetic(void)
{
    assert_num("1+2*3", 7);
    assert_num("(1+2)*3", 9);
    assert_num("2^10", 1024);
    assert_num("7/2", 3.5);
    assert_num("-3+1", -2);
    assert_num("10%", 0.1);
    assert_num("2+3*4-5", 9);
}

static void test_references(void)
{
    assert_num("A1", 10);
    assert_num("A1+A2", 30);
    assert_num("A1+A2+A3", 60);
    assert_num("$A$1*2", 20);
}

static void test_aggregates(void)
{
    assert_num("SUM(A1:A3)", 60);
    assert_num("SUM(A1:A3, 100)", 160);
    assert_num("AVERAGE(A1:A3)", 20);
    assert_num("MIN(A1:A3)", 10);
    assert_num("MAX(A1:A3)", 30);
    assert_num("COUNT(A1:C1)", 1);    /* only A1 is a number */
    assert_num("COUNTA(A1:C1)", 3);   /* A1, B1, C1 non-blank */
}

static void test_logical(void)
{
    assert_num("IF(A1>5, 100, 200)", 100);
    assert_num("IF(A1>50, 100, 200)", 200);
    assert_bool("AND(A1>5, A2>5)", 1);
    assert_bool("OR(A1>50, A2>50)", 0);
    assert_bool("NOT(A1>50)", 1);
    assert_num("IFERROR(10/0, -1)", -1);
}

static void test_comparisons(void)
{
    assert_bool("3<>4", 1);
    assert_bool("3=3", 1);
    assert_bool("A1<=10", 1);
    assert_bool("A1>A2", 0);
}

static void test_text(void)
{
    assert_str("\"foo\"&\"bar\"", "foobar");
    assert_str("CONCATENATE(B1, \" world\")", "hello world");
    assert_num("LEN(B1)", 5);
    assert_str("UPPER(B1)", "HELLO");
    assert_str("LEFT(B1, 3)", "hel");
    assert_str("RIGHT(B1, 2)", "lo");
    assert_str("MID(B1, 2, 3)", "ell");
    assert_str("TRIM(\"  a   b  \")", "a b");
}

static void test_math_funcs(void)
{
    assert_num("ABS(-5)", 5);
    assert_num("INT(3.9)", 3);
    assert_num("ROUND(3.14159, 2)", 3.14);
    assert_num("ROUNDUP(3.111, 1)", 3.2);
    assert_num("MOD(7, 3)", 1);
    assert_num("POWER(2, 8)", 256);
    assert_num("SQRT(144)", 12);
}

static void test_errors(void)
{
    assert_err("10/0", LXLSX_FERR_DIV0);
    assert_err("SQRT(-1)", LXLSX_FERR_NUM);
    assert_err("FOOBAR(1)", LXLSX_FERR_NAME);   /* unknown function */
    assert_err("NoSuchName", LXLSX_FERR_NAME);  /* unknown name */
    assert_err("SUM(A1:A3) + 1/0", LXLSX_FERR_DIV0); /* error propagation */
}

int main(void)
{
    UNITY_BEGIN();
    RUN_TEST(test_arithmetic);
    RUN_TEST(test_references);
    RUN_TEST(test_aggregates);
    RUN_TEST(test_logical);
    RUN_TEST(test_comparisons);
    RUN_TEST(test_text);
    RUN_TEST(test_math_funcs);
    RUN_TEST(test_errors);
    return UNITY_END();
}
