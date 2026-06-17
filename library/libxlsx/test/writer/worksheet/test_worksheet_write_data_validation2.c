/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/lxlsx/worksheet.h"


/* Test: Integer between 1 and 10. */
CTEST(worksheet, test_write_data_validations_201) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Integer not between 1 and 10. */
CTEST(worksheet, test_write_data_validations_202) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"notBetween\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_NOT_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}


/* Test: Integer == 1. */
CTEST(worksheet, test_write_data_validations_203) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Integer != 1. */
CTEST(worksheet, test_write_data_validations_204) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"notEqual\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Integer > 1. */
CTEST(worksheet, test_write_data_validations_205) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Integer < 1. */
CTEST(worksheet, test_write_data_validations_206) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"lessThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_LESS_THAN;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: 1Integer >= 1. */
CTEST(worksheet, test_write_data_validations_207) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"greaterThanOrEqual\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Integer <= 1. */
CTEST(worksheet, test_write_data_validations_208) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"lessThanOrEqual\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO;
    data_validation->value_number = 1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Ignore blank off. */
CTEST(worksheet, test_write_data_validations_209) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->ignore_blank = LXW_VALIDATION_OFF;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Error style == warning. */
CTEST(worksheet, test_write_data_validations_210) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" errorStyle=\"warning\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->error_type = LXW_VALIDATION_ERROR_TYPE_WARNING;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Error style == information */
CTEST(worksheet, test_write_data_validations_211) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" errorStyle=\"information\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->error_type = LXW_VALIDATION_ERROR_TYPE_INFORMATION;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Input title. */
CTEST(worksheet, test_write_data_validations_212) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"Input title January\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Input title + input message. */
CTEST(worksheet, test_write_data_validations_213) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Input title + input message + error title. */
CTEST(worksheet, test_write_data_validations_214) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" errorTitle=\"Error title March\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";
    data_validation->error_title = "Error title March";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Input title + input message + error title + error message. */
CTEST(worksheet, test_write_data_validations_215) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" errorTitle=\"Error title March\" error=\"Error message April\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";
    data_validation->error_title = "Error title March";
    data_validation->error_message = "Error message April";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}
/* Test: Input title. + input message + error title + error message - input message box. */
CTEST(worksheet, test_write_data_validations_216) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showErrorMessage=\"1\" errorTitle=\"Error title March\" error=\"Error message April\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";
    data_validation->error_title = "Error title March";
    data_validation->error_message = "Error message April";
    data_validation->show_input = LXW_VALIDATION_OFF;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Input title + input message + error title + error message - input message box - error message box. */
CTEST(worksheet, test_write_data_validations_217) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" errorTitle=\"Error title March\" error=\"Error message April\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";
    data_validation->error_title = "Error title March";
    data_validation->error_message = "Error message April";
    data_validation->show_input = LXW_VALIDATION_OFF;
    data_validation->show_error = LXW_VALIDATION_OFF;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: 'any' value on its own shouldn't produce a DV record. */
CTEST(worksheet, test_write_data_validations_218) {
    char* got;
    char exp[] = "";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_ANY;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Decimal = 1.2345 */
CTEST(worksheet, test_write_data_validations_219) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"decimal\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>1.2345</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DECIMAL;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_number = 1.2345;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: 28 List = a,bb,ccc */
CTEST(worksheet, test_write_data_validations_220) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>\"a,bb,ccc\"</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    const char* list[] = {"a", "bb", "ccc", NULL};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LIST;
    data_validation->value_list = list;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: List = a,bb,ccc, No dropdown */
CTEST(worksheet, test_write_data_validations_221) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"list\" allowBlank=\"1\" showDropDown=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>\"a,bb,ccc\"</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    const char* list[] = {"a", "bb", "ccc", NULL};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LIST;
    data_validation->value_list = list;
    data_validation->dropdown = LXW_VALIDATION_OFF;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: List = $D$1:$D$5 */
CTEST(worksheet, test_write_data_validations_222) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1\"><formula1>$D$1:$D$5</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LIST_FORMULA;
    data_validation->value_formula = "=$D$1:$D$5";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("A1"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Date = 39653 (2008-07-24) */
CTEST(worksheet, test_write_data_validations_223) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"date\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>39653</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DATE_NUMBER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_number = 39653;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Date = 2008-07-24 */
CTEST(worksheet, test_write_data_validations_224) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"date\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>39653</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_datetime datetime1 = {2008,  7,  24, 0, 0, 0};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DATE;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_datetime = datetime1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Date between ranges. */
CTEST(worksheet, test_write_data_validations_225) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"date\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_datetime datetime1 = {2008,  1,  1, 0, 0, 0};
    lxw_datetime datetime2 = {2008, 12, 12, 0, 0, 0};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DATE;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_datetime = datetime1;
    data_validation->maximum_datetime = datetime2;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Time = 0.5 (12:00:00) */
CTEST(worksheet, test_write_data_validations_226) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"time\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>0.5</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_TIME_NUMBER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_number = 0.5;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5:B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Time = 12:00:00 */
CTEST(worksheet, test_write_data_validations_227) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"time\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>0.5</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_datetime datetime1 = {0,  0,  0, 12, 0, 0};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_TIME;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_datetime = datetime1;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Check data validation range. */
CTEST(worksheet, test_write_data_validations_228) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5:B10\"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 1;
    data_validation->maximum_number = 10;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_range(worksheet, RANGE("B5:B10"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Multiple validations. */
CTEST(worksheet, test_write_data_validations_229) {
    char* got;
    char exp[] = "<dataValidations count=\"2\"><dataValidation type=\"whole\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>10</formula1></dataValidation><dataValidation type=\"whole\" operator=\"lessThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"C10\"><formula1>10</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation->value_number = 10;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);

    data_validation->validate = LXW_VALIDATION_TYPE_INTEGER;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_LESS_THAN;
    data_validation->value_number = 10;
    worksheet_data_validation_cell(worksheet, CELL("C10"), data_validation);

    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: 'any' with an input message should produce a dataValidation record. */
CTEST(worksheet, test_write_data_validations_230) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"Input title January\" prompt=\"Input message February\" sqref=\"B5\"/></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_ANY;
    data_validation->input_title = "Input title January";
    data_validation->input_message = "Input message February";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: length */
CTEST(worksheet, test_write_data_validations_231) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"textLength\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1\"><formula1>5</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LENGTH;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation->value_number = 5;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("A1"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}


/* Test: length */
CTEST(worksheet, test_write_data_validations_232) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"textLength\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1\"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LENGTH;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_number = 5;
    data_validation->maximum_number = 10;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("A1"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: length formula. */
CTEST(worksheet, test_write_data_validations_233) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"textLength\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1\"><formula1>H1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LENGTH_FORMULA;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_GREATER_THAN;
    data_validation->value_formula = "=H1";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("A1"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: length formula. */
CTEST(worksheet, test_write_data_validations_234) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"textLength\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1\"><formula1>H1</formula1><formula2>H2</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_LENGTH_FORMULA;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_formula = "=H1";
    data_validation->maximum_formula = "=H2";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("A1"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Date = 2008-07-24 */
CTEST(worksheet, test_write_data_validations_235) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"date\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>H1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DATE_FORMULA;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_formula = "=H1";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Date between ranges. */
CTEST(worksheet, test_write_data_validations_236) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"date\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>H1</formula1><formula2>H2</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_DATE_FORMULA;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_formula = "=H1";
    data_validation->maximum_formula = "=H2";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Time between ranges. */
CTEST(worksheet, test_write_data_validations_237) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"time\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>0</formula1><formula2>0.5</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);
    lxw_datetime datetime1 = {0,  0,  0,  0, 0, 0};
    lxw_datetime datetime2 = {0,  0,  0, 12, 0, 0};

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_TIME;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_datetime = datetime1;
    data_validation->maximum_datetime = datetime2;

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Time formula */
CTEST(worksheet, test_write_data_validations_238) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"time\" operator=\"equal\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>H1</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_TIME_FORMULA;
    data_validation->criteria = LXW_VALIDATION_CRITERIA_EQUAL_TO;
    data_validation->value_formula = "=H1";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: Time formula */
CTEST(worksheet, test_write_data_validations_239) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"time\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>H1</formula1><formula2>H2</formula2></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_TIME_FORMULA;
        data_validation->criteria = LXW_VALIDATION_CRITERIA_BETWEEN;
    data_validation->minimum_formula = "=H1";
    data_validation->maximum_formula = "=H2";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}

/* Test: length formula. */
CTEST(worksheet, test_write_data_validations_240) {
    char* got;
    char exp[] = "<dataValidations count=\"1\"><dataValidation type=\"custom\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B5\"><formula1>10</formula1></dataValidation></dataValidations>";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_data_validation *data_validation = calloc(1, sizeof(lxw_data_validation));
    data_validation->validate = LXW_VALIDATION_TYPE_CUSTOM_FORMULA;
    data_validation->value_formula = "10";

    lxw_worksheet *worksheet = lxw_worksheet_new(NULL);
    worksheet->file = testfile;

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation);
    _worksheet_write_data_validations(worksheet);

    RUN_XLSX_STREQ(exp, got);

    lxw_worksheet_free(worksheet);
    free(data_validation);
}
