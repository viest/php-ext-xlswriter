/*
 * Tests for the lib_xlsx_writer library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2026, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/libxlsx/worksheet.h"


// Function used for testing.
uint32_t
width_to_pixels(double width)
{
    double max_digit_width = 7.0;
    double padding = 5.0;
    double pixels;

    if (width < 1.0)
        pixels = (uint32_t) (width * (max_digit_width + padding) + 0.5);
    else
        pixels = (uint32_t) (width * max_digit_width + 0.5) + 5;

    return (uint32_t)pixels;
}

// Function used for testing.
uint32_t
height_to_pixels(double height)
{
    return (uint32_t) (height / 0.75);
}


// Test the Worksheet _pixels_to_width() function.
CTEST(worksheet, pixel_to_width01) {

    int pixels;
    double got;
    double exp;

    for (pixels = 0; pixels <= 1790; pixels++) {
        exp = pixels;
        got = width_to_pixels(_pixels_to_width(pixels));
        ASSERT_DBL_NEAR(exp, got);
    }
}

// Test the Worksheet _pixels_to_height() function.
CTEST(worksheet, pixel_to_height01) {

    int pixels;
    double got;
    double exp;

    for (pixels = 0; pixels <= 545; pixels++) {
        exp = pixels;
        got = height_to_pixels(_pixels_to_height(pixels));
        ASSERT_DBL_NEAR(exp, got);
    }
}
