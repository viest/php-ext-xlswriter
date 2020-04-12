/*****************************************************************************
Copyright (C)  2016  Brecht Sanders  All Rights Reserved

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*****************************************************************************/

/**
 * @file xlsxio_write.h
 * @brief XLSX I/O header file for writing .xlsx files.
 * @author Brecht Sanders
 * @date 2016
 * @copyright MIT
 *
 * Include this header file to use XLSX I/O for writing .xlsx files and
 * link with -lxlsxio_write.
 */

#ifndef INCLUDED_XLSXIO_WRITE_H
#define INCLUDED_XLSXIO_WRITE_H

#include <stdlib.h>
#if defined(_MSC_VER) && _MSC_VER < 1600
typedef signed __int64 int64_t;
#else
#include <stdint.h>
#endif
#include <time.h>

/*! \cond PRIVATE */
#ifndef DLL_EXPORT_XLSXIO
#ifdef _WIN32
#if defined(BUILD_XLSXIO_DLL) || defined(BUILD_XLSXIO_SHARED) || defined(xlsxio_write_SHARED_EXPORTS)
#define DLL_EXPORT_XLSXIO __declspec(dllexport)
#elif !defined(STATIC) && !defined(BUILD_XLSXIO_STATIC) && !defined(BUILD_XLSXIO)
#define DLL_EXPORT_XLSXIO __declspec(dllimport)
#else
#define DLL_EXPORT_XLSXIO
#endif
#else
#define DLL_EXPORT_XLSXIO
#endif
#endif
/*! \endcond */

#ifdef __cplusplus
extern "C" {
#endif

/*! \brief get xlsxio_write version
 * \param  pmajor        pointer to integer that will receive major version number
 * \param  pminor        pointer to integer that will receive minor version number
 * \param  pmicro        pointer to integer that will receive micro version number
 * \sa     xlsxiowrite_get_version_string()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_get_version (int* pmajor, int* pminor, int* pmicro);

/*! \brief get xlsxio_write version string
 * \return version string
 * \sa     xlsxiowrite_get_version()
 */
DLL_EXPORT_XLSXIO const char* xlsxiowrite_get_version_string ();

/*! \brief write handle for .xlsx object */
typedef struct xlsxio_write_struct* xlsxiowriter;

/*! \brief create and open .xlsx file
 * \param  filename      path of .xlsx file to open
 * \param  sheetname     name of worksheet
 * \return write handle for .xlsx object or NULL on error
 * \sa     xlsxiowrite_close()
 */
DLL_EXPORT_XLSXIO xlsxiowriter xlsxiowrite_open (const char* filename, const char* sheetname);

/*! \brief close .xlsx file
 * \param  handle        write handle for .xlsx object
 * \return zero on success, non-zero on error
 * \sa     xlsxiowrite_open()
 */
DLL_EXPORT_XLSXIO int xlsxiowrite_close (xlsxiowriter handle);

/*! \brief specify how many initial rows will be buffered in memory to determine column widths
 * \param  handle        write handle for .xlsx object
 * \param  rows          number of rows to buffer in memory, zero for none
 * Must be called before the first call to xlsxiowrite_next_row()
 * \sa     xlsxiowrite_add_column()
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_set_detection_rows (xlsxiowriter handle, size_t rows);

/*! \brief specify the row height to use for the current and next rows
 * \param  handle        write handle for .xlsx object
 * \param  height        row height (in text lines), zero for unspecified
 * Must be called before the first call to any xlsxiowrite_add_ function of the current row
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_set_row_height (xlsxiowriter handle, size_t height);

/*! \brief add a column cell
 * \param  handle        write handle for .xlsx object
 * \param  name          column name
 * \param  width         column width (in characters)
 * Only one row of column names is supported or none.
 * Call for each column, and finish column row by calling xlsxiowrite_next_row().
 * Must be called before any xlsxiowrite_next_row() or the xlsxiowrite_add_cell_ functions.
 * \sa     xlsxiowrite_next_row()
 * \sa     xlsxiowrite_set_detection_rows()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_column (xlsxiowriter handle, const char* name, int width);

/*! \brief add a cell with string data
 * \param  handle        write handle for .xlsx object
 * \param  value         string value
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_string (xlsxiowriter handle, const char* value);

/*! \brief add a cell with integer data
 * \param  handle        write handle for .xlsx object
 * \param  value         integer value
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_int (xlsxiowriter handle, int64_t value);

/*! \brief add a cell with floating point data
 * \param  handle        write handle for .xlsx object
 * \param  value         floating point value
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_float (xlsxiowriter handle, double value);

/*! \brief add a cell with date and time data
 * \param  handle        write handle for .xlsx object
 * \param  value         date and time value
 * \sa     xlsxiowrite_next_row()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_add_cell_datetime (xlsxiowriter handle, time_t value);

/*! \brief mark the end of a row (next cell will start on a new row)
 * \param  handle        write handle for .xlsx object
 * \sa     xlsxiowrite_add_cell_string()
 */
DLL_EXPORT_XLSXIO void xlsxiowrite_next_row (xlsxiowriter handle);

#ifdef __cplusplus
}
#endif

#endif
