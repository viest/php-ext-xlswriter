PHP_ARG_WITH(xlsxwriter, xlswriter support,
[  --with-xlswriter           Include xlswriter support], yes)

PHP_ARG_WITH(libxlsxwriter, system libxlsswriter,
[  --with-libxlsxwriter=DIR Use system library], no, no)

if test "$PHP_XLSWRITER" != "no"; then
    xls_writer_sources="
    xls_writer.c \
    kernel/exception.c \
    kernel/resource.c \
    kernel/common.c \
    kernel/excel.c \
    kernel/write.c \
    kernel/format.c \
    "
    libxlsxwriter_sources="
    library/third_party/minizip/ioapi.c \
    library/third_party/minizip/mztools.c \
    library/third_party/minizip/unzip.c \
    library/third_party/minizip/zip.c \
    library/third_party/tmpfileplus/tmpfileplus.c \
    library/src/app.c \
    library/src/chart.c \
    library/src/content_types.c \
    library/src/core.c \
    library/src/custom.c \
    library/src/drawing.c \
    library/src/xlsx_format.c \
    library/src/hash_table.c \
    library/src/packager.c \
    library/src/relationships.c \
    library/src/shared_strings.c \
    library/src/styles.c \
    library/src/theme.c \
    library/src/utility.c \
    library/src/workbook.c \
    library/src/worksheet.c \
    library/src/xmlwriter.c \
    "

    AC_MSG_CHECKING([Check libxlsxwriter library])
    if test "$PHP_LIBXLSXWRITER" != "no"; then

        for i in $PHP_LIBXLSXWRITER /usr/local /usr; do
            if test -r $i/include/xlsxwriter.h; then
                XLSXWRITER_DIR=$i
                AC_MSG_RESULT([found in $i])
                break
            fi
        done
        if test -z "$XLSXWRITER_DIR"; then
            AC_MSG_ERROR([libxlsxwriter library not found])
        else
            PHP_ADD_INCLUDE($XLSXWRITER_DIR/include)
            PHP_CHECK_LIBRARY(xlsxwriter, worksheet_write_string,
            [
                PHP_ADD_LIBRARY_WITH_PATH(xlsxwriter, $i/$PHP_LIBDIR, XLSWRITER_SHARED_LIBADD)
            ],[
                AC_MSG_ERROR([Wrong libxlsxwriter version or library not found])
            ],[
                -L$XLSXWRITER_DIR/$PHP_LIBDIR -lm
            ])
        fi
        AC_DEFINE(HAVE_LIBXLSXWRITER,1,[ ])
    else
        AC_MSG_RESULT([use the bundled library])
        xls_writer_sources="$xls_writer_sources $libxlsxwriter_sources"
        PHP_ADD_INCLUDE([$srcdir/library/include])
    fi

    if test -z "$PHP_DEBUG"; then
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no])
    fi

    PHP_SUBST(XLSWRITER_SHARED_LIBADD)

    PHP_NEW_EXTENSION(xlswriter, $xls_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
