PHP_ARG_WITH(xlswriter, xlswriter support,
[  --with-xlswriter         Include xlswriter support], yes)

PHP_ARG_WITH(libxlsxwriter, system libxlsxwriter,
[  --with-libxlsxwriter=DIR Use system libxlsxwriter], no, no)

PHP_ARG_WITH(libxlsxio, system libxlsxio,
[  --with-libxlsxio=DIR     Use system libxlsxio], no, no)

PHP_ARG_ENABLE(reader, enable xlsx reader support,
[  --enable-reader          Enable xlsx reader?], yes, yes)

if test "$PHP_XLSWRITER" != "no"; then
    xls_writer_sources="
    xlswriter.c \
    kernel/chart.c \
    kernel/common.c \
    kernel/excel.c \
    kernel/exception.c \
    kernel/format.c \
    kernel/help.c \
    kernel/resource.c \
    kernel/rich_string.c \
    kernel/validation.c \
    kernel/write.c \
    "

    xls_read_sources="
    kernel/read.c \
    kernel/csv.c \
    "

    minizip_sources="
    library/libxlsxwriter/third_party/minizip/ioapi.c \
    library/libxlsxwriter/third_party/minizip/mztools.c \
    library/libxlsxwriter/third_party/minizip/unzip.c \
    library/libxlsxwriter/third_party/minizip/zip.c \
    "

    libxlsxwriter_sources="
    library/libxlsxwriter/third_party/tmpfileplus/tmpfileplus.c \
    library/libxlsxwriter/third_party/md5/md5.c \
    library/libxlsxwriter/src/app.c \
    library/libxlsxwriter/src/chart.c \
    library/libxlsxwriter/src/chartsheet.c \
    library/libxlsxwriter/src/comment.c \
    library/libxlsxwriter/src/content_types.c \
    library/libxlsxwriter/src/core.c \
    library/libxlsxwriter/src/custom.c \
    library/libxlsxwriter/src/drawing.c \
    library/libxlsxwriter/src/format.c \
    library/libxlsxwriter/src/hash_table.c \
    library/libxlsxwriter/src/metadata.c \
    library/libxlsxwriter/src/packager.c \
    library/libxlsxwriter/src/relationships.c \
    library/libxlsxwriter/src/shared_strings.c \
    library/libxlsxwriter/src/styles.c \
    library/libxlsxwriter/src/theme.c \
    library/libxlsxwriter/src/utility.c \
    library/libxlsxwriter/src/vml.c \
    library/libxlsxwriter/src/workbook.c \
    library/libxlsxwriter/src/worksheet.c \
    library/libxlsxwriter/src/xmlwriter.c \
    "

    libexpat="
    library/libexpat/expat/lib/loadlibrary.c \
    library/libexpat/expat/lib/xmlparse.c \
    library/libexpat/expat/lib/xmlrole.c \
    library/libexpat/expat/lib/xmltok.c \
    library/libexpat/expat/lib/xmltok_impl.c \
    library/libexpat/expat/lib/xmltok_ns.c \
    "

    libxlsxio="
    library/libxlsxio/lib/xlsxio_read.c \
    library/libxlsxio/lib/xlsxio_read_sharedstrings.c \
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
            PHP_CHECK_LIBRARY(xlsxwriter, lxw_worksheet_find_cell_in_row,
            [
                PHP_ADD_LIBRARY_WITH_PATH(xlsxwriter, $i/$PHP_LIBDIR, XLSWRITER_SHARED_LIBADD)
            ],[
                AC_MSG_ERROR([Wrong libxlsxwriter version or library not found, 0.9.3 required])
            ],[
                -L$XLSXWRITER_DIR/$PHP_LIBDIR -lm
            ])
            PHP_CHECK_LIBRARY(xlsxwriter, lxw_fopen,
            [
                AC_DEFINE(HAVE_LXW_OPEN, 1, [ lxw_fopen available in 0.8.8 ])
            ],[
            ],[
                -L$XLSXWRITER_DIR/$PHP_LIBDIR -lm
            ])
        fi

        AC_DEFINE(HAVE_LIBXLSXWRITER, 1, [ use system libxlsxwriter ])
    else
        AC_MSG_RESULT([use the bundled library])
        xls_writer_sources="$xls_writer_sources $libxlsxwriter_sources"
        PHP_ADD_INCLUDE([$srcdir/library/libxlsxwriter/include])

        dnl see library/CMakeLists.txt
        LIBOPT="-DNOCRYPT -DNOUNCRYPT"
    fi

    if test "$PHP_READER" = "yes"; then
        AC_MSG_CHECKING([Check libxlsxwriter library])
        if test "$PHP_LIBXLSXIO" != "no"; then

            for i in $PHP_LIBXLSXIO /usr/local /usr; do
                if test -r $i/include/xlsxio_read.h; then
                    XLSXIO_DIR=$i
                    AC_MSG_RESULT([found in $i])
                    break
                fi
            done

            if test -z "$XLSXIO_DIR"; then
                AC_MSG_ERROR([libxlsxio library not found])
            else
                PHP_ADD_INCLUDE($XLSXIO_DIR/include)
                PHP_CHECK_LIBRARY(xlsxio_read, xlsxioread_sheet_open,
                [
                    PHP_ADD_LIBRARY_WITH_PATH(xlsxio_read, $i/$PHP_LIBDIR, XLSWRITER_SHARED_LIBADD)
                ],[
                    AC_MSG_ERROR([Wrong libxlsxio version or library not found])
                ],[
                    -L$XLSXWRITER_DIR/$PHP_LIBDIR -lm
                ])
            fi

            AC_DEFINE(HAVE_LIBXLSXIO, 1, [ use system libxlsxwriter ])
        else
            AC_MSG_RESULT([use the bundled library])

            xls_writer_sources="$xls_writer_sources $libexpat"
            PHP_ADD_INCLUDE([$srcdir/library/libexpat/expat/lib])
            PHP_ADD_BUILD_DIR([$abs_builddir/library/libexpat/expat/lib])
            LIBOPT="$LIBOPT -DXML_POOR_ENTROPY"

            xls_writer_sources="$xls_writer_sources $libxlsxio"
            PHP_ADD_INCLUDE([$srcdir/library/libxlsxio/include])
            PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxio/lib])
            LIBOPT="$LIBOPT -DUSE_MINIZIP"

        fi

        xls_writer_sources="$xls_writer_sources $xls_read_sources"
        AC_DEFINE(ENABLE_READER, 1, [enable reader])
    fi

    if test "$PHP_LIBXLSXIO" = "no" || test "$PHP_LIBXLSXWRITER" = "no"; then
        xls_writer_sources="$xls_writer_sources $minizip_sources"
    fi

    if test -z "$PHP_DEBUG"; then
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no])
    fi

    PHP_SUBST(XLSWRITER_SHARED_LIBADD)

    PHP_NEW_EXTENSION(xlswriter, $xls_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1 $LIBOPT)

    PHP_ADD_INCLUDE([$srcdir])
    PHP_ADD_INCLUDE([$srcdir/include])

    PHP_ADD_BUILD_DIR([$abs_builddir/kernel])
    PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxwriter/src])
    PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxwriter/third_party/minizip])
    PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxwriter/third_party/tmpfileplus])
    PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxwriter/third_party/md5])

    PHP_ADD_BUILD_DIR([$abs_builddir/library/libexpat/expat/lib])
    PHP_ADD_BUILD_DIR([$abs_builddir/library/libxlsxio/lib])
fi
