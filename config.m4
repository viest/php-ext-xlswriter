PHP_ARG_WITH(xlsxwriter, xlswriter support,
[  --with-xlswriter           Include xlswriter support])

if test "$PHP_XLSWRITER" != "no"; then
    xls_writer_sources="xls_writer.c \
    kernel/exception.c \
    kernel/resource.c \
    kernel/common.c \
    "

    AC_MSG_CHECKING([Check libxlsxwriter support])

    for i in /usr/local /usr; do
        if test -r $i/include/xlsxwriter.h; then
            AC_MSG_CHECKING([Check libxlsxwriter library])
            XLSXWRITER_DIR=$i
            PHP_ADD_INCLUDE($i/include)
            PHP_CHECK_LIBRARY(xlsxwriter, worksheet_write_string,
            [
                PHP_ADD_LIBRARY_WITH_PATH(xlsxwriter, $i/$PHP_LIBDIR, XLSWRITER_SHARED_LIBADD)
                xls_writer_sources="$xls_writer_sources \
                kernel/excel.c \
                kernel/write.c \
                kernel/format.c"
            ],[
                AC_MSG_ERROR([Wrong libxlsxwriter version or library not found])
            ],[
                -L$i/$PHP_LIBDIR -lm
            ])

            break
        else
            AC_MSG_RESULT([no, found in $i])
        fi
    done

    if test -z "$XLSXWRITER_DIR"; then
        AC_MSG_ERROR([libxlsxwriter library not found])
    fi

    if test -z "$PHP_DEBUG"; then
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no])
    fi

    PHP_SUBST(XLSWRITER_SHARED_LIBADD)

    PHP_NEW_EXTENSION(xlswriter, $xls_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
