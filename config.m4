PHP_ARG_ENABLE(excel_writer, whether to enable excel_writer support,
[  --enable-excel_writer           Enable excel_writer support])

if test "$PHP_EXCEL_WRITER" != "no"; then
    excel_writer_sources="excel_writer.c \
    kernel/exception.c \
    kernel/common/resource.c \
    "

    AC_MSG_CHECKING([Check libxlsxwriter support])

    for i in /usr/local /usr; do
        if test -r $i/include/xlsxwriter.h; then
            AC_MSG_CHECKING([Check libxlsxwriter library])
            XLSXWRITER_DIR=$i
            PHP_ADD_INCLUDE($i/include)
            PHP_CHECK_LIBRARY(xlsxwriter, worksheet_write_string,
            [
                PHP_ADD_LIBRARY_WITH_PATH(xlsxwriter, $i/$PHP_LIBDIR, EXCEL_WRITER_SHARED_LIBADD)
                excel_writer_sources="$excel_writer_sources \
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

    PHP_SUBST(EXCEL_WRITER_SHARED_LIBADD)

    PHP_NEW_EXTENSION(excel_writer, $excel_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
