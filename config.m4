PHP_ARG_ENABLE(vtiful, whether to enable vtiful support,
[  --enable-vtiful           Enable vtiful support])

if test "$PHP_VTIFUL" != "no"; then
    vtiful_sources="vtiful.c \
    kernel/exception.c \
    "

    AC_MSG_CHECKING([Check libxlsxwriter support])

    for i in /usr/local /usr; do
        if test -r $i/include/xlsxwriter.h; then
            AC_MSG_CHECKING([Check libxlsxwriter library])
            XLSXWRITER_DIR=$i
            PHP_ADD_INCLUDE($i/include)
            PHP_CHECK_LIBRARY(xlsxwriter, worksheet_write_string,
            [
                PHP_ADD_LIBRARY_WITH_PATH(xlsxwriter, $i/$PHP_LIBDIR, VTIFUL_SHARED_LIBADD)
                AC_DEFINE([VTIFUL_XLSX_WRITER], [1], [Have libxlsxwriter support])
                vtiful_sources="$vtiful_sources kernel/excel.c kernel/write.c"
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
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no] )
    fi

    PHP_SUBST(VTIFUL_SHARED_LIBADD)

    PHP_NEW_EXTENSION(vtiful, $vtiful_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
