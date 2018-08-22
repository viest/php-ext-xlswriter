PHP_ARG_WITH(xlsxwriter, xlswriter support,
[  --with-xlswriter           Include xlswriter support])

PHP_ARG_WITH(zlib-dir,if the location of ZLIB install directory is defined,
[  --with-zlib-dir=<DIR>   Define the location of zlib install directory], no, no)

if test "$PHP_XLSWRITER" != "no" || $PHP_ZLIB_DIR != "no"; then

    for i in /usr/local /usr $PHP_ZLIB_DIR; do
        if test -f $i/include/zlib/zlib.h; then
            ZLIB_DIR=$i
            ZLIB_INCDIR=$i/include/zlib
        elif test -f $i/include/zlib.h; then
            ZLIB_DIR=$i
            ZLIB_INCDIR=$i/include
        fi
    done

    if test -z "$ZLIB_DIR"; then
        AC_MSG_ERROR(Cannot find zlib)
    fi

    AC_MSG_CHECKING([for zlib version >= 1.2.8])

    ZLIB_VERSION=`$EGREP "define ZLIB_VERSION" $ZLIB_INCDIR/zlib.h | $SED -e 's/[[^0-9\.]]//g'`

    if test `echo $ZLIB_VERSION | $SED -e 's/[[^0-9]]/ /g' | $AWK '{print $1*1000000 + $2*10000 + $3*100}'` -lt 1020800; then
        AC_MSG_ERROR([zlib version greater or equal to 1.2.8 required])
    fi

    xls_writer_sources="
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

    xls_writer.c \
    kernel/exception.c \
    kernel/resource.c \
    kernel/common.c \
    kernel/excel.c \
    kernel/write.c \
    kernel/format.c"

    if test -z "$PHP_DEBUG"; then
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no])
    fi

    PHP_ADD_INCLUDE($XLSWRITER_DIR/include)
    PHP_ADD_INCLUDE($ZLIB_INCDIR)

    PHP_NEW_EXTENSION(xlswriter, $xls_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
