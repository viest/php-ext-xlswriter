PHP_ARG_WITH(xlsxwriter, xlswriter support,
[  --with-xlswriter           Include xlswriter support])

if test "$PHP_XLSWRITER" != "no"; then
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
    library/src/format.c \
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

    PHP_NEW_EXTENSION(xlswriter, $xls_writer_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1)

    PHP_ADD_BUILD_DIR([$ext_builddir/kernel])
fi
