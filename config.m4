PHP_ARG_WITH([xlswriter],
    [xlswriter support],
    [AS_HELP_STRING([--without-xlswriter],
        [Disable xlswriter support])],
    [yes])

PHP_ARG_WITH([openssl],
    [openssl MD5],
    [AS_HELP_STRING([[--with-openssl[=DIR]]],
        [Use openssl MD5])],
    [no],
    [no])

if test "$PHP_XLSWRITER" != "no"; then
    xlswriter_sources="
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
    kernel/conditional_format.c \
    kernel/table.c \
    kernel/formula_ast.c \
    kernel/read.c \
    kernel/csv.c \
    kernel/write.c \
    library/libexpat/expat/lib/loadlibrary.c \
    library/libexpat/expat/lib/xmlparse.c \
    library/libexpat/expat/lib/xmlrole.c \
    library/libexpat/expat/lib/xmltok.c \
    library/libxlsx/third_party/tmpfileplus/tmpfileplus.c \
    library/libxlsx/third_party/dtoa/emyg_dtoa.c \
    library/libxlsx/src/app.c \
    library/libxlsx/src/chart.c \
    library/libxlsx/src/chartsheet.c \
    library/libxlsx/src/comment.c \
    library/libxlsx/src/content_types.c \
    library/libxlsx/src/core.c \
    library/libxlsx/src/custom.c \
    library/libxlsx/src/drawing.c \
    library/libxlsx/src/format.c \
    library/libxlsx/src/hash_table.c \
    library/libxlsx/src/metadata.c \
    library/libxlsx/src/packager.c \
    library/libxlsx/src/relationships.c \
    library/libxlsx/src/shared_strings.c \
    library/libxlsx/src/styles.c \
    library/libxlsx/src/table.c \
    library/libxlsx/src/theme.c \
    library/libxlsx/src/utility.c \
    library/libxlsx/src/vml.c \
    library/libxlsx/src/workbook.c \
    library/libxlsx/src/worksheet.c \
    library/libxlsx/src/xmlwriter.c \
    library/libxlsx/src/rich_value.c \
    library/libxlsx/src/rich_value_rel.c \
    library/libxlsx/src/rich_value_structure.c \
    library/libxlsx/src/rich_value_types.c \
    library/libxlsx/src/source_package.c \
    library/libxlsx/src/edit.c \
    library/libxlsx/src/formula.c \
    library/libxlsx/src/common.c \
    library/libxlsx/src/xlsx_util.c \
    library/libxlsx/src/zip_io.c \
    library/libxlsx/src/xml_pump.c \
    library/libxlsx/src/sst.c \
    library/libxlsx/third_party/minizip/ioapi.c \
    library/libxlsx/third_party/minizip/mztools.c \
    library/libxlsx/third_party/minizip/unzip.c \
    library/libxlsx/third_party/minizip/zip.c \
    "

    md5_sources="
    library/libxlsx/third_party/md5/md5.c \
    "

    AC_MSG_CHECKING([Check libxlsx library])
    AC_MSG_RESULT([use the bundled library])

    if test "$PHP_OPENSSL" != "no"; then
        AC_MSG_CHECKING([Check openssl MD5 library])
        AC_MSG_RESULT([use the openssl md5 library])
        for i in $PHP_OPENSSL /usr/local /usr /usr/local/opt; do
            if test -r $i/include/openssl/md5.h; then
                OPENSSL_DIR=$i
                AC_MSG_RESULT([found in $i])
                break
            fi
        done

        if test -z "$OPENSSL_DIR"; then
            PHP_SETUP_OPENSSL(XLSWRITER_SHARED_LIBADD,
            [
                AC_DEFINE(USE_OPENSSL_MD5, 1, [ use openssl md5 ])
            ], [
                AC_MSG_ERROR([openssl library not found])
            ])
        else
            PHP_ADD_INCLUDE($OPENSSL_DIR/include)

            PHP_CHECK_LIBRARY(crypto, MD5_Init,
            [
                PHP_ADD_LIBRARY_WITH_PATH(crypto, $OPENSSL_DIR/lib, XLSWRITER_SHARED_LIBADD)
            ],[
                AC_MSG_ERROR([Wrong openssl MD5_Init not found])
            ],[
                -L$OPENSSL_DIR/lib -lcrypto
            ])

            AC_DEFINE(USE_OPENSSL_MD5, 1, [ use openssl md5 ])
        fi
    else
        AC_MSG_RESULT([use the bundled md5 library])
        xlswriter_sources="$xlswriter_sources $md5_sources"
    fi

    LIBOPT="$LIBOPT -DNOCRYPT -DNOUNCRYPT"

    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR])
    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR/include])
    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR/library/libxlsx/include])
    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR/library/libxlsx/internal])
    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR/library/libxlsx/third_party/minizip])
    PHP_ADD_INCLUDE([PHP_EXT_SRCDIR/library/libexpat/expat/lib])

    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/kernel])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libxlsx/src])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libexpat/expat/lib])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libxlsx/third_party/minizip])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libxlsx/third_party/tmpfileplus])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libxlsx/third_party/dtoa])
    PHP_ADD_BUILD_DIR([PHP_EXT_BUILDDIR/library/libxlsx/third_party/md5])

    LIBOPT="$LIBOPT -DXML_POOR_ENTROPY -DUSE_DTOA_LIBRARY"

    if test -z "$PHP_DEBUG"; then
        AC_ARG_ENABLE(debug, [--enable-debug compile with debugging system], [PHP_DEBUG=$enableval],[PHP_DEBUG=no])
    fi

    PHP_SUBST(XLSWRITER_SHARED_LIBADD)

    PHP_NEW_EXTENSION(xlswriter, $xlswriter_sources, $ext_shared,, -DZEND_ENABLE_STATIC_TSRMLS_CACHE=1 $LIBOPT)
fi
