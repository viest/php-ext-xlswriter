#!/bin/sh
set -eu

status=0

check() {
    pattern=$1
    shift

    if git grep -n -E "$pattern" -- "$@" >/tmp/libxlsx-audit-prefixes.$$; then
        cat /tmp/libxlsx-audit-prefixes.$$
        status=1
    fi

    rm -f /tmp/libxlsx-audit-prefixes.$$
}

check '(^|[^A-Za-z0-9_])lxw_' \
    library/libxlsx/include/lxlsx.h \
    library/libxlsx/include/lxlsx \
    library/libxlsx/src \
    library/libxlsx/test/writer

check '(^|[^A-Za-z0-9_])LXW_' \
    library/libxlsx/include/lxlsx.h \
    library/libxlsx/include/lxlsx \
    library/libxlsx/src \
    library/libxlsx/test/writer

check '#[[:space:]]*include[[:space:]]+[<"]xlsxwriter' \
    library/libxlsx/include/lxlsx.h \
    library/libxlsx/include/lxlsx \
    library/libxlsx/src \
    library/libxlsx/test/writer

check '^[A-Za-z_][A-Za-z0-9_ *]+[[:space:]]+(workbook|worksheet|format|chart|chartsheet)_[A-Za-z0-9_]+[[:space:]]*\(' \
    library/libxlsx/include/lxlsx.h \
    library/libxlsx/include/lxlsx

check '(^|[^A-Za-z0-9_])lxr_' \
    library/libxlsx/include/lxlsx/reader.h \
    library/libxlsx/include/lxlsx/reader \
    library/libxlsx/reader/src \
    library/libxlsx/test/reader

check '(^|[^A-Za-z0-9_])LXR_' \
    library/libxlsx/include/lxlsx/reader.h \
    library/libxlsx/include/lxlsx/reader \
    library/libxlsx/reader/src \
    library/libxlsx/test/reader

check 'libxlsxreader|xlsxreader' \
    library/libxlsx/include/lxlsx/reader.h \
    library/libxlsx/include/lxlsx/reader \
    library/libxlsx/reader/src \
    library/libxlsx/test/reader

check 'lxlsx_reader_cell_type|LXLSX_READER_CELL_|lxlsx_reader_formula_kind|LXLSX_READER_FORMULA_(NORMAL|ARRAY|DATATABLE|SHARED)' \
    library/libxlsx/include/lxlsx/reader.h \
    library/libxlsx/include/lxlsx/reader \
    library/libxlsx/reader/src \
    library/libxlsx/test/reader

exit "$status"
