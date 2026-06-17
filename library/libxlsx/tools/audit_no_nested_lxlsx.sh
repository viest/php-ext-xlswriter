#!/bin/sh
set -eu

if [ -d "library/libxlsx/lxlsx" ]; then
    printf '%s\n' "unexpected nested library/libxlsx/lxlsx directory" >&2
    exit 1
fi
