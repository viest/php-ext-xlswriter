#!/bin/sh
set -eu

nested_dir="library/libxlsx/"'lxlsx'

if [ -d "$nested_dir" ]; then
    printf '%s\n' "unexpected nested %s directory" "$nested_dir" >&2
    exit 1
fi
