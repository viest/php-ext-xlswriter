# libxlsxreader

XLSX reading library for php-ext-xlswriter, replacing the libxlsxio dependency.
Designed to mirror the layout of `library/libxlsxwriter` so reading and writing
sit symmetrically inside the project.

See `plans/reader.md` for the full design and phase plan.

## Status

**Phase 0 — skeleton.** Public API surface is declared; all entry points return
`LXR_ERROR_UNSUPPORTED_FEATURE`. Implementation lands in subsequent phases:

| Phase | Module | What lands |
|---|---|---|
| 1 | `zip_io.c`, `xml_pump.c` | minizip + expat foundation |
| 2 | `workbook.c`, `sst.c`, `styles.c` | metadata, shared strings (FULL + STREAMING), styles |
| 3 | `worksheet.c` | streaming row/cell + explicit FSM |
| 4 | (kernel/) | PHP layer switchover, libxlsxio removal |

## Layout

```
include/xlsxreader.h          umbrella header
include/xlsxreader/           public API headers (common, workbook, worksheet, styles)
src/                          implementation + private headers (lxr_internal.h)
test/                         Unity-based unit tests (vendor/Unity is FetchContent'd at configure time)
```

## Build

CMake (recommended):

```sh
cmake -S . -B build -DLXR_BUILD_TESTING=ON
cmake --build build
ctest --test-dir build
```

Plain make (library only, no tests):

```sh
make
```

## Conventions

- C99, no GNU extensions (Windows MSVC must build).
- All public symbols prefixed `lxr_` / `LXR_`.
- No PHP dependencies inside this library — the PHP glue lives in
  `kernel/read.c` of the parent project.
- Errors returned as `lxr_error`; output values via pointer parameters.
- See `plans/reader.md` §4 for full conventions.
