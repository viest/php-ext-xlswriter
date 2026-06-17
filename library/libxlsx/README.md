# libxlsx

`library/libxlsx/` is the single C library root for xlsx functionality in this
repository.

Public C symbols exported by this library use the `lxlsx_*` prefix. The writer
and reader implementations are internal modules of this library, not separate
split libraries.

PHP and Zend extension glue is intentionally kept outside this directory.
