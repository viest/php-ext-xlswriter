# Read file

The reader opens an existing XLSX and exposes its data either in full (load all rows into an array) or via a cursor (one row at a time, constant memory). It also surfaces sheet metadata — styles, merged cells, hyperlinks, validations, conditional formats, page setup, defined names, images, comments and charts.

Reader features require the extension to be built with `--enable-reader` (the default in PECL packages).

Pages in this section:

* Iteration: [worksheet list](sheet_list.md), [worksheet list with metadata](sheet-list-with-meta.md), [full read](read-file-full.md), [cursor read](read-file-cursor.md), [cell callback mode](cell_callback.md)
* Filtering: [skip empty cells](skip-empty-cells.md), [skip empty rows](skip-empty-row.md), [set global type](set-type.md), [read by data type](data-type.md), [data type constants](data-type-const.md)
* Cell-level: [next row with formula](next-row-with-formula.md), [next row with rich text](next-row-rich.md), [read style format](get-style-format.md)
* Sheet-level metadata: [merged cells](get-merged-cells.md), [hyperlinks](get-hyperlinks.md), [sheet protection](get-sheet-protection.md), [row / column options](row-column-options.md), [conditional formats](get-conditional-formats.md), [data validations](get-data-validations.md), [auto filter](get-auto-filter.md), [defined names](get-defined-names.md), [page setup](get-page-setup.md), [formula AST](get-formula-ast.md)
* Embedded objects: [iterate images](iterate-images.md), [iterate comments](iterate-comments.md), [iterate charts](iterate-charts.md)
