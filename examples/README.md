# Examples

Runnable, self-contained scripts demonstrating common xlswriter usage.
Each example writes to (and cleans up) a temporary file in
`sys_get_temp_dir()` so you can run them in any order without leaving
stale fixtures behind.

```sh
# After building the extension:
php -d extension=$(pwd)/modules/xlswriter.so examples/01_basic_write.php
```

| Example | What it shows |
|---|---|
| `01_basic_write.php` | Header + tabular data + output() |
| `02_read_full.php` | `getSheetData()` for small files |
| `03_read_streaming.php` | `nextRow()` loop for large files |
| `04_const_memory.php` | Constant-memory write mode for 1M+ rows |
| `05_styles.php` | `Format` chains: colour, bold, alignment, number format |
| `06_merge_cells.php` | `mergeCells()` for header bars |
| `07_csv_export.php` | `putCSV()` & `putCSVCallback()` |
| `08_chart.php` | Building a chart with `Chart` & `insertChart()` |
| `09_images_roundtrip.php` | `insertImage()` + `iterateImages()` round-trip |
| `10_data_validation.php` | Drop-down lists via `Validation` |
| `11_read_styles.php` | Reading back per-cell styles via `getStyleFormat()` |
