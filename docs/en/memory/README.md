# Memory model

xlswriter offers two memory strategies for writing files:

* **Normal mode** — the entire workbook is held in memory; fastest but proportional to data size.
* **[Fixed memory mode](fixed-memory.md)** — rows are flushed to disk as they are written, keeping peak memory under ~1 MB regardless of file size. Recommended for exporting millions of rows.

Pick fixed memory mode whenever the dataset may not fit comfortably in PHP's `memory_limit`.
