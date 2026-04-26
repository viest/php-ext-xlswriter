# Get page setup

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns the active worksheet's page setup — margins, paper, scaling, headers / footers, etc.

## Methods

```php
getPageSetup(): ?array
```

## Return value

Returns `null` when the worksheet declares no page setup at all. Otherwise the fields below, grouped by topic.

**Margins**

* `margins` *(array\|null)* — `{left, right, top, bottom, header, footer}`, in inches.

**Print / scale**

* `paper_size` *(int\|null)* — paper code;
* `fit_to_width` *(int\|null)* — fit-to-width pages;
* `fit_to_height` *(int\|null)* — fit-to-height pages;
* `scale` *(int\|null)* — scale percentage;
* `orientation` *(string\|null)* — `portrait` or `landscape`;
* `horizontal_dpi` *(int\|null)* — horizontal DPI;
* `vertical_dpi` *(int\|null)* — vertical DPI;
* `first_page_number` *(int\|null)* — starting page number;
* `use_first_page_number` *(bool)* — use the custom starting page number.

**Centering / gridlines**

* `print_horizontal_centered` *(bool)* — center horizontally;
* `print_vertical_centered` *(bool)* — center vertically;
* `print_grid_lines` *(bool)* — print gridlines;
* `print_headings` *(bool)* — print row/column headings.

**Header / footer**

* `odd_header` / `odd_footer` *(string\|null)* — odd page header / footer;
* `even_header` / `even_footer` *(string\|null)* — even page header / footer;
* `first_header` / `first_footer` *(string\|null)* — first page header / footer;
* `different_odd_even` *(bool)* — different odd / even pages;
* `different_first` *(bool)* — different first page;
* `scale_with_doc` *(bool)* — scale with document;
* `align_with_margins` *(bool)* — align with page margins.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$ps = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getPageSetup();

var_export($ps);
```
