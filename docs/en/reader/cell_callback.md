# Cell callback mode read

> Maximum Memory = Maximum Cell Data Volume
>
> This mode can satisfy `xlsx` large file reading

## Function Prototype

```php
nextCellCallback(callable $callback, string $sheetName = NULL): void
```

## Callback function prototype

```php
function(int $row, int $cell, string|double|int $data)
```

## Single line end identifier

At the end of each line, `callback call` will be appended and pass `XLSX_ROW_END` to the end of the current line.

##example

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('tutorial.xlsx')->nextCellCallback(function ($row, $cell, $data) {
    echo 'cell:' . $cell . ', row:' . $row . ', value:' . $data . PHP_EOL;
});
```

## Sample output

```php
cell:0, row:0, value:Item
cell:1, row:0, value:Cost
cell:1, row:0, value:XLSX_ROW_END // end identifier
cell:0, row:1, value:Item_1
cell:1, row:1, value:Cost_1
cell:1, row:1, value:XLSX_ROW_END // end identifier
```