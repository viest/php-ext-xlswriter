# Ignore blank lines

##Test data preparation

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

// Write test data
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', '']
    ])
    ->output();
```

## Example 1

```php
// Read the full amount of data
// Use \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW to ignore blank lines

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();
```

## Example 2

```php
// cursor mode
// Use \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS to ignore blank lines

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW);

while ($data = $excel->nextRow()) {
    var_dump($data);
}
```