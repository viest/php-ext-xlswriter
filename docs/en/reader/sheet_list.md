# Worksheet list

* The file is not supported for the `windows` system.
* Extended version is greater than or equal to `1.2.7`;
* PECL will prompt you to open the read function when installing, please type `yes`;

## Function Prototype

```php
sheetList(): array
```

## Example

```php
$excel = new \Vtiful\Kernel\Excel(['path' => './tests']);

// Build the sample file
$filePath = $excel
    // First worksheet
    ->fileName('tutorial.xlsx', 'test1')
    ->header(['sheet'])
    ->data([['test1']])

    // Second worksheet
    ->addSheet('test2')
    ->header(['sheet'])
    ->data([['test2']])

    ->output();

// Open the sample file
$sheetList = $excel->openFile('tutorial.xlsx')
    ->sheetList();

Foreach ($sheetList as $sheetName) {
    echo 'Sheet Name:' . $sheetName . PHP_EOL;

    // Get the worksheet data by the worksheet name
    $sheetData = $excel
        ->openSheet($sheetName)
        ->getSheetData();

    var_dump($sheetData);
}
```

## Sample output

```php
Sheet Name: test1
Array(2) {
  [0]=>
  Array(1) {
    [0]=>
    String(5) "sheet"
  }
  [1]=>
  Array(1) {
    [0]=>
    String(5) "test1"
  }
}
Sheet Name: test2
Array(2) {
  [0]=>
  Array(1) {
    [0]=>
    String(5) "sheet"
  }
  [1]=>
  Array(1) {
    [0]=>
    String(5) "test2"
  }
}
```