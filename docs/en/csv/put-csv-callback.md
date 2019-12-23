# Export CSV - Callback Function Mode

## Application scenario

The data needs to be written to CSV after secondary processing.

1. Many xlsx file fragments are merged into a single CSV file for unified processing;
2. The speed of adding xlsx files is greater than the speed of task processing. After asynchronously converting files to CSV, you can use more efficient tools to process them (for example, the database tool directly imports CSV)

More application scenarios are waiting for you to discover ...

## **Function prototype**

```php
putCSVCallback (callable $ callback, resource $ handler): bool
```

### **callable $ callback**

Return function

### **resource $ handler**

> The file pointer must be valid and must point to a file successfully opened by fopen () or fsockopen () (and not yet closed by fclose ()).

## Example

```php
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel ($config);
$filePath = $excel-> fileName('tutorial.xlsx', 'TestSheet1')
    ->header(['String', 'Int', 'Double'])
    ->data([
        ['Item_1', 10, 10.9999995],
    ])
    ->output ();

$fp = fopen ('./ tests / file.csv', 'w');

$csvResult = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->putCSVCallback(function (array $row) {
        return $row;
    }, $fp);
```