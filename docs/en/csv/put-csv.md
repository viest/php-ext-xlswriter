# Export CSV

## Application Scenario

1. More xlsx file fragments are merged into a single CSV file for unified processing;
2. The xlsx file is added faster than the task processing speed. After converting the file to CSV asynchronously, it can be processed with more efficient tools (for example, the database tool directly imports CSV);

More application scenarios are waiting for you to discover...........

## **Function Prototype**

```php
putCSV(resource $handler): bool
```

### **resource $handler**

> The file pointer must be valid and must point to a file that was successfully opened by fopen() or fsockopen() (and not yet closed by fclose()).

## Example

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial.xlsx', 'TestSheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->output();

// Write mode is on, pointing the file pointer to the end of the file.
$fp = fopen('./tests/file.csv', 'a');

$csvResult = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->putCSV($fp);
```