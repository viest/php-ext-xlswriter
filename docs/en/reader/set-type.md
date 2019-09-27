# Global read type

## Function Prototype

```php
setType(array $type)
```

## Test data preparation

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Name', 'Age', 'Date'])
    ->data([
        ['Viest', 24]
    ])
    ->insertDate(1, 2, 1568877706)
    ->output();
```

## Example

```php
$data = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_INT,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ])
    ->getSheetData();

var_dump($data);
```

## Sample output

```php
array(2) {
   [0]=>
   Array(3) {
     [0]=>
     String(4) "Name"
     [1]=>
     String(3) "Age"
     [2]=>
     String(4) "Date"
   }
   [1]=>
   Array(3) {
     [0]=>
     String(5) "Viest"
     [1]=>
     Int(24)
     [2]=>
     Int(1568877706)
   }
}
```