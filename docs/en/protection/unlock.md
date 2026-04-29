# Unlock cells inside a protected sheet

When a worksheet has edit protection enabled, this style lifts protection on specific rows, columns, or individual cells so they remain editable.

## Function Prototype

```php
unlocked();
```

## Example

```php
$config = [
    'path' => './'
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// Build the unlocked style
$format = new \Vtiful\Kernel\Format($fileHandle);
$unlockedStyle = $format->unlocked()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['wjx',   21]
    ])
    ->setRow('A2', 50, $unlockedStyle) // Unlock row 2 of the worksheet
    ->protection() // Protect the rest of the worksheet
    ->output();
```
