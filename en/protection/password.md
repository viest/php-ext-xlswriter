# Password protection

## Function Prototype

```php
protection(?string $password = null): self
```

## Example 1

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection() // Enable sheet protection without a password
    ->output();
```

## Example 2

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection('viest') // Enable sheet protection with the password "viest"
    ->output();
```
