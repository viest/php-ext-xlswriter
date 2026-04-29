# 密码保护

## **函数原型**

```php
protection([string $password]);
```

## 示例一

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection() // 启用工作表保护，无密码保护
    ->output();
```

## 示例二

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection('viest') // 启用工作表保护，密码保护，密码为：viest
    ->output();
```

