# 数字格式

## **函数原型**

```php
number(string $format): self
```

### **string $format**

> 数组格式字符串

```
"0.000"
"#,##0"
"#,##0.00"
"0.00"
```

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject  = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// 创建样式资源
$format      = new \Vtiful\Kernel\Format($fileHandle);
$numberStyle = $format->number('#,##0')->toResource();

$filePath = $fileObject->header(['name', 'balance'])
    ->data([
        ['viest', 10000],
        ['wjx',   100000]
    ])
    ->setColumn('B:B', 50, $numberStyle) // 应用样式
    ->output();
```

