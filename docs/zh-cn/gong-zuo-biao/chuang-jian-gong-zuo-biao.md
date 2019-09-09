# 创建工作表

### **函数原型**

```php
addSheet([string $sheetName]);
```

### 示例

```php
$config = [
    'path' => './filePath'
];

$excel = new \Vtiful\Kernel\Excel($config);

// 此处会自动创建一个工作表
$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]]);

// 向文件中追加一个工作表
$fileObject->addSheet()
    ->header(['name', 'age'])
    ->data([['wjx', 22]]);

// 最后的最后，输出文件
$filePath = $fileObject->output();
```

#### 

