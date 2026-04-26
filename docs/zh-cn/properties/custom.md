# 自定义属性

写入用户自定义的工作簿属性。每次调用写入一条键值对。

## 函数原型

```php
setCustomProperty(string $name, mixed $value, ?string $type = null): self
```

### **string $name**

> 属性名。

### **mixed $value**

> 属性值。

### **string $type**

> 属性类型。可选值：
>
> * `string`   —— 字符串
> * `number`   —— 数值（整数或浮点数）
> * `boolean`  —— 布尔
> * `datetime` —— 日期时间，`$value` 须为 **Unix 时间戳**
>
> 省略时根据 `$value` 的 PHP 类型自动推断（string / int / double / bool）。`datetime` 必须显式指定。
>
> 传入非法类型时抛出异常（错误码 `222`）。

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setCustomProperty('Department',   'Sales')          // 自动按 string 处理
    ->setCustomProperty('Confidential', true)             // 自动按 boolean 处理
    ->setCustomProperty('Revision',     3)                // 自动按 number 处理
    ->setCustomProperty('Reviewed',     time(), 'datetime') // datetime 必须显式指定
    ->output();
```
