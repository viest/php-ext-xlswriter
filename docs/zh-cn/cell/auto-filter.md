# 自动过滤

## **函数原型**

```php
autoFilter(string $range): self
```

### **string $range**

> 筛选/过滤 数据范围

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("tutorial.xlsx")
    ->header(['name', 'age'])
    ->data([
        ['one', 10],
        ['two', 20],
        ['three', 30],
    ])
    ->autoFilter("A1:B3") // 添加 过滤/筛选
    ->output();
```