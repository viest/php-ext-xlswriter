# 数据验证

数据验证用于约束单元格输入，例如限定为整数、小数、日期、列表选项等。当用户输入不符合约束的值时，Excel 会在打开文件时给出警告。需要注意的是，扩展在 **写入** 时不会做校验，只是把规则随文件保存下来。

整体流程分两步：

1. 通过 `Vtiful\Kernel\Validation` 构造一条规则；
2. 调用 `Excel::validation(string $range, $validationHandle)` 把规则应用到指定区域。其中 `$validationHandle` 是 `Validation::toResource()` 返回的资源句柄。

## 函数原型

```
\Vtiful\Kernel\Excel::validation(string $range, resource $validationHandle): self
```

### **string $range**

> 应用规则的单元格区域，使用 A1 表示法，例如 `A1`、`A1:A10`、`B2:D8`。

### **resource $validationHandle**

> `Validation::toResource()` 返回的资源句柄。

## 子章节

* [下拉列表](type_list.md)
* [范围约束](criteria_between.md)
* [大于约束](criteria_greater_than.md)
* [Validation 类完整 API](api-reference.md)

## 示例

下面是一个最常见的整数范围约束示例：单元格 `A1:A10` 必须填入 1 到 100 之间的整数。

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_BETWEEN)
    ->minimumNumber(1)
    ->maximumNumber(100);

$excel->fileName('tutorial.xlsx')
    ->header(['Score'])
    ->validation('A1:A10', $validation->toResource())
    ->output();
```
