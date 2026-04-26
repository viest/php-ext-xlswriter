# 工作簿属性

为 xlsx 写入文档属性。Excel 的「文件 → 信息 → 属性」面板会读取这些字段。属性分两类：

* **标准属性** —— 由 OOXML 规范定义的固定字段（标题、作者、公司等），统一通过 `Excel::setProperties` 一次写入。
* **自定义属性** —— 由用户自定义键名 / 类型，通过 `Excel::setCustomProperty` 逐条写入。

## 函数原型

```php
setProperties(array $properties): self
setCustomProperty(string $name, mixed $value, ?string $type = null): self
```

本节内容：

* [标准属性](standard.md)
* [自定义属性](custom.md)
