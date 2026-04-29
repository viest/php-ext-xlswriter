# 工作表保护

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表 `<sheetProtection>` 元素。文件未启用工作表保护时返回 `null`。

## 函数原型

```php
getSheetProtection(): ?array
```

## 返回值

返回数组，字段按主题分组如下：

**密码**

* `password_hash` *(string\|null)* — `<sheetProtection password="...">` 的哈希值。

**保护项**

* `sheet` *(bool)* — 是否保护工作表本身；
* `content` *(bool)* — 保护工作表内容；
* `objects` *(bool)* — 保护图形对象；
* `scenarios` *(bool)* — 保护方案。

**编辑权限**

* `format_cells` *(bool)* — 允许设置单元格格式；
* `format_columns` *(bool)* — 允许设置列格式；
* `format_rows` *(bool)* — 允许设置行格式；
* `insert_columns` *(bool)* — 允许插入列；
* `insert_rows` *(bool)* — 允许插入行；
* `insert_hyperlinks` *(bool)* — 允许插入超链接；
* `delete_columns` *(bool)* — 允许删除列；
* `delete_rows` *(bool)* — 允许删除行。

**选择 / 其他**

* `select_locked_cells` *(bool)* — 允许选择已锁定单元格；
* `select_unlocked_cells` *(bool)* — 允许选择未锁定单元格；
* `sort` *(bool)* — 允许排序；
* `auto_filter` *(bool)* — 允许使用自动筛选；
* `pivot_tables` *(bool)* — 允许使用数据透视表。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$prot = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getSheetProtection();

var_export($prot);
```
