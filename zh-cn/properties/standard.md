# 标准属性

通过 `Excel::setProperties` 的关联数组写入标准文档属性。所有键均为可选，未设置的字段不会写入。

## 函数原型

```php
setProperties(array $properties): self
```

### **array $properties**

> 支持的键：
>
> * `title`          —— 标题
> * `subject`        —— 主题
> * `author`         —— 作者
> * `manager`        —— 经理
> * `company`        —— 公司
> * `category`       —— 类别
> * `keywords`       —— 关键字
> * `comments`       —— 备注
> * `status`         —— 状态
> * `hyperlink_base` —— 超链接基址
> * `created`        —— 创建时间，**Unix 时间戳**（整数秒），未传则采用文件生成当时的时间

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
    ->setProperties([
        'title'    => 'Q1 销售报表',
        'subject'  => '销售',
        'author'   => 'viest',
        'manager'  => 'wjx',
        'company'  => 'Vtiful',
        'category' => '财务',
        'keywords' => 'sales, q1, 2025',
        'comments' => '由 xlswriter 生成',
        'status'   => 'Draft',
        'created'  => time(),
    ])
    ->output();
```
