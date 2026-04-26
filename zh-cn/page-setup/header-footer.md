# 页眉与页脚

为打印输出设置页眉、页脚文本。文本支持 libxlsxwriter 的格式控制码（如 `&L` 左对齐、`&C` 居中、`&R` 右对齐、`&P` 当前页码、`&N` 总页数、`&D` 当前日期、`&T` 当前时间、`&"Arial,Bold"` 指定字体、`&14` 指定字号等）。

## 函数原型

```php
setHeader(string $value, ?array $options = null): self
setFooter(string $value, ?array $options = null): self
```

### **string $value**

> 页眉 / 页脚文本，支持格式控制码。

### **array $options**

> 可选参数，键名如下：
>
> * `margin`        —— 页眉 / 页脚到纸张边缘的距离（英寸），默认 `0.3`
> * `image_left`    —— 左侧图片路径，配合文本中的 `&G` 占位符使用
> * `image_center`  —— 居中图片路径
> * `image_right`   —— 右侧图片路径

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
    ->setHeader('&L&"Arial,Bold"&14销售报表&R&D')
    ->setFooter('&C第 &P 页 / 共 &N 页', ['margin' => 0.4])
    ->output();
```
