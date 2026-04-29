# 工作表背景图片

为当前工作表设置一张铺满整张工作表的背景图。两种方法二选一：从磁盘读取，或直接传入二进制内容（适用于流式 / 内存场景）。

## **函数原型**

```php
setBackgroundImage(string $imagePath): self

setBackgroundImageBuffer(string $imageBuffer): self
```

### **string $imagePath**

> 图片文件路径，支持 PNG / JPEG。

### **string $imageBuffer**

> 图片的二进制内容，例如 `file_get_contents()` 的返回值。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx', 'sheet1')
      ->setBackgroundImage('./assets/logo.png')
      ->output();
```

```php
$excel = new \Vtiful\Kernel\Excel($config);

$buffer = file_get_contents('./assets/logo.png');

$excel->fileName('tutorial.xlsx', 'sheet1')
      ->setBackgroundImageBuffer($buffer)
      ->output();
```
