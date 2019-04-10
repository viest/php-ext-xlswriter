## xlswriter [ Excel-Export ]

Authon: viest [dev@service.viest.me](mailto:dev@service.viest.me)

## 安装

- [编译安装](#编译安装)
- [PECL](#PECL)
  
## 使用

1. [创建一个简单的xlsx文件](#创建一个简单的xlsx文件)
2. 图表
   1. [图表添加数据](#图表添加数据)
   2. [直方图](#直方图)
   3. [面积图](#面积图)
3. [单元格插入文字](#单元格插入文字)
4. [单元格插入链接](#单元格插入链接)
5. [单元格插入公式](#单元格插入公式)
6. [单元格插入本地图片](#单元格插入本地图片)
7. [数据过滤](#数据过滤)
8. [合并单元格](#合并单元格)
9. [设置列单元格样式](#设置列单元格格式)
10. [设置行单元格样式](#设置行单元格格式)
11. [设置文字颜色](#设置文字颜色)
12. [固定内存导出](#固定内存导出)
13. [创建工作表](#创建工作表)
14. [组合样式](#组合样式)
15. [样式列表](#样式列表)
16. [颜色常量](#颜色常量)

## PECL

```bash
pecl install xlswriter

# 添加 extension = xlswriter.so 到 ini 配置
```

## 编译安装

### Unix

##### Ubuntu

```bash
# 依赖

sudo apt-get install -y zlib1g-dev

# 扩展

git clone https://github.com/viest/php-ext-excel-export.git

cd php-ext-excel-export

git submodule update --init

phpize && ./configure --with-php-config=/path/to/php-config

make && make install

# 添加 extension = xlswriter.so 到 ini 配置
```

##### Mac

```bash
# 依赖

brew install zlib

# 扩展

git clone https://github.com/viest/php-ext-excel-export.git

cd php-ext-excel-export

git submodule update --init

phpize && ./configure --with-php-config=/path/to/php-config

make && make install

# 添加 extension = xlswriter.so 到 ini 配置
```

#### Windows

##### 依赖

> 请预先搭建PHP编译环境，教程详见 [php.net](https://wiki.php.net/internals/windows/stepbystepbuild)

```bash
cd PHP_BUILD_PATH/deps

DownloadFile http://zlib.net/zlib-1.2.11.tar.gz
7z x zlib-1.2.11.tar.gz > NUL
7z x zlib-1.2.11.tar > NUL
cd zlib-1.2.11
cmake -G "Visual Studio 14 2015" -DCMAKE_BUILD_TYPE="Release" -DCMAKE_C_FLAGS_RELEASE="/MT"
cmake --build . --config "Release"
```

##### 扩展

```bash
cd PHP_PATH/ext

git clone https://github.com/viest/php-ext-excel-export.git

cd EXT_PATH

git submodule update --init

phpize

configure.bat --with-xlswriter --with-extra-libs=PATH\zlib-1.2.11\Release --with-extra-includes=PATH\zlib-1.2.11

nmake
```

### 创建一个简单的xlsx文件

```php
$config = ['path' => '/home/viest'];

$excel = new \Vtiful\Kernel\Excel($config);

// fileName 会自动创建一个工作表，你可以自定义该工作表名称，工作表名称为可选参数
$filePath = $excel->fileName('tutorial01.xlsx', 'sheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Rent', 1000],
        ['Gas',  100],
        ['Food', 300],
        ['Gym',  50],
    ])
    ->output();
```

### 图表添加数据

##### 函数原型

```php
series(string $value,[ string $categories])
```

##### string $value

> 图表单个类别数据所在的工作表及单元格跨度
>
> `Sheet1 !   $A$1   : $A$5`
>
> `工作表  ! 开始单元格 : 结束单元格`

##### string $categories

> 类别名称

### 直方图

##### 类型

```php
\Vtiful\Kernel\Chart::CHART_COLUMN
```

##### 图表

![php-excel](https://github.com/viest/php-ext-excel-export/blob/master/resource/chart_simple.png)

##### 实例

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_COLUMN);

$chartResource = $chart->series('Sheet1!$A$1:$A$5')
    ->series('Sheet1!$B$1:$B$5')
    ->series('Sheet1!$C$1:$C$5')
    ->toResource();

$filePath = $fileObject->data([
    [1, 2, 3],
    [2, 4, 6],
    [3, 6, 9],
    [4, 8, 12],
    [5, 10, 15],
])->insertChart(0, 3, $chartResource)->output();
```

### 面积图

##### 类型

```php
\Vtiful\Kernel\Chart::CHART_AREA
```

##### 图标

![php-excel](https://github.com/viest/php-ext-excel-export/blob/master/resource/chart_area1.png)

##### 实例

```php
<?php

$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_AREA);

$chartResource = $chart
    ->series('=Sheet1!$B$2:$B$7', '=Sheet1!$A$2:$A$7')
    ->seriesName('=Sheet1!$B$1')
    ->series('=Sheet1!$C$2:$C$7', '=Sheet1!$A$2:$A$7')
    ->seriesName('=Sheet1!$C$1')
    ->style(11)// 值为 1 - 48，可参考 Excel 2007 "设计" 选项卡中的 48 种样式
    ->axisNameX('Test number') // 设置 X 轴名称
    ->axisNameY('Sample length (mm)') // 设置 Y 轴名称
    ->title('Results of sample analysis') // 设置图表 Title
    ->toResource();

$filePath = $fileObject->header(['Number', 'Batch 1', 'Batch 2'])
    ->data([
        [2, 40, 30],
        [3, 40, 25],
        [4, 50, 30],
        [5, 30, 10],
        [6, 25, 5],
        [7, 50, 10],
    ])->insertChart(0, 3, $chartResource)->output();
```

### 单元格插入文字

#### 函数原型

```php
insertText(int $row, int $column, string|int|double $data[, string $format])
```

##### int $row

> 单元格所在行

##### int $column

> 单元格所在列

##### string | int | double $data

> 需要写入的内容

##### string $format

> 内容格式

##### 实例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$textFile = $excel->fileName("free.xlsx")
    ->header(['name', 'money']);

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index+1, 0, 'viest');
    $textFile->insertText($index+1, 1, 10000, '#,##0');
}

$textFile->output();
```

### 单元格插入链接

#### 函数原型

```php
insertUrl(int $row, int $column, string $url[, resource $format])
```

##### int $row

> 单元格所在行

##### int $column

> 单元格所在列

##### string $url

> 链接地址

##### resource $format

> 链接样式

##### 实例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$urlFile = $excel->fileName("free.xlsx")
    ->header(['url']);

$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$urlStyle = $format->bold()
    ->underline(Format::UNDERLINE_SINGLE)
    ->toResource();

$urlFile->insertUrl(1, 0, 'https://github.com', $urlStyle);

$textFile->output();
```

### 单元格插入公式

#### 函数原型

```php
insertFormula(int $row, int $column, string $formula)
```

##### int $row

> 单元格所在行

##### int $column

> 单元格所在列

##### string $formula

> 公式

##### 实例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx")
    ->header(['name', 'money']);

for($index = 1; $index < 10; $index++) {
    $textFile->insertText($index, 0, 'viest');
    $textFile->insertText($index, 1, 10);
}

$textFile->insertText(12, 0, "Total");
$textFile->insertFormula(12, 1, '=SUM(B2:B11)');

$freeFile->output();
```

### 单元格插入本地图片

#### 函数原型

```php
insertImage(int $row, int $column, string $localImagePath[, double $widthScale, double $heightScale])
```

##### int $row

> 单元格所在行

##### int $column

> 单元格所在列

##### string $localImagePath

> 图片路径

##### double $widthScale

> 对图像X轴进行缩放处理；
> 默认为1，保持图像原始宽度；值为0.5时，图像宽度为原图的1/2；

##### double $heightScale

> 对图像轴进行缩放处理；
> 默认为1，保持图像原始高度；值为0.5时，图像高度为原图的1/2；

##### 实例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx");

$freeFile->insertImage(5, 0, '/vagrant/ASW-G-66.jpg');

$freeFile->output();
```

### 数据过滤

#### 函数原型

```php
autoFilter(string $scope);
```

##### string $scope

> 过滤范围

##### 实例

```php
$excel->fileName('test.xlsx')
        ->header(['name', 'age'])
        ->data($data)
        ->autoFilter('A1:B11')
        ->output();
```

### 合并单元格

#### 函数原型

```php
mergeCells(string $scope, string $data);
```

##### string $scope

> 单元格范围

##### string $data

> 数据

##### 实例

```php
$excel->fileName("test.xlsx")
        ->mergeCells('A1:C1', 'Merge cells')
        ->output();
```

### 设置列单元格格式

#### 函数原型

```php
setColumn(string $range, double $width [, resource $format]);
```

##### string $range

> 单元格范围

##### double $width

> 单元格宽度

##### string $format

> 单元格样式

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setColumn('A:A', 200, $boldStyle)
    ->output();
```

### 设置行单元格格式

#### 函数原型

```php
setRow(string $range, double $height [, resource $format]);
```

##### string $range

> 单元格范围

##### double $height

> 单元格高度

##### string $format

> 单元格样式

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setRow('A1', 20, $boldStyle,)
    ->output();
```

### 设置文字颜色

#### 函数原型

```php
color(int $color)
```

##### int $color

> RGB 十六进制值

##### 实例

```php
$config     = ['path' => './tests'];
$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->color(0xFF0000)->toResource();
//或 $colorStyle = $format->color(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setRow('A1', 50, $colorStyle)
    ->output();

var_dump($filePath);
```

### 固定内存导出

#### 内存

最大内存使用量 = 最大一行的数据占用量

#### 函数原型

```php
constMemory(string $fileName);
```

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->constMemory('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setRow($boldStyle, 'A1')
    ->output();
```

### 创建工作表

#### 函数原型

```php
addSheet([string $sheetName]);
```

##### 实例

```php
$config = [
    'path' => './filePath'
];

$excel = new \Vtiful\Kernel\Excel($config);

// 此处会自动创建一个表格
$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]]);

// 向文件中追加一个表格
$fileObject->addSheet()
    ->header(['name', 'age'])
    ->data([['wjx', 22]]);

// 最后的最后，输出文件
$filePath = $fileObject->output();
```

### 组合样式

将多个样式合并为一个新样式应用在单元格上

```php
// 将粗体与斜体合并为一个样式
$format          = new \Vtiful\Kernel\Format($fileHandle);
$boldItalicStyle = $format->bold()->italic()->toResource();
```

### 样式列表

##### 粗体

```php
$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();
```
##### 斜体

```php
$format      = new \Vtiful\Kernel\Format($fileHandle);
$italicStyle = $format->italic()->toResource();
```
##### 下划线

###### 函数原型

```php
underline(resource $resourchHandle, Format::const $style): \Vtiful\Kernel\Format
```

###### 实例

```php
$format         = new \Vtiful\Kernel\Format($fileHandle);
$underlineStyle = $format->underline(Format::UNDERLINE_SINGLE)->toResource();
```
###### Style 

```php
Format::UNDERLINE_SINGLE;            // 单下划线
Format::UNDERLINE_DOUBLE;            // 双下划线
Format::UNDERLINE_SINGLE_ACCOUNTING; // 会计用单下划线
Format::UNDERLINE_DOUBLE_ACCOUNTING; // 会计用双下划线
```

##### 对齐

###### 函数原型

```php
align(resource $resourchHandle, Format::const ...$style): \Vtiful\Kernel\Format
```

###### 实例

```php
$format     = new \Vtiful\Kernel\Format($fileHandle);
$alignStyle = $format
    ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
    ->toResource();
```
###### Style 

```php
Format::FORMAT_ALIGN_LEFT;                 // 水平左对齐
Format::FORMAT_ALIGN_CENTER;               // 水平剧中对齐
Format::FORMAT_ALIGN_RIGHT;                // 水平右对齐
Format::FORMAT_ALIGN_FILL;                 // 水平填充对齐
Format::FORMAT_ALIGN_JUSTIFY;              // 水平两端对齐
Format::FORMAT_ALIGN_CENTER_ACROSS;        // 横向中心对齐
Format::FORMAT_ALIGN_DISTRIBUTED;          // 分散对齐
Format::FORMAT_ALIGN_VERTICAL_TOP;         // 顶部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_BOTTOM;      // 底部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_CENTER;      // 垂直剧中对齐
Format::FORMAT_ALIGN_VERTICAL_JUSTIFY;     // 垂直两端对齐
Format::FORMAT_ALIGN_VERTICAL_DISTRIBUTED; // 垂直分散对齐
```

### 颜色常量

```php
Format::COLOR_BLACK
Format::COLOR_BLUE
Format::COLOR_BROWN
Format::COLOR_CYAN
Format::COLOR_GRAY
Format::COLOR_GREEN
Format::COLOR_LIME
Format::COLOR_MAGENTA
Format::COLOR_NAVY
Format::COLOR_ORANGE
Format::COLOR_PINK
Format::COLOR_PURPLE
Format::COLOR_RED
Format::COLOR_SILVER
Format::COLOR_WHITE
Format::COLOR_YELLOW
```

