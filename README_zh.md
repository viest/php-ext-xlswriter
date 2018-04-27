## Excel-Export

Authon: viest [dev@service.viest.me](mailto:dev@service.viest.me)

-------

- **安装**
    - **[编译安装 Excel-Export](#编译安装)**
- **使用**
    - **[创建一个简单的xlsx文件](#创建一个简单的xlsx文件)**
    - **[1、向单元格插入文字](#向单元格插入文字)**
    - **[2、单元格插入公式](#单元格插入公式)**
    - **[3、单元格插入本地图片](#单元格插入本地图片)**
    - **[4、数据过滤](#数据过滤)**
    - **[5、合并单元格](#合并单元格)**
    - **[6、设置列单元格格式](#设置列单元格格式)**
    - **[7、设置行单元格格式](#设置行单元格格式)**
    - **[8、固定内存导出](#固定内存导出)**

--------

### 编译安装

#### Unix

##### Ubuntu

```bash
# 依赖

sudo apt-get install -y zlib1g-dev

git clone https://github.com/jmcnamara/libxlsxwriter.git

cd libxlsxwriter && make && sudo make install

# Excel-Export

git clone https://github.com/viest/php-ext-excel-export.git

cd php-ext-excel-export

phpize && ./configure --with-php-config=/path/to/php-config

make && make install

# 添加 extension = xlswriter.so 到 ini 配置
```

##### Mac

```bash
# 依赖

brew install libxlsxwriter

# Excel-Export

git clone https://github.com/viest/php-ext-excel-export.git

cd php-ext-excel-export

phpize && ./configure --with-php-config=/path/to/php-config

make && make install

# 添加 extension = xlswriter.so 到 ini 配置
```

#### Windows

##### 依赖

> 请预先搭建PHP编译环境，教程详见 [php.net](https://wiki.php.net/internals/windows/stepbystepbuild)

```bash
cd PHP_BUILD_PATH/deps

git clone --recursive https://github.com/jmcnamara/MSVCLibXlsxWriter.git
```

要构建依赖库的DLL，打开LibXlsxWriterProj/LibXlsxWriter。在MS Visual Studio中使用sln项目，并使用`"构建 ->构建解决方案"`菜单项构建解决方案。

在默认配置中，这将构建一个`x64调试`LibXlsxWriter .lib和.dll，你也可以选择`释放`版本构建，构建完成后，你可以在一下路径中找到对应的DLL

```bash
PHP_BUILD_PATH\deps\MSVCLibXlsxWriter\LibXlsxWriterProj\[x64|x86]\Debug

# or

PHP_BUILD_PATH\deps\MSVCLibXlsxWriter\LibXlsxWriterProj\[x64|x86]\Release
```

32\64位: 复制 .dll 文件到 c:\Windows\System*

最后添加系统环境变量LIB，指向DLL文件目录

##### 扩展

```bash
cd PHP_PATH/ext

git clone https://github.com/viest/php-ext-excel-export.git

cd PHP_PATH

buildconf.bat

configure.bat --disable-all --enable-cli --with-xlswriter

nmake
```

### 创建一个简单的xlsx文件

```php
$config = ['path' => '/home/viest'];

$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial01.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Rent', 1000],
        ['Gas',  100],
        ['Food', 300],
        ['Gym',  50],
    ])
    ->output();
```

### 向单元格插入文字

#### 语法

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

### 单元格插入公式

#### 语法

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

#### 语法

```php
insertImage(int $row, int $column, string $localImagePath)
```

##### int $row

> 单元格所在行

##### int $column

> 单元格所在列

##### string $localImagePath

> 图片路径

##### 实例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx");

$freeFile->insertImage(5, 0, '/vagrant/ASW-G-66.jpg');

$freeFile->output();
```

### 数据过滤

#### 语法

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

#### 语法

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

#### 语法

```php
setColumn(resourch $format, string $range[, double $width]);
```

##### string $format

> 单元格样式

##### string $range

> 单元格范围

##### double $width

> 单元格宽度

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$boldStyle = \Vtiful\Kernel\Format::bold($fileHandle);

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setColumn($boldStyle, 'A:A', 200)
    ->output();
```

### 设置行单元格格式

#### 语法

```php
setRow(resourch $format, string $range[, double $height]);
```

##### string $format

> 单元格样式

##### string $range

> 单元格范围

##### double $height

> 单元格高度

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$boldStyle = \Vtiful\Kernel\Format::bold($fileHandle);

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setRow($boldStyle, 'A1')
    ->output();
```

### 固定内存导出

#### 内存

最大内存使用量 = 最大一行的数据占用量

#### 语法

```php
constMemory(string $fileName);
```

##### 实例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->constMemory('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$boldStyle = \Vtiful\Kernel\Format::bold($fileHandle);

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setRow($boldStyle, 'A1')
    ->output();```

