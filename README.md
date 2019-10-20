![php-excel](resource/logo.png)

[![Build Status](https://travis-ci.org/viest/php-ext-xlswriter.svg?branch=master)](https://travis-ci.org/viest/php-ext-xlswriter)
[![Build status](https://ci.appveyor.com/api/projects/status/w4cfjo9e4gsrs6rn/branch/master?svg=true)](https://ci.appveyor.com/project/viest/php-ext-excel-export/branch/master)
[![](https://img.shields.io/github/release/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-xlswriter)
[![](https://img.shields.io/badge/PHP-%3E%3D%207.0-brightgreen.svg)](https://github.com/viest/php-ext-xlswriter)
[![](https://img.shields.io/github/contributors/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-xlswriter)
[![](https://img.shields.io/badge/platform-macos%20%7C%20linux%20%7C%20windows-brightgreen.svg)](https://github.com/viest/php-ext-xlswriter)
[![](https://img.shields.io/badge/license-BSD-green.svg)](https://github.com/viest/php-ext-xlswriter)
[![](https://img.shields.io/github/issues/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-xlswriter)
[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=shield)](https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_shield)

#### Why use xlswriter

Please refer to the image below. PHPExcel has been unable to work properly for memory reasons at 40,000 and 100000 points, but it can be resolved by modifying the ini configuration, but the time may take longer to complete the work;

![php-excel](resource/performance_comparison.png)

xlswriter is a PHP C Extension that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file. It supports features such as:

###### Writer

* 100% compatible Excel XLSX files.
* Full Excel formatting.
* Merged cells.
* Defined names.
* Autofilters.
* Charts.
* Data validation and drop down lists.
* Worksheet PNG/JPEG images.
* Memory optimization mode for writing large files.
* Works on Linux, FreeBSD, OpenBSD, OS X, Windows.
* Compiles for 32 and 64 bit.
* FreeBSD License.
* The only dependency is on zlib.

###### Reader

* Full read data
* Cursor read data
* Read by data type

#### Benchmark

Test environment: Macbook Pro 13 inch, Intel Core i5, 16GB 2133MHz LPDDR3 Memory, 128GB SSD Storage.

##### Export

> Two memory modes export 1 million rows of data (27 columns, data is string)

* Normal mode: only 29S is needed, and the memory only needs 2083MB;
* Fixed memory mode: only need 52S, memory only needs <1MB;

##### Import

> 1 million rows of data (1 columns, data is inter)

* Full mode: Just 3S, the memory is only 558MB;
* Cursor mode: Just 2.8S, memory is only <1MB;

#### Documents

[文档|Documents](https://xlswriter-docs.viest.me/)

#### PECL Repository

[![pecl](resource/pecl.png)](https://pecl.php.net/package/xlswriter)

#### IDE Helper

```bash
composer require viest/php-ext-xlswriter-ide-helper:dev-master
```

#### Exchange group

<img width="160" src="resource/qq.jpg"/>

#### License

FreeBSD license


[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=large)](https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_large)