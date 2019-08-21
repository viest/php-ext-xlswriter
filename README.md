![php-excel](https://github.com/viest/php-excel-writer/blob/master/resource/logo.png)

[![Build Status](https://travis-ci.org/viest/php-ext-excel-export.svg?branch=master)](https://travis-ci.org/viest/php-ext-excel-export)
[![Build status](https://ci.appveyor.com/api/projects/status/w4cfjo9e4gsrs6rn/branch/master?svg=true)](https://ci.appveyor.com/project/viest/php-ext-excel-export/branch/master)
[![](https://img.shields.io/github/release/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-excel-export)
[![](https://img.shields.io/badge/PHP-%3E%3D%207.0-brightgreen.svg)](https://github.com/viest/php-ext-excel-export)
[![](https://img.shields.io/github/contributors/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-excel-export)
[![](https://img.shields.io/badge/platform-macos%20%7C%20linux%20%7C%20windows-brightgreen.svg)](https://github.com/viest/php-ext-excel-export)
[![](https://img.shields.io/badge/license-BSD-green.svg)](https://github.com/viest/php-ext-excel-export)
[![](https://img.shields.io/github/issues/viest/php-ext-excel-export.svg)](https://github.com/viest/php-ext-excel-export)

#### Why use xlswriter

Please refer to the image below. PHPExcel has been unable to work properly for memory reasons at 40,000 and 100000 points, but it can be resolved by modifying the ini configuration, but the time may take longer to complete the work;

![php-excel](https://github.com/viest/php-excel-writer/blob/master/resource/performance_comparison.png)

xlswriter is a PHP C Extension that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file. It supports features such as:

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

[文档|Documents](https://xlswriter-docs.viest.me/)

[![pecl](https://github.com/viest/php-excel-writer/blob/master/resource/pecl.png)](https://pecl.php.net/package/xlswriter)

#### License

FreeBSD license
