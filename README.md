<div align=center>
<img height="214" src="resource/logo_now.png"/>
</div>

<div align=center>
<a href="https://github.com/viest/php-ext-xlswriter/releases"><img src="https://img.shields.io/github/release/viest/php-ext-excel-export.svg"/></a>
</div>

<div align=center>
<a href="https://github.com/viest/php-ext-xlswriter"><img src="https://img.shields.io/badge/platform-macos%20%7C%20linux%20%7C%20windows-brightgreen.svg"/></a>
</div>

<div align=center>
<a href="https://travis-ci.com/viest/php-ext-xlswriter"><img src="https://travis-ci.com/viest/php-ext-xlswriter.svg?branch=master"/></a>
<a href="https://ci.appveyor.com/project/viest/php-ext-excel-export/branch/master"><img src="https://ci.appveyor.com/api/projects/status/w4cfjo9e4gsrs6rn/branch/master?svg=true"/></a>
<a href="https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_shield"><img src="https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=shield"/></a>
</div>

<div align=center>
<a href="https://opencollective.com/php-ext-xlswriter"><img src="https://opencollective.com/php-ext-xlswriter/all/badge.svg?label=financial+contributors"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img src="https://img.shields.io/badge/PHP-%3E%3D%207.0-brightgreen.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img src="https://img.shields.io/github/contributors/viest/php-ext-excel-export.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img src="https://img.shields.io/badge/license-BSD-green.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img src="https://img.shields.io/github/issues/viest/php-ext-excel-export.svg"/></a>
<a href="https://hits.seeyoufarm.com"><img src="https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https%3A%2F%2Fgithub.com%2Fviest%2Fphp-ext-xlswriter&count_bg=%2379C83D&title_bg=%23555555&icon=&icon_color=%23E7E7E7&title=hits&edge_flat=false"/></a>
</div>

## Why use xlswriter

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

#### Install

###### Unix

```bash
pecl install xlswriter
```

###### Windows

[download dll](https://github.com/viest/php-ext-xlswriter/releases)

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

## [Documents](https://xlswriter-docs.viest.me/)

Includes extensive and detailed instructions that make it easy to get started with xlswriter.

## PECL Repository

[![pecl](resource/pecl.png)](https://pecl.php.net/package/xlswriter)

## IDE Helper

```bash
composer require viest/php-ext-xlswriter-ide-helper:dev-master
```

## Exchange group

<img width="160" src="resource/qq.jpg"/>

## Financial donation

<img height="220" src="resource/pay.jpg"/>

## Contributors

### Code Contributors

This project exists thanks to all the people who contribute. [[Contribute](CONTRIBUTING.md)].
<a href="https://github.com/viest/php-ext-xlswriter/graphs/contributors"><img src="https://opencollective.com/php-ext-xlswriter/contributors.svg?width=890&button=false" /></a>

### Financial Contributors

Become a financial contributor and help us sustain our community. [[Contribute](https://opencollective.com/php-ext-xlswriter/contribute)]

#### Individuals

<a href="https://opencollective.com/php-ext-xlswriter"><img src="https://opencollective.com/php-ext-xlswriter/individuals.svg?width=890"></a>

#### Organizations

Support this project with your organization. Your logo will show up here with a link to your website. [[Contribute](https://opencollective.com/php-ext-xlswriter/contribute)]

<a href="https://opencollective.com/php-ext-xlswriter/organization/0/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/0/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/1/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/1/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/2/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/2/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/3/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/3/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/4/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/4/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/5/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/5/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/6/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/6/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/7/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/7/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/8/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/8/avatar.svg"></a>
<a href="https://opencollective.com/php-ext-xlswriter/organization/9/website"><img src="https://opencollective.com/php-ext-xlswriter/organization/9/avatar.svg"></a>

## License

BSD License

[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=large)](https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_large)

## Stargazers over time

[![Stargazers over time](https://starchart.cc/viest/php-ext-xlswriter.svg)](https://starchart.cc/viest/php-ext-xlswriter)
