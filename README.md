<div align=center>
<img alt="php-ext-xlswriter" height="214" src="resource/logo_now.png"/>
</div>

<div align=center>
<a href="https://github.com/viest/php-ext-xlswriter/releases"><img alt="php-ext-xlswriter" src="https://img.shields.io/github/release/viest/php-ext-excel-export.svg"/></a>
</div>

<div align=center>
<a href="https://github.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://img.shields.io/badge/platform-macos%20%7C%20linux%20%7C%20windows-brightgreen.svg"/></a>
</div>

<div align=center>
<a href="https://gitee.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://gitee.com/viest/php-ext-xlswriter/badge/star.svg?theme=gvp"/></a>
<a href="https://gitcode.com/viest/php-ext-xlsxwriter"><img alt="php-ext-xlswriter" src="https://gitcode.com/viest/php-ext-xlsxwriter/star/badge.svg"/></a>
</div>

<div align=center>
<a href="https://github.com/viest/php-ext-xlswriter/actions"><img alt="php-ext-xlswriter" src="https://img.shields.io/github/actions/workflow/status/viest/php-ext-xlswriter/main.yml?branch=master&logo=github"></a>
<a href="https://ci.appveyor.com/project/viest/php-ext-xlswriter/branch/master"><img alt="php-ext-xlswriter" src="https://ci.appveyor.com/api/projects/status/w4cfjo9e4gsrs6rn/branch/master?svg=true"/></a>
<a href="https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_shield"><img alt="php-ext-xlswriter" src="https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=shield"/></a>
</div>

<div align=center>
<a href="https://opencollective.com/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://opencollective.com/php-ext-xlswriter/all/badge.svg?label=financial+contributors"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://img.shields.io/badge/PHP-%3E%3D%207.0-brightgreen.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://img.shields.io/github/contributors/viest/php-ext-excel-export.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://img.shields.io/badge/license-BSD-green.svg"/></a>
<a href="https://github.com/viest/php-ext-xlswriter"><img alt="php-ext-xlswriter" src="https://img.shields.io/github/issues/viest/php-ext-excel-export.svg"/></a>
</div>

## Why use xlswriter

The chart below compares xlswriter with PhpSpreadsheet (the maintained successor to PHPExcel) when exporting an XLSX file, scaled all the way to Excel's row limit. Writing 1,048,576 rows × 10 columns, xlswriter is about 20× faster, and its fixed-memory mode keeps peak memory flat at ~30 MB no matter how many rows you write — whereas a pure-PHP library's memory grows with the data (≈7 GB for the same file).

![xlswriter vs PhpSpreadsheet performance](resource/performance_comparison.png)

> The two xlswriter modes track within ~10% of each other on time. Fixed-memory mode is marginally faster because it streams each row straight to disk and frees it immediately — a single pass, with no full in-memory model to build up and then serialize a second time at the end. The trade-off is that, unlike normal mode, it can no longer revisit a cell once it has been written (and its strings are stored inline rather than de-duplicated, so the file can be slightly larger). Normal mode keeps the whole workbook in memory, which is what lets you write cells in any order and re-style them before saving.

xlswriter is a PHP C Extension for Excel 2007+ XLSX files. It writes text, numbers, formulas, dates, charts, images and hyperlinks to new workbooks, opens existing files to edit them and save the result, reads their contents back, and evaluates formulas to a computed value. It supports features such as:

###### Writer

* 100% compatible Excel XLSX files.
* Full Excel formatting.
* Merged cells.
* Defined names.
* Autofilters.
* Charts.
* Data validation and drop down lists.
* Conditional formatting.
* Rich text, comments and hyperlinks.
* Worksheet PNG/JPEG images.
* Edit existing workbooks — open a file, update cell values, styles, merged ranges and row/column sizes, add worksheets, images and charts, then save the result.
* Formula calculation — evaluate a formula and get its computed value, and write formulas with a pre-computed cached result.
* Memory optimization mode for writing large files.
* Works on Linux, FreeBSD, OpenBSD, OS X, Windows.
* Compiles for 32 and 64 bit.
* FreeBSD License.
* The only dependency is on zlib.

###### Reader

* Full read mode and cursor read mode.
* Read by data type.
* Read cell styles and number formats.
* Read merged cells.
* Read images, charts and comments.
* Read formulas together with their cached values.

#### Install

###### Unix

xlswriter requires the zlib development headers at build time. Install them
first if they are missing (common on minimal images):

```bash
# Debian / Ubuntu
apt-get install -y zlib1g-dev
# Alpine
apk add zlib-dev
# RHEL / CentOS / Fedora
yum install -y zlib-devel
```

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
