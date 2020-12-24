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

## 为什么使用xlswriter

请参考下方对比图；由于内存原因，PHPExcel数据量`相对较大`的情况下无法正常工作，虽然可以通过`修改memory_limit`配置来解决内存问题，但完成工作的时间可能会更长;

![php-excel](resource/performance_comparison.png)

xlswriter是一个 PHP C 扩展，可用于在 Excel 2007+ XLSX 文件中读取数据，插入多个工作表，写入文本、数字、公式、日期、图表、图片和超链接。

它具备以下特性：

###### 一、写入

* 100％兼容的Excel XLSX文件
* 完整的Excel格式
* 合并单元格
* 定义工作表名称
* 过滤器
* 图表
* 数据验证和下拉列表
* 工作表PNG/JPEG图像
* 用于写入大文件的内存优化模式
* 适用于Linux，FreeBSD，OpenBSD，OS X，Windows
* 编译为32位和64位
* FreeBSD许可证
* 唯一的依赖是zlib

###### 二、读取

* 完整读取数据
* 光标读取数据
* 按数据类型读取

#### 基准测试

测试环境: Macbook Pro 13 inch, Intel Core i5, 16GB 2133MHz LPDDR3 Memory, 128GB SSD Storage.

##### 导出

> 两种内存模式导出100万行数据（单行27列，数据类型均为字符串，单个字符串长度为19）

* 普通模式：耗时 `29S`，内存只需 `2083MB`；
* 固定内存模式：仅需 `52S`，内存仅需 `<1MB`；

##### 导入

> 100万行数据（单行1列，数据类型为INT）

* 全量模式：耗时 `3S`，内存仅 `558MB`；
* 游标模式：耗时 `2.8S`，内存仅 `<1MB`；

## 从这里开始

[文档|Documents](https://xlswriter-docs.viest.me/)

## PECL 仓库

[![pecl](resource/pecl.png)](https://pecl.php.net/package/xlswriter)

## IDE Helper

```bash
composer require viest/php-ext-xlswriter-ide-helper:dev-master
```

## 交流群

<img width="160" src="resource/qq.jpg"/>

## 贡献者

### 代码贡献者

这个项目的存在要感谢所有贡献者。 [[Contribute](CONTRIBUTING.md)].
<a href="https://github.com/viest/php-ext-xlswriter/graphs/contributors"><img src="https://opencollective.com/php-ext-xlswriter/contributors.svg?width=890&button=false" /></a>

### 财务捐赠者

成为财务捐赠者，并帮助我们维持我们的社区。[[Contribute](https://opencollective.com/php-ext-xlswriter/contribute)]

#### 个人

<a href="https://opencollective.com/php-ext-xlswriter"><img src="https://opencollective.com/php-ext-xlswriter/individuals.svg?width=890"></a>

#### 组织机构

与您的组织一起支持该项目。您的徽标将显示在此处，并带有指向您网站的链接。[[Contribute](https://opencollective.com/php-ext-xlswriter/contribute)]

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

BSD license

[![FOSSA Status](https://app.fossa.io/api/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter.svg?type=large)](https://app.fossa.io/projects/git%2Bgithub.com%2Fviest%2Fphp-ext-xlswriter?ref=badge_large)

## Stargazers over time

[![Stargazers over time](https://starchart.cc/viest/php-ext-xlswriter.svg)](https://starchart.cc/viest/php-ext-xlswriter)
