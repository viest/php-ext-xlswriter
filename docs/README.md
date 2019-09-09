# Why use  xlswriter

Please refer to the image below. PHPExcel has been unable to work properly for memory reasons at 40,000 and 100000 points, but it can be resolved by modifying the ini configuration, but the time may take longer to complete the work;

![](.gitbook/assets/performance_comparison.png)



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

