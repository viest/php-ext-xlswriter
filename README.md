![php-excel](https://github.com/viest/php-excel-writer/blob/master/resource/logo.png)

[![Build Status](https://travis-ci.org/viest/php-ext-excel-export.svg?branch=master)](https://travis-ci.org/viest/php-ext-excel-export)

#### 1、Install the dependencies

##### Ubuntu

```bash
sudo apt-get install -y zlib1g-dev

git clone https://github.com/jmcnamara/libxlsxwriter.git && cd libxlsxwriter && make && sudo make install
```

##### Mac

```bash
brew install libxlsxwriter
```

##### Windows

> Build the basic PHP build environment.

```bash
cd PHP_BUILD_PATH/deps

git clone --recursive https://github.com/jmcnamara/MSVCLibXlsxWriter.git
```

To build the DLL of the library open the LibXlsxWriterProj/LibXlsxWriter.sln project in MS Visual Studio and build the solution using the "Build -> Build Solution" menu item.

In the default configuration this will build an x64 debug LibXlsxWriter .lib and .dll in:

```bash
PHP_BUILD_PATH\deps\MSVCLibXlsxWriter\LibXlsxWriterProj\x64\Debug
```

32Bit: Copy .dll files to c:\Windows\System32
64Bit: Same thing

Add the lib path to the `LIB` environment variable.

#### 2、Get the source code via Git

##### Unix

```bash
git clone https://github.com/viest/php-ext-excel-export.git
cd php-ext-excel-export
phpize && ./configure
make && make install
```
add the `extension=excel_writer.so` to `php.ini` file.

##### Windows

Clone the project to the ext directory in PHP, `configure` add `--with-excel_writer` parameter.

[PHP compilation tutorial](https://wiki.php.net/internals/windows/stepbystepbuild)

#### 3、Documents

[Wiki](https://github.com/viest/php-excel-writer/wiki)

#### License

Apache License 2.0
