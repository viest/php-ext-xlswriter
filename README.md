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

>Please look forward to the next update.

#### 3、Documents

[Wiki](https://github.com/viest/php-excel-writer/wiki)

#### License

Apache License 2.0
