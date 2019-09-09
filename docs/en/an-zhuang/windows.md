# Windows

#### Building a PHP build environment

See [php.net](https://wiki.php.net/internals/windows/stepbystepbuild)

#### Installation dependencies

```bash
cd PHP_BUILD_PATH/deps

DownloadFile http://zlib.net/zlib-1.2.11.tar.gz

7z x zlib-1.2.11.tar.gz > NUL
7z x zlib-1.2.11.tar > NUL

cd zlib-1.2.11

cmake -G "Visual Studio 14 2015" -DCMAKE_BUILD_TYPE="Release" -DCMAKE_C_FLAGS_RELEASE="/MT"
cmake --build . --config "Release"
```

#### Compiling extensions

```bash
cd PHP_PATH/ext

git clone https://github.com/viest/php-ext-excel-export.git

cd php-ext-excel-export

git submodule update --init

phpize

configure.bat --with-xlswriter --with-extra-libs=PATH\zlib-1.2.11\Release --with-extra-includes=PATH\zlib-1.2.11

nmake
```

