# Ubuntu

#### Installation dependencies

```bash
apt-get install -y zlib1g-dev
```

#### Compiling extensions

```bash
git clone https://github.com/viest/php-ext-excel-export

cd php-ext-excel-export

git submodule update --init

phpize && ./configure --with-php-config=/path/to/php-config --enable-reader

make && make install
```

#### Function test

```bash
make && make test
```

#### Modify php.ini

```text
extension = xlswriter.so
```

