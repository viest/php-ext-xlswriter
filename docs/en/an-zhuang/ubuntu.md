# Ubuntu

#### 安装依赖

```bash
apt-get install -y zlib1g-dev
```

#### 编译扩展

```bash
git clone https://github.com/viest/php-ext-excel-export

cd php-ext-excel-export

git submodule update --init

phpize && ./configure --with-php-config=/path/to/php-config --enable-reader

make && make install
```

#### 功能测试

```bash
make && make test
```

#### 修改 php.ini

```text
extension = xlswriter.so
```

