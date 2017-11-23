PHP-Excel-Writer

#### 1、Install the dependencies

```bash
sudo apt-get install -y zlib1g-dev

git clone https://github.com/jmcnamara/libxlsxwriter.git

cd libxlsxwriter

make
sudo make install
```

#### 2、Get the source code via Git

```bash
git clone https://github.com/viest/php-excel-writer.git

cd php-excel-writer

phpize 

./configure

make

make install
```

#### 3、Examples

```php
try {
    $config = [
        'path' => '/vagrant/'
    ];

    $excel = new \Vtiful\Kernel\Excel($config);

    for($a = 0; $a < 200000; ++$a) {
        $data[$a] = ['viest', 20];
    }

    $excel->fileName("test.xlsx")
        ->header(['name', 'age'])
        ->data([
            ['viest', 20]
        ])
        ->output();

} catch (\Exception $exception) {
    //....
}
```

#### License

Apache License 2.0