# Insert link

### **Function Prototype**

```php
insertUrl(int $row, int $column, string $url[, resource $format])
```

#### **int $row**

> cell row

#### **int $column**

> cell column

#### **string $url**

> link address

#### **resource $format**

> link style

###example

```php
$excel = new \Vtiful\Kernel\Excel($config);

$urlFile = $excel->fileName("free.xlsx")
     ->header(['url']);

$fileHandle = $fileObject->getHandle();

$format   = new \Vtiful\Kernel\Format($fileHandle);
$urlStyle = $format->bold()
     ->underline(Format::UNDERLINE_SINGLE)
     ->toResource();

$urlFile->insertUrl(1, 0, 'https://github.com', $urlStyle);

$textFile->output();
```