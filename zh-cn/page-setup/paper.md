# 设置纸张大小

## 函数原型

```php
setPaper(int $parper): array
```

## 纸张大小常量

### 所属类

```php
\Vtiful\Kernel\Excel
```

### 常量列表

```php
const PAPER_DEFAULT              = 0;
const PAPER_LETTER               = 1;
const PAPER_LETTER_SMALL         = 2;
const PAPER_TABLOID              = 3;
const PAPER_LEDGER               = 4;
const PAPER_LEGAL                = 5;
const PAPER_STATEMENT            = 6;
const PAPER_EXECUTIVE            = 7;
const PAPER_A3                   = 8;
const PAPER_A4                   = 9;
const PAPER_A4_SMALL             = 10;
const PAPER_A5                   = 11;
const PAPER_B4                   = 12;
const PAPER_B5                   = 13;
const PAPER_FOLIO                = 14;
const PAPER_QUARTO               = 15;
const PAPER_NOTE                 = 18;
const PAPER_ENVELOPE_9           = 19;
const PAPER_ENVELOPE_10          = 20;
const PAPER_ENVELOPE_11          = 21;
const PAPER_ENVELOPE_12          = 22;
const PAPER_ENVELOPE_14          = 23;
const PAPER_C_SIZE_SHEET         = 24;
const PAPER_D_SIZE_SHEET         = 25;
const PAPER_E_SIZE_SHEET         = 26;
const PAPER_ENVELOPE_DL          = 27;
const PAPER_ENVELOPE_C3          = 28;
const PAPER_ENVELOPE_C4          = 29;
const PAPER_ENVELOPE_C5          = 30;
const PAPER_ENVELOPE_C6          = 31;
const PAPER_ENVELOPE_C65         = 32;
const PAPER_ENVELOPE_B4          = 33;
const PAPER_ENVELOPE_B5          = 34;
const PAPER_ENVELOPE_B6          = 35;
const PAPER_ENVELOPE_1           = 36;
const PAPER_MONARCH              = 37;
const PAPER_ENVELOPE_2           = 38;
const PAPER_FANFOLD              = 39;
const PAPER_GERMAN_STD_FANFOLD   = 40;
const PAPER_GERMAN_LEGAL_FANFOLD = 41;
```

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setPaper(\Vtiful\Kernel\Excel::PAPER_A3)
    ->setLandscape()
    ->output();
```
