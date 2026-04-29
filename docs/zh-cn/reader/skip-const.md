# 忽略跳过动作常量

```php
const SKIP_NONE = 0x00;         // 不忽略任何单元格、行
const SKIP_EMPTY_ROW = 0x01;    // 忽略空行
const SKIP_EMPTY_CELLS = 0x02;  // 忽略空单元格（肉眼观察单元格内无数据，并不代表单元格未定义、未使用）
const SKIP_EMPTY_VALUE = 0X100; // 忽略单元格空数据
```