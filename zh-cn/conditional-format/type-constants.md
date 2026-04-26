# 类型常量

`type()` 方法接受的常量，用于声明条件格式规则的整体类型。

## 所属类

```php
\Vtiful\Kernel\ConditionalFormat
```

## 常量列表

```php
const TYPE_CELL          = ?; // 单元格值规则（配合 criteria + value/minimum/maximum）
const TYPE_TEXT          = ?; // 文本规则（包含 / 不含 / 开头 / 结尾）
const TYPE_TIME_PERIOD   = ?; // 日期/时间规则（昨天 / 今天 / 本周 等）
const TYPE_AVERAGE       = ?; // 平均值规则（高于 / 低于均值）
const TYPE_DUPLICATE     = ?; // 重复值
const TYPE_UNIQUE        = ?; // 唯一值
const TYPE_TOP           = ?; // 排名前 N 或前 N%
const TYPE_BOTTOM        = ?; // 排名后 N 或后 N%
const TYPE_BLANKS        = ?; // 空白单元格
const TYPE_NO_BLANKS     = ?; // 非空白单元格
const TYPE_ERRORS        = ?; // 错误值
const TYPE_NO_ERRORS     = ?; // 非错误值
const TYPE_FORMULA       = ?; // 自定义公式（valueString 设置公式）
const TYPE_2_COLOR_SCALE = ?; // 双色色阶
const TYPE_3_COLOR_SCALE = ?; // 三色色阶
const TYPE_DATA_BAR      = ?; // 数据条
const TYPE_ICON_SETS     = ?; // 图标集
```

> 上方等号右侧的具体数值由扩展内部决定，请直接使用常量名称引用，不要硬编码数字。

## 使用示例

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
   ->value(50);
```
