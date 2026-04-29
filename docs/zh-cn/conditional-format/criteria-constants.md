# 条件常量

`criteria()` 方法接受的常量，用于在 `TYPE_CELL` / `TYPE_TEXT` / `TYPE_TIME_PERIOD` / `TYPE_AVERAGE` / `TYPE_TOP` / `TYPE_BOTTOM` 等类型下进一步细化匹配条件。

## 所属类

```php
\Vtiful\Kernel\ConditionalFormat
```

## 常量列表

```php
// 数值比较
const CRITERIA_EQUAL_TO;
const CRITERIA_NOT_EQUAL_TO;
const CRITERIA_GREATER_THAN;
const CRITERIA_LESS_THAN;
const CRITERIA_GREATER_THAN_OR_EQUAL_TO;
const CRITERIA_LESS_THAN_OR_EQUAL_TO;
const CRITERIA_BETWEEN;
const CRITERIA_NOT_BETWEEN;

// 文本匹配（配合 TYPE_TEXT 与 valueString）
const CRITERIA_TEXT_CONTAINING;
const CRITERIA_TEXT_NOT_CONTAINING;
const CRITERIA_TEXT_BEGINS_WITH;
const CRITERIA_TEXT_ENDS_WITH;

// 时间区间（配合 TYPE_TIME_PERIOD）
const CRITERIA_TIME_PERIOD_YESTERDAY;
const CRITERIA_TIME_PERIOD_TODAY;
const CRITERIA_TIME_PERIOD_TOMORROW;
const CRITERIA_TIME_PERIOD_LAST_7_DAYS;
const CRITERIA_TIME_PERIOD_LAST_WEEK;
const CRITERIA_TIME_PERIOD_THIS_WEEK;
const CRITERIA_TIME_PERIOD_NEXT_WEEK;
const CRITERIA_TIME_PERIOD_LAST_MONTH;
const CRITERIA_TIME_PERIOD_THIS_MONTH;
const CRITERIA_TIME_PERIOD_NEXT_MONTH;

// 平均值（配合 TYPE_AVERAGE）
const CRITERIA_AVERAGE_ABOVE;
const CRITERIA_AVERAGE_BELOW;

// 排名（配合 TYPE_TOP / TYPE_BOTTOM 使用 value 指定 N 或 N%）
const CRITERIA_TOP_OR_BOTTOM_PERCENT; // 把 value 解释为百分比
```

## 规则类型常量（用于色阶 / 数据条）

`minimumRule()` / `middleRule()` / `maximumRule()` 接受的常量：

```php
const RULE_MINIMUM;     // 取数据集合的最小值
const RULE_NUMBER;      // 用 minimum/middle/maximum 设置的具体数字
const RULE_PERCENT;     // 百分比（0~100）
const RULE_PERCENTILE;  // 百分位（0~100）
const RULE_FORMULA;     // 公式（用 minimumString 等给出）
const RULE_MAXIMUM;     // 取数据集合的最大值
```

## 数据条方向常量

`barDirection()` 接受：

```php
const BAR_DIRECTION_CONTEXT;       // 跟随系统默认（一般是 LTR）
const BAR_DIRECTION_LEFT_TO_RIGHT; // 从左向右
const BAR_DIRECTION_RIGHT_TO_LEFT; // 从右向左
```

## 数据条轴位置常量

`barAxisPosition()` 接受：

```php
const BAR_AXIS_AUTOMATIC; // 自动
const BAR_AXIS_MIDPOINT;  // 取中点（适合正负数据条）
const BAR_AXIS_NONE;      // 不显示轴线
```

## 使用示例

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_BETWEEN)
   ->minimum(60)
   ->maximum(90);
```
