# Criteria constants

Constants accepted by `criteria()`. They refine the match for `TYPE_CELL` / `TYPE_TEXT` / `TYPE_TIME_PERIOD` / `TYPE_AVERAGE` / `TYPE_TOP` / `TYPE_BOTTOM` rules.

## Defining class

```php
\Vtiful\Kernel\ConditionalFormat
```

## Constants

```php
// Numeric comparison
const CRITERIA_EQUAL_TO;
const CRITERIA_NOT_EQUAL_TO;
const CRITERIA_GREATER_THAN;
const CRITERIA_LESS_THAN;
const CRITERIA_GREATER_THAN_OR_EQUAL_TO;
const CRITERIA_LESS_THAN_OR_EQUAL_TO;
const CRITERIA_BETWEEN;
const CRITERIA_NOT_BETWEEN;

// Text matching (with TYPE_TEXT and valueString)
const CRITERIA_TEXT_CONTAINING;
const CRITERIA_TEXT_NOT_CONTAINING;
const CRITERIA_TEXT_BEGINS_WITH;
const CRITERIA_TEXT_ENDS_WITH;

// Time periods (with TYPE_TIME_PERIOD)
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

// Average (with TYPE_AVERAGE)
const CRITERIA_AVERAGE_ABOVE;
const CRITERIA_AVERAGE_BELOW;

// Ranking (with TYPE_TOP / TYPE_BOTTOM; use value() for N or N%)
const CRITERIA_TOP_OR_BOTTOM_PERCENT; // Interprets value() as a percentage
```

## Rule type constants (colour scales / data bars)

Accepted by `minimumRule()` / `middleRule()` / `maximumRule()`:

```php
const RULE_MINIMUM;     // Minimum of the data
const RULE_NUMBER;      // Literal number set with minimum/middle/maximum
const RULE_PERCENT;     // Percentage (0..100)
const RULE_PERCENTILE;  // Percentile (0..100)
const RULE_FORMULA;     // Formula (provided via minimumString etc.)
const RULE_MAXIMUM;     // Maximum of the data
```

## Bar direction constants

Accepted by `barDirection()`:

```php
const BAR_DIRECTION_CONTEXT;       // Follow the system default (typically LTR)
const BAR_DIRECTION_LEFT_TO_RIGHT;
const BAR_DIRECTION_RIGHT_TO_LEFT;
```

## Bar axis position constants

Accepted by `barAxisPosition()`:

```php
const BAR_AXIS_AUTOMATIC; // Auto
const BAR_AXIS_MIDPOINT;  // Midpoint (suitable for bars with positive and negative values)
const BAR_AXIS_NONE;      // No axis
```

## Example

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_BETWEEN)
   ->minimum(60)
   ->maximum(90);
```
