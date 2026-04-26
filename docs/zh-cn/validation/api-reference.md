# Validation 类完整 API

> _此页面正在补充中（v1.6.0+ 新功能）。_

## 方法签名

`Validation::validationType(int $type)`

`Validation::criteriaType(int $criteria)`

`Validation::ignoreBlank(bool $ignore = true)`

`Validation::showInput(bool $show = true)`

`Validation::showError(bool $show = true)`

`Validation::errorType(int $type)`

`Validation::dropdown(bool $on = true)`

`Validation::valueNumber(float $value)`

`Validation::valueFormula(string $formula)`

`Validation::valueList(array $values)`

`Validation::valueDatetime(int $timestamp)`

`Validation::minimumNumber(float $value)` / `minimumFormula(string)` / `minimumDatetime(int)`

`Validation::maximumNumber(float $value)` / `maximumFormula(string)` / `maximumDatetime(int)`

`Validation::inputTitle(string $title)`

`Validation::inputMessage(string $message)`

`Validation::errorTitle(string $title)`

`Validation::errorMessage(string $message)`

`Validation::toResource()`
