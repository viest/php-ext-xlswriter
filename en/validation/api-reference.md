# Validation API reference

`Vtiful\Kernel\Validation` builds a single data-validation rule. Every method except `__construct()` and `toResource()` returns `$this`, so calls can be chained.

## Class constants

### Validation types — `TYPE_*`

```
TYPE_INTEGER          TYPE_INTEGER_FORMULA
TYPE_DECIMAL          TYPE_DECIMAL_FORMULA
TYPE_LIST             TYPE_LIST_FORMULA
TYPE_DATE             TYPE_DATE_FORMULA       TYPE_DATE_NUMBER
TYPE_TIME             TYPE_TIME_FORMULA       TYPE_TIME_NUMBER
TYPE_LENGTH           TYPE_LENGTH_FORMULA
TYPE_CUSTOM_FORMULA
TYPE_ANY
```

### Criteria — `CRITERIA_*`

```
CRITERIA_BETWEEN              CRITERIA_NOT_BETWEEN
CRITERIA_EQUAL_TO             CRITERIA_NOT_EQUAL_TO
CRITERIA_GREATER_THAN         CRITERIA_LESS_THAN
CRITERIA_GREATER_THAN_OR_EQUAL_TO
CRITERIA_LESS_THAN_OR_EQUAL_TO
```

### Error severity — `ERROR_TYPE_*`

```
ERROR_TYPE_STOP
ERROR_TYPE_WARNING
ERROR_TYPE_INFORMATION
```

## Function prototypes

```
\Vtiful\Kernel\Validation::__construct()
\Vtiful\Kernel\Validation::validationType(int $type): self
\Vtiful\Kernel\Validation::criteriaType(int $criteria): self
\Vtiful\Kernel\Validation::ignoreBlank(bool $ignore = true): self
\Vtiful\Kernel\Validation::showInput(bool $show = true): self
\Vtiful\Kernel\Validation::showError(bool $show = true): self
\Vtiful\Kernel\Validation::errorType(int $type): self
\Vtiful\Kernel\Validation::dropdown(bool $on = true): self

\Vtiful\Kernel\Validation::valueNumber(float $value): self
\Vtiful\Kernel\Validation::valueFormula(string $formula): self
\Vtiful\Kernel\Validation::valueList(array $values): self
\Vtiful\Kernel\Validation::valueDatetime(int $timestamp): self

\Vtiful\Kernel\Validation::minimumNumber(float $value): self
\Vtiful\Kernel\Validation::minimumFormula(string $formula): self
\Vtiful\Kernel\Validation::minimumDatetime(int $timestamp): self

\Vtiful\Kernel\Validation::maximumNumber(float $value): self
\Vtiful\Kernel\Validation::maximumFormula(string $formula): self
\Vtiful\Kernel\Validation::maximumDatetime(int $timestamp): self

\Vtiful\Kernel\Validation::inputTitle(string $title): self
\Vtiful\Kernel\Validation::inputMessage(string $message): self
\Vtiful\Kernel\Validation::errorTitle(string $title): self
\Vtiful\Kernel\Validation::errorMessage(string $message): self

\Vtiful\Kernel\Validation::toResource(): resource
```

### **int $type**

> Validation type — one of the `Validation::TYPE_*` constants.

### **int $criteria**

> Criteria — one of the `Validation::CRITERIA_*` constants. `CRITERIA_BETWEEN` and `CRITERIA_NOT_BETWEEN` require both `minimum*` and `maximum*` to be set; the other comparison criteria use `value*`.

### **bool $ignore**

> Whether blank cells are accepted. Defaults to `true`.

### **bool $show**

> Whether the input prompt (`showInput`) or error alert (`showError`) is displayed when the cell is selected or invalid. Defaults to `true`.

### **bool $on**

> Whether to show the in-cell drop-down arrow. Defaults to `true`. Only meaningful for `TYPE_LIST` / `TYPE_LIST_FORMULA`.

### **float $value**

> Target value for comparison criteria, e.g. `valueNumber(20)` compares against 20.

### **string $formula**

> Express the bound as an Excel formula, e.g. `valueFormula('=A1')`, `minimumFormula('=$A$1')`.

### **array $values**

> List of strings used for `TYPE_LIST`. Each element must be a non-empty string, otherwise `Vtiful\Exception` is thrown.

### **int $timestamp**

> Unix timestamp; the extension converts it to an Excel serial date.

### **string $title / $message**

> Title and body for the input prompt and error alert.

### **toResource()**

> Convert the current `Validation` builder into a resource handle that `Excel::validation()` accepts.

## Example

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_BETWEEN)
    ->minimumNumber(1)
    ->maximumNumber(100)
    ->ignoreBlank(true)
    ->errorType(\Vtiful\Kernel\Validation::ERROR_TYPE_STOP)
    ->inputTitle('Score')
    ->inputMessage('Integer between 1 and 100')
    ->errorTitle('Invalid value')
    ->errorMessage('Score must be an integer in 1..100');

$excel->fileName('tutorial.xlsx')
    ->header(['Score'])
    ->validation('A1:A10', $validation->toResource())
    ->output();
```
