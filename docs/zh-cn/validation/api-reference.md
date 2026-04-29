# Validation 类完整 API

`Vtiful\Kernel\Validation` 用于构造一条数据验证规则。除 `__construct()` 与 `toResource()` 之外，所有方法均返回 `$this`，可链式调用。

## 类常量

### 验证类型 `TYPE_*`

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

### 约束类型 `CRITERIA_*`

```
CRITERIA_BETWEEN              CRITERIA_NOT_BETWEEN
CRITERIA_EQUAL_TO             CRITERIA_NOT_EQUAL_TO
CRITERIA_GREATER_THAN         CRITERIA_LESS_THAN
CRITERIA_GREATER_THAN_OR_EQUAL_TO
CRITERIA_LESS_THAN_OR_EQUAL_TO
```

### 错误提示等级 `ERROR_TYPE_*`

```
ERROR_TYPE_STOP
ERROR_TYPE_WARNING
ERROR_TYPE_INFORMATION
```

## 函数原型

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

> 验证类型，对应 `Validation::TYPE_*` 常量。

### **int $criteria**

> 约束类型，对应 `Validation::CRITERIA_*` 常量。`CRITERIA_BETWEEN` / `CRITERIA_NOT_BETWEEN` 需要同时设置 `minimum*` 与 `maximum*`；其余比较类约束使用 `value*`。

### **bool $ignore**

> 是否忽略空白单元格，默认 `true`。

### **bool $show**

> 是否显示提示信息（`showInput`）或错误信息（`showError`），默认 `true`。

### **bool $on**

> 是否显示下拉箭头，默认 `true`，仅对 `TYPE_LIST` / `TYPE_LIST_FORMULA` 有效。

### **float $value**

> 比较类约束的目标值，例如 `valueNumber(20)` 表示与 20 比较。

### **string $formula**

> 用公式表达取值，例如 `valueFormula('=A1')`、`minimumFormula('=$A$1')`。

### **array $values**

> 仅用于 `TYPE_LIST`，传入字符串数组作为下拉项。数组元素必须是非空字符串，否则会抛出 `Vtiful\Exception`。

### **int $timestamp**

> Unix 时间戳，扩展会自动换算成 Excel 序列日期。

### **string $title / $message**

> 输入提示与错误提示的标题/正文。

### **toResource()**

> 把当前 `Validation` 对象转换成资源句柄，提供给 `Excel::validation()` 使用。

## 示例

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
    ->inputTitle('请输入分数')
    ->inputMessage('1 到 100 之间的整数')
    ->errorTitle('输入错误')
    ->errorMessage('分数必须是 1~100 的整数');

$excel->fileName('tutorial.xlsx')
    ->header(['Score'])
    ->validation('A1:A10', $validation->toResource())
    ->output();
```
