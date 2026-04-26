# Data validation

Data validation restricts what users can type into a cell — for example only allowing integers, decimals, dates, or values picked from a list. When a user enters something that breaks the rule, Excel surfaces a warning when the file is opened. Note that the extension itself does **not** check values at write time; it just persists the rule in the workbook.

There are two steps:

1. Build a rule with `Vtiful\Kernel\Validation`.
2. Apply it to a range with `Excel::validation(string $range, $validationHandle)`. `$validationHandle` is the resource returned by `Validation::toResource()`.

## Function prototype

```
\Vtiful\Kernel\Excel::validation(string $range, resource $validationHandle): self
```

### **string $range**

> The cell range to apply the rule to, in A1 notation, for example `A1`, `A1:A10`, `B2:D8`.

### **resource $validationHandle**

> The resource handle returned by `Validation::toResource()`.

## Sub-sections

* [Drop-down list](type-list.md)
* [Range constraint](criteria-between.md)
* [Greater than constraint](criteria-greater-than.md)
* [Validation API reference](api-reference.md)

## Example

The most common case — restrict `A1:A10` to integers between 1 and 100.

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_BETWEEN)
    ->minimumNumber(1)
    ->maximumNumber(100);

$excel->fileName('tutorial.xlsx')
    ->header(['Score'])
    ->validation('A1:A10', $validation->toResource())
    ->output();
```
