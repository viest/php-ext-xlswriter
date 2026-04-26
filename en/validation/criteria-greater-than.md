# Greater than constraint

## Example

The validation type is `integer` and the criteria is `greater than`, with target value `20`. The value of `A1` therefore must be an integer strictly greater than 20.

```php
$config = [
    'path' => './'
];

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_GREATER_THAN)
    ->valueNumber(20);

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial.xlsx')
    ->validation('A1', $validation->toResource())
    ->insertText(0, 0, 21) // Out-of-range; the write succeeds, Excel flags the cell on open.
    ->output();
```

The same pattern applies to the other one-sided criteria — `CRITERIA_LESS_THAN`, `CRITERIA_GREATER_THAN_OR_EQUAL_TO`, `CRITERIA_LESS_THAN_OR_EQUAL_TO`, `CRITERIA_EQUAL_TO`, `CRITERIA_NOT_EQUAL_TO` — each used together with `valueNumber()`, `valueFormula()` or `valueDatetime()`.
