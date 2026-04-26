# Drop-down list

Use `TYPE_LIST` together with `valueList()` to render an in-cell drop-down whose items come from a PHP array of strings.

## Example

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
    ->valueList(['wjx', 'viest']);

$excel->fileName('tutorial.xlsx')
    ->validation('A1', $validation->toResource())
    ->output();
```

The drop-down items can also be sourced from a worksheet range with `TYPE_LIST_FORMULA` plus `valueFormula('=$E$1:$E$5')`.
