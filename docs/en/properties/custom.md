# Custom properties

Write user-defined workbook properties. Each call writes one name / value pair.

## Function Prototype

```php
setCustomProperty(string $name, mixed $value, ?string $type = null): self
```

### **string $name**

> Property name.

### **mixed $value**

> Property value.

### **string $type**

> Property type. Allowed values:
>
> * `string`   — string
> * `number`   — integer or floating-point number
> * `boolean`  — boolean
> * `datetime` — date / time; `$value` must be a **Unix timestamp**
>
> When omitted the type is inferred from PHP's type of `$value` (string / int / double / bool). `datetime` must always be specified explicitly.
>
> An unsupported type throws an exception (error code `222`).

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setCustomProperty('Department',   'Sales')          // inferred as string
    ->setCustomProperty('Confidential', true)             // inferred as boolean
    ->setCustomProperty('Revision',     3)                // inferred as number
    ->setCustomProperty('Reviewed',     time(), 'datetime') // datetime must be explicit
    ->output();
```
