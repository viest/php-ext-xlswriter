# Header and footer

Set the printed header and footer text. The text supports libxlsxwriter's format codes such as `&L` (left), `&C` (center), `&R` (right), `&P` (page number), `&N` (total pages), `&D` (date), `&T` (time), `&"Arial,Bold"` (font), `&14` (size).

## Function Prototype

```php
setHeader(string $value, ?array $options = null): self
setFooter(string $value, ?array $options = null): self
```

### **string $value**

> Header / footer text with optional format codes.

### **array $options**

> Optional. Supported keys:
>
> * `margin`        — distance from the page edge in inches, default `0.3`
> * `image_left`    — left-side image path, used together with the `&G` placeholder in the text
> * `image_center`  — center image path
> * `image_right`   — right-side image path

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
    ->setHeader('&L&"Arial,Bold"&14Sales Report&R&D')
    ->setFooter('&CPage &P of &N', ['margin' => 0.4])
    ->output();
```
