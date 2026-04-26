# Iterate images

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Iterates every embedded image on a worksheet via a callback, exposing the binary content directly.

## Methods

```php
iterateImages(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> Invoked once per image. Return `false` from the callback to stop iteration.

### **string $sheet**

> Worksheet name. When `null`, the currently `openSheet()`-ed worksheet is used.

## Callback payload

The callback receives an associative array:

* `from_row`, `from_col` *(int)* — anchor start cell (0-based);
* `to_row`, `to_col` *(int)* — anchor end cell;
* `mime` *(string)* — MIME type, e.g. `image/png`, `image/jpeg`;
* `data` *(string)* — binary contents;
* `name` *(string)* — file name under `xl/media/`.

## Return value

`true` on successful completion or callback-driven early exit; `false` on error (e.g. the worksheet could not be opened).

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateImages(function (array $img) {
    echo "image at ({$img['from_row']}, {$img['from_col']}), MIME={$img['mime']}\n";
    file_put_contents('./out_' . $img['name'], $img['data']);
}, 'Sheet1');
```
