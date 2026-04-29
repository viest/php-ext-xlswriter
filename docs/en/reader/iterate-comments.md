# Iterate comments

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Iterates every comment on a worksheet via a callback, including both legacy and threaded comments.

## Methods

```php
iterateComments(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> Invoked once per comment. Return `false` from the callback to stop iteration.

### **string $sheet**

> Worksheet name. When `null`, the currently `openSheet()`-ed worksheet is used.

## Callback payload

The callback receives an associative array:

* `row` *(int)* — row index (0-based);
* `col` *(int)* — column index (0-based);
* `text` *(string)* — comment body;
* `author` *(string\|null)* — author;
* `visible` *(bool)* — whether the comment is permanently visible;
* `threaded` *(bool)* — whether this is a threaded comment;
* `parent_id` *(string\|null)* — parent id for threaded replies.

## Return value

`true` on successful completion or callback-driven early exit; `false` on error.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateComments(function (array $c) {
    echo "[{$c['author']}] R{$c['row']}C{$c['col']}: {$c['text']}\n";
}, 'Sheet1');
```
