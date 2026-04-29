# Indent

Add a left indent to the text inside a cell.

## **Function Prototype**

```php
indent(int $level): self
```

### **int $level**

> Indent level, in the range `0..15`. Each level is roughly the width of three spaces.

## Example

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$level1 = (new \Vtiful\Kernel\Format($fileHandle))
    ->indent(1)
    ->toResource();

$level3 = (new \Vtiful\Kernel\Format($fileHandle))
    ->indent(3)
    ->toResource();

$fileObject->header(['title'])
    ->insertText(1, 0, 'level 1', null, $level1)
    ->insertText(2, 0, 'level 3', null, $level3)
    ->output();
```
