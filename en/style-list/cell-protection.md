# Cell protection

Control whether a cell is locked and whether its formula is visible. These three flags only take effect once the worksheet itself has been protected through `protection()` — calling `locked()` / `unlocked()` alone does not turn protection on.

## **Function Prototype**

```php
locked(): self

unlocked(): self

hidden(): self
```

### **locked()**

> Mark the cell as locked. This is Excel's default; call it explicitly only when you need to override an earlier `unlocked()`.

### **unlocked()**

> Mark the cell as unlocked, so it remains editable after the worksheet has been protected.

### **hidden()**

> Hide the cell's formula from the formula bar once the worksheet is protected.

## Example

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$editable = (new \Vtiful\Kernel\Format($fileHandle))
    ->unlocked()
    ->toResource();

$secret = (new \Vtiful\Kernel\Format($fileHandle))
    ->locked()
    ->hidden()
    ->toResource();

$fileObject->header(['input', 'formula'])
    ->insertText(1, 0, 'editable', null, $editable)
    ->insertFormula(1, 1, '=A2*2', $secret)
    ->protection()
    ->output();
```
