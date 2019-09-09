# Aligning

### **Function Prototype**

```php
align(resource $resourchHandle, Format::const ...$style): \Vtiful\Kernel\Format
```

### Example

```php
$format     = new \Vtiful\Kernel\Format($fileHandle);
$alignStyle = $format
    ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
    ->toResource();
```

### **Style**

```php
Format::FORMAT_ALIGN_LEFT;                 // horizontally left aligned
Format::FORMAT_ALIGN_CENTER;               // Align in horizontal drama
Format::FORMAT_ALIGN_RIGHT;                // horizontal right alignment
Format::FORMAT_ALIGN_FILL;                 // horizontal fill alignment
Format::FORMAT_ALIGN_JUSTIFY;              // Horizontal justification
Format::FORMAT_ALIGN_CENTER_ACROSS;        // Horizontal center alignment
Format::FORMAT_ALIGN_DISTRIBUTED;          // Disperse alignment
Format::FORMAT_ALIGN_VERTICAL_TOP;         // top vertical alignment
Format::FORMAT_ALIGN_VERTICAL_BOTTOM;      // bottom vertical alignment
Format::FORMAT_ALIGN_VERTICAL_CENTER;      // Align in vertical drama
Format::FORMAT_ALIGN_VERTICAL_JUSTIFY;     // Vertical justification
Format::FORMAT_ALIGN_VERTICAL_DISTRIBUTED; // Vertically dispersing alignment
```