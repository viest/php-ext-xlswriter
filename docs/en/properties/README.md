# Workbook properties

Write document properties into the xlsx. They appear in Excel's "File → Info → Properties" panel. There are two kinds:

* **Standard properties** — fixed fields defined by OOXML (title, author, company, etc.), written in one call to `Excel::setProperties`.
* **Custom properties** — user-defined name / value pairs, written one at a time via `Excel::setCustomProperty`.

## Function Prototype

```php
setProperties(array $properties): self
setCustomProperty(string $name, mixed $value, ?string $type = null): self
```

Topics:

* [Standard properties](standard.md)
* [Custom properties](custom.md)
