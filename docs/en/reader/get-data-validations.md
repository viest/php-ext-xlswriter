# Get data validations

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns every `<dataValidation>` element on the active worksheet — the read-side counterpart of the `Validation` writer class.

## Methods

```php
getDataValidations(): ?array
```

## Return value

An array of validation entries, each with:

* `type` *(string)* — `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`;
* `operator` *(string\|null)* — `between`, `equal`, `greaterThan`, ...;
* `error_style` *(string\|null)* — `stop`, `warning`, `information`;
* `formula1` *(string\|null)* — first formula / list source;
* `formula2` *(string\|null)* — second formula;
* `prompt` *(string\|null)* — input prompt body;
* `prompt_title` *(string\|null)* — input prompt title;
* `error` *(string\|null)* — error message body;
* `error_title` *(string\|null)* — error message title;
* `sqref` *(string)* — applied range, e.g. `A1:A10`;
* `allow_blank` *(bool)* — whether blanks are allowed;
* `show_drop_down` *(bool)* — whether the drop-down arrow is shown (note: inverted in XML);
* `show_input_message` *(bool)* — whether to show the input prompt;
* `show_error_message` *(bool)* — whether to show the error message.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$dvs = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getDataValidations();

print_r($dvs);
```
