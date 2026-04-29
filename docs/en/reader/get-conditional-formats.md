# Get conditional formats

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns every `<conditionalFormatting>` block defined on the active worksheet, with the rules inside each block.

## Methods

```php
getConditionalFormats(): ?array
```

## Return value

An array of regions:

* `range` *(string)* — applied range, e.g. `A1:A10`;
* `rules` *(array)* — list of rules; each rule has:
  * `type` *(string)* — rule type, e.g. `cellIs`, `expression`, `top10`, `containsText`;
  * `operator` *(string\|null)* — comparison operator, e.g. `between`, `greaterThan`;
  * `priority` *(int)* — rule priority;
  * `stop_if_true` *(bool)* — stop processing further rules when this one matches;
  * `dxf_id` *(int\|null)* — index into the differential format table;
  * `percent` *(bool)* — `top10`: rank by percent;
  * `bottom` *(bool)* — `top10`: bottom N instead of top N;
  * `rank` *(int\|null)* — `top10`: N value;
  * `text` *(string\|null)* — text used by text rules;
  * `time_period` *(string\|null)* — time-period rule keyword;
  * `formula1` *(string\|null)* — first formula;
  * `formula2` *(string\|null)* — second formula.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$cfs = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getConditionalFormats();

var_export($cfs);
```
