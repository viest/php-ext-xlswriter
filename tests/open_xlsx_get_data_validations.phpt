--TEST--
getDataValidations: round-trip writer.validation() + fixture variants
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Round-trip: writer -> reader. */
$v = (new \Vtiful\Kernel\Validation())
    ->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_BETWEEN)
    ->minimumNumber(1)
    ->maximumNumber(10)
    ->inputTitle('Range')
    ->inputMessage('1..10')
    ->errorTitle('Bad')
    ->errorMessage('out of range');

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_data_validations.xlsx')
    ->insertText(0, 0, 'x')
    ->validation('A1:A5', $v->toResource())
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_data_validations.xlsx')
    ->openSheet();
$dv = $reader->getDataValidations()[0];
printf("type=%s sqref=%s formula1=%s formula2=%s prompt='%s' error_title='%s'\n",
    $dv['type'], $dv['sqref'], $dv['formula1'], $dv['formula2'],
    $dv['prompt'], $dv['error_title']);

/* Fixture: list-type and rich props. */
$reader2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase3.xlsx')
    ->openSheet();
$dvs = $reader2->getDataValidations();
echo "fixture count: " . count($dvs) . "\n";
echo "[0] " . $dvs[0]['type'] . " " . $dvs[0]['operator'] . " " . $dvs[0]['sqref'] . "\n";
echo "[1] " . $dvs[1]['type'] . " sqref=" . $dvs[1]['sqref'] .
     " formula1=" . $dvs[1]['formula1'] . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_data_validations.xlsx');
?>
--EXPECT--
type=whole sqref=A1:A5 formula1=1 formula2=10 prompt='1..10' error_title='Bad'
fixture count: 2
[0] whole between D2:D8
[1] list sqref=C1:C2 formula1="alpha,beta,gamma"
