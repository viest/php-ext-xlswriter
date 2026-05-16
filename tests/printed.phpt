--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
try {
    $config = ['path' => './tests'];
    $excel  = new \Vtiful\Kernel\Excel($config);

    $excel->setPortrait();
} catch (\Exception $exception) {
    var_dump($exception->getCode());
    var_dump($exception->getMessage());
}

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_portrait.xlsx', 'sheet1')
    ->setPortrait()
    ->output();

var_dump($excel);

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_landscape.xlsx', 'sheet1')
    ->setLandscape()
    ->output();

var_dump($excel);

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_scale.xlsx', 'sheet1')
    ->setPrintScale(180)
    ->output();

var_dump($excel);

/* Round-trip: each printed-* writer setting is recoverable via getPageSetup. */
foreach (['printed_landscape' => 'landscape', 'printed_scale' => null] as $name => $expectedOrient) {
    $ps = (new \Vtiful\Kernel\Excel($config))->openFile($name . '.xlsx')->openSheet()->getPageSetup();
    if ($name === 'printed_landscape') {
        echo "landscape.orientation: " . $ps['orientation'] . "\n";
    } else {
        echo "scale.scale: " . $ps['scale'] . "\n";
    }
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/printed_portrait.xlsx');
@unlink(__DIR__ . '/printed_landscape.xlsx');
@unlink(__DIR__ . '/printed_scale.xlsx');
?>
--EXPECTF--
int(130)
string(51) "Please create a file first, use the filename method"
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(29) "./tests/printed_portrait.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(30) "./tests/printed_landscape.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(26) "./tests/printed_scale.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
landscape.orientation: landscape
scale.scale: 180
