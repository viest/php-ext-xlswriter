--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* #535: header() after data() used to silently overwrite the first data
 * row (data() bumped write_line to 4, header() wrote to row 0 anyway).
 * Now it throws with code 132. */
$excel = (new \Vtiful\Kernel\Excel($config))
    ->fileName('header_after_data.xlsx', 'sheet1')
    ->data([
        ['Rent', 1000],
        ['Gas',  100],
    ]);

try {
    $excel->header(['Item', 'Cost']);
    echo "no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "code=" . $e->getCode() . "\n";
    echo "hasOrderHint=" . var_export(strpos($e->getMessage(), 'before data') !== false, true) . "\n";
}

/* The documented order — header() then data() — still works. */
$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('header_after_data_ok.xlsx', 'sheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Rent', 1000],
        ['Gas',  100],
    ])
    ->output();
$rows = (new \Vtiful\Kernel\Excel($config))->openFile('header_after_data_ok.xlsx')->openSheet()->getSheetData();
echo "rows: " . count($rows) . "\n";
var_dump($rows);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/header_after_data.xlsx');
@unlink(__DIR__ . '/header_after_data_ok.xlsx');
?>
--EXPECT--
code=132
hasOrderHint=true
rows: 3
array(3) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "Item"
    [1]=>
    string(4) "Cost"
  }
  [1]=>
  array(2) {
    [0]=>
    string(4) "Rent"
    [1]=>
    int(1000)
  }
  [2]=>
  array(2) {
    [0]=>
    string(3) "Gas"
    [1]=>
    int(100)
  }
}
