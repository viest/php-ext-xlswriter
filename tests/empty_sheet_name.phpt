--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

/* fileName(file, '') -> throw */
try {
    (new \Vtiful\Kernel\Excel($config))->fileName('empty_sheet_name.xlsx', '');
    echo "fileName: no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "fileName: " . $e->getCode() . " " . $e->getMessage() . "\n";
}

/* addSheet('') -> throw, original sheet still usable */
$excel = (new \Vtiful\Kernel\Excel($config))->fileName('empty_sheet_name.xlsx', 'S1');
try {
    $excel->addSheet('');
    echo "addSheet: no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "addSheet: " . $e->getCode() . " " . $e->getMessage() . "\n";
}
$excel->insertText(0, 0, 'x')->output();

/* constMemory(file, '') -> throw */
try {
    (new \Vtiful\Kernel\Excel($config))->constMemory('empty_sheet_name_cm.xlsx', '');
    echo "constMemory: no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "constMemory: " . $e->getCode() . " " . $e->getMessage() . "\n";
}

/* null still auto-generates the default sheet name (regression guard). */
$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('empty_sheet_name_null.xlsx', null)
    ->insertText(0, 0, 'x')
    ->output();
echo "nullAccepted: " . var_export(is_file($path), true) . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/empty_sheet_name.xlsx');
@unlink(__DIR__ . '/empty_sheet_name_null.xlsx');
?>
--EXPECT--
fileName: 131 Sheet name must not be an empty string; pass null to auto-generate one
addSheet: 131 Sheet name must not be an empty string; pass null to auto-generate one
constMemory: 131 Sheet name must not be an empty string; pass null to auto-generate one
nullAccepted: true
