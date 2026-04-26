--TEST--
defineName: writer-emitted defined names readable via getDefinedNames
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('define_name_round_trip.xlsx', 'Sales')
    ->defineName('Exchange_rate', '=0.96')
    ->defineName('LocalRange', '=Sales!$A$1:$B$2', 'Sales')
    ->insertText(0, 0, 'x')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('define_name_round_trip.xlsx');

foreach ($reader->getDefinedNames() as $dn) {
    printf("%s = %s scope=%s hidden=%s\n",
        $dn['name'], $dn['formula'],
        $dn['scope'] ?? '(workbook)',
        var_export($dn['hidden'], true));
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/define_name_round_trip.xlsx');
?>
--EXPECT--
Exchange_rate = 0.96 scope=(workbook) hidden=false
LocalRange = Sales!$A$1:$B$2 scope=Sales hidden=false
