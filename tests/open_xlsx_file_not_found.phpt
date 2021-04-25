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

    $excel->openFile('tutorial_not_found.xlsx');
} catch (Vtiful\Kernel\Exception $exception) {
    var_dump($exception->getMessage());
}
?>
--CLEAN--
<?php
//
?>
--EXPECT--
string(57) "File not found, file path:./tests/tutorial_not_found.xlsx"