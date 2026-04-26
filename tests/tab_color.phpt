--TEST--
setTabColor: file produced successfully (visual property, no read-back)
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('tab_color.xlsx', 'Q1')
    ->setTabColor(0xFF6600)
    ->addSheet('Q2')
    ->setTabColor(\Vtiful\Kernel\Format::COLOR_GREEN)
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file($path));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tab_color.xlsx');
?>
--EXPECT--
bool(true)
