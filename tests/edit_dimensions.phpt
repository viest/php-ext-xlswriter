--TEST--
setColumn()/setRow() work in edit mode (column widths and row heights)
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_dim_src.xlsx', 'S')
    ->insertText(0, 0, 'a')
    ->insertText(0, 1, 'b')
    ->output();

(new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_dim_src.xlsx')
    ->openSheet('S')
    ->setColumn('A:C', 25.0)
    ->setRow('A1', 30.0)
    ->output('edit_dim_out.xlsx');

var_dump((new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_dim_out.xlsx')->openSheet('S')->getColumnOptions('A')['width']);
var_dump((new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_dim_out.xlsx')->openSheet('S')->getRowOptions(0)['height']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_dim_src.xlsx');
@unlink(__DIR__ . '/edit_dim_out.xlsx');
?>
--EXPECT--
float(25)
float(30)
