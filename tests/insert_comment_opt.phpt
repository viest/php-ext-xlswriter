--TEST--
insertCommentOpt: writes file successfully with full options
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('insert_comment_opt.xlsx')
    ->insertText(0, 0, 'cell')
    ->insertCommentOpt(0, 0, 'Reviewed', [
        'author'    => 'QA',
        'visible'   => 1,
        'color'     => 0xFFFFCC,
        'font_name' => 'Calibri',
        'font_size' => 12,
        'x_offset'  => 5,
        'y_offset'  => 5,
        'x_scale'   => 1.5,
        'y_scale'   => 1.5,
    ])
    ->output();

var_dump(is_file($path));
var_dump(filesize($path) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_comment_opt.xlsx');
?>
--EXPECT--
bool(true)
bool(true)
