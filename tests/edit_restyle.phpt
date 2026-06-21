--TEST--
Applying a Format in edit mode restyles a cell (font/fill/alignment) round-trip
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_restyle_src.xlsx', 'S')
    ->insertText(0, 0, 'plain')
    ->output();

$x = (new \Vtiful\Kernel\Excel($config))->openFile('edit_restyle_src.xlsx')->openSheet('S');
$fmt = (new \Vtiful\Kernel\Format($x->getHandle()))
    ->bold()
    ->fontColor(0xFF0000)
    ->background(0x00FF00)
    ->align(\Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER)
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->toResource();
$x->insertText(2, 0, 'styled', null, $fmt)->output('edit_restyle_out.xlsx');

// Read the styled cell's style id back, then resolve the full style.
$rd = (new \Vtiful\Kernel\Excel($config))->openFile('edit_restyle_out.xlsx')->openSheet('S');
$sid = null;
while (($row = $rd->nextRowWithFormula()) !== null) {
    foreach ($row as $cell) {
        if (is_array($cell) && !empty($cell['style_id'])) {
            $sid = $cell['style_id'];
        }
    }
}

$style = (new \Vtiful\Kernel\Excel($config))->openFile('edit_restyle_out.xlsx')->openSheet('S')->getStyleFormat($sid);

var_dump($style['font']['bold']);
var_dump($style['font']['color']);
var_dump($style['fill']['pattern_type']);
var_dump($style['fill']['fg_color']);
var_dump($style['alignment']['horizontal']);
var_dump($style['border']['left']['style']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_restyle_src.xlsx');
@unlink(__DIR__ . '/edit_restyle_out.xlsx');
?>
--EXPECT--
bool(true)
string(8) "FFFF0000"
string(5) "solid"
string(8) "FF00FF00"
string(6) "center"
string(4) "thin"
