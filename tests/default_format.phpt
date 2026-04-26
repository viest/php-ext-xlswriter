--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('default_format.xlsx');

$format        = new \Vtiful\Kernel\Format($excel->getHandle());
$colorOneStyle = $format
    ->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)
    ->border(\Vtiful\Kernel\Format::BORDER_DASH_DOT)
    ->toResource();

$format        = new \Vtiful\Kernel\Format($excel->getHandle());
$colorTwoStyle = $format
    ->fontColor(\Vtiful\Kernel\Format::COLOR_GREEN)
    ->toResource();

$filePath = $excel
    // Apply the first style as the default
    ->defaultFormat($colorOneStyle)
    ->header(['hello', 'xlswriter'])
    // Apply the second style as the default style
    ->defaultFormat($colorTwoStyle)
    ->data([
        ['hello', 'xlswriter'],
    ])
    ->output();

var_dump($filePath);

/* Round-trip: written cells have a non-default style after read-back. */
$v_     = new \Vtiful\Kernel\Excel($config);
$v_->openFile('default_format.xlsx')->openSheet();
$styled_ = false;
while (($r_ = $v_->nextRowWithFormula()) !== null) {
    foreach ($r_ as $c_) { if ($c_['style_id'] > 0) { $styled_ = true; break 2; } }
}
var_dump($styled_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/default_format.xlsx');
?>
--EXPECT--
string(27) "./tests/default_format.xlsx"
bool(true)
