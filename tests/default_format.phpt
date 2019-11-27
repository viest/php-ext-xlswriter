--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx');

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
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
string(21) "./tests/tutorial.xlsx"
