--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
try {
    $excel = new \Vtiful\Kernel\Excel([
        'path' => './',
    ]);

    $fileObject = $excel->constMemory('const_memory_index_out_range.xlsx');
    $fileHandle = $fileObject->getHandle();

    $format    = new \Vtiful\Kernel\Format($fileHandle);
    $boldStyle = $format->bold()->toResource();

    $fileObject->header(['name', 'age'])
        ->data([['viest', 21]])
        ->mergeCells('A1:C1', 'aaaa')
        ->output();
} catch (\Vtiful\Kernel\Exception $exception) {
    echo $exception->getCode() . PHP_EOL;
    echo $exception->getMessage() . PHP_EOL;
}

try {
    $excel = new \Vtiful\Kernel\Excel([
        'path' => './',
    ]);

    $fileObject = $excel->constMemory('const_memory_index_out_range.xlsx');
    $fileHandle = $fileObject->getHandle();

    $format    = new \Vtiful\Kernel\Format($fileHandle);
    $boldStyle = $format->bold()->toResource();

    $fileObject->header(['name', 'age'])
        ->data([['viest', 21]])
        ->setRow('A1', 200)
        ->output();
} catch (\Vtiful\Kernel\Exception $exception) {
    echo $exception->getCode() . PHP_EOL;
    echo $exception->getMessage() . PHP_EOL;
}

try {
    $excel = new \Vtiful\Kernel\Excel([
        'path' => './',
    ]);

    $fileObject = $excel->constMemory('const_memory_index_out_range.xlsx');
    $fileHandle = $fileObject->getHandle();

    $format    = new \Vtiful\Kernel\Format($fileHandle);
    $boldStyle = $format->bold()->toResource();

    $fileObject->header(['name', 'age'])
        ->data([['viest', 21]])
        ->autoFilter('A1:C1')
        ->output();
} catch (\Vtiful\Kernel\Exception $exception) {
    echo $exception->getCode() . PHP_EOL;
    echo $exception->getMessage() . PHP_EOL;
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/const_memory_index_out_range.xlsx');
?>
--EXPECT--
170
In const memory mode, you cannot modify the placed cells
170
In const memory mode, you cannot modify the placed cells
170
In const memory mode, you cannot modify the placed cells