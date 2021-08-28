--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
    try {
        $config = ['path' => './tests'];

        $fileObject = new \Vtiful\Kernel\Excel($config);

        $fileObject->constMemory('writer_exception.xlsx', 'DemoSheet')
            ->insertText(1, 0, 'viest')
            ->insertText(0, 0, 'viest');
    } catch (\Vtiful\Kernel\Exception $exception) {
          echo $exception->getCode() . PHP_EOL;
          echo $exception->getMessage() . PHP_EOL;
    }
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/writer_exception.xlsx');
?>
--EXPECT--
23
Worksheet row or column index out of range.
