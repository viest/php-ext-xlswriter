--TEST--
getHyperlinks round-trip: writer insertUrl -> reader returns url + range
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName('open_xlsx_get_hyperlinks.xlsx')
    ->insertUrl(0, 0, 'https://example.com/a')
    ->insertUrl(2, 1, 'https://example.com/b')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_hyperlinks.xlsx')
    ->openSheet();

$rows = $reader->getHyperlinks();
foreach ($rows as $r) {
    printf("(%d,%d)-(%d,%d) %s\n",
        $r['first_row'], $r['first_col'],
        $r['last_row'],  $r['last_col'],
        $r['url']);
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_hyperlinks.xlsx');
?>
--EXPECT--
(1,1)-(1,1) https://example.com/a
(3,2)-(3,2) https://example.com/b
