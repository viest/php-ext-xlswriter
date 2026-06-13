--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Short list — accepted, file produced. */
$short = ['Apple', 'Banana', 'Cherry'];
$v = (new \Vtiful\Kernel\Validation())
    ->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
    ->valueList($short);
$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('validation_list_short.xlsx')
    ->validation('A1', $v->toResource())
    ->insertText(0, 0, 'pick')
    ->output();
var_dump(is_file($path));

/* Long list — projected CSV exceeds 255 chars, must throw cleanly instead
 * of crashing libxlsxwriter's fixed-size buffer (#486 / #530 / #546). */
$long = [];
for ($i = 1; $i <= 30; $i++) {
    $long[] = sprintf('option-with-some-padding-%02d', $i); /* 27 chars each */
}
try {
    (new \Vtiful\Kernel\Validation())
        ->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
        ->valueList($long);
    echo "no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "code=" . $e->getCode() . "\n";
    echo "msgHasLimit=" . var_export(strpos($e->getMessage(), '255') !== false, true) . "\n";
}

/* Chinese / multi-byte stress — the limit is in UTF-8 codepoints, not
 * bytes, so a list that ranks under 255 bytes naively but over 255 chars
 * should also throw and NOT trip the libxlsxwriter buffer overflow.
 * 16 items × 16 cn chars = 256 codepoints + commas + quotes > 255. */
$cn = [];
for ($i = 0; $i < 16; $i++) {
    $cn[] = '数据列表中的一个较长选项编号' . sprintf('%02d', $i); /* 14 cn + 2 digit = 16 chars */
}
try {
    (new \Vtiful\Kernel\Validation())
        ->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
        ->valueList($cn);
    echo "cn: no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "cn.code=" . $e->getCode() . "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/validation_list_short.xlsx');
?>
--EXPECT--
bool(true)
code=302
msgHasLimit=true
cn.code=302
