--TEST--
getPageSetup: round-trip writer's setMargins / setHeader / setFooter / setPaper
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_page_setup.xlsx')
    ->setHeader('&CHello')
    ->setFooter('&D')
    ->setMargins(0.5, 0.6, 0.7, 0.8)
    ->setPaper(9)
    ->insertText(0, 0, 'x')
    ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_page_setup.xlsx')
    ->openSheet();
$ps = $r->getPageSetup();
printf("paper=%d orientation=%s\n", $ps['paper_size'], $ps['orientation']);
printf("odd_header=%s odd_footer=%s\n", $ps['odd_header'], $ps['odd_footer']);
printf("ml=%.2f mr=%.2f mt=%.2f mb=%.2f\n",
    $ps['margins']['left'],
    $ps['margins']['right'],
    $ps['margins']['top'],
    $ps['margins']['bottom']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_page_setup.xlsx');
?>
--EXPECTF--
paper=9 orientation=portrait
odd_header=&CHello odd_footer=&D
ml=%f mr=%f mt=%f mb=%f
