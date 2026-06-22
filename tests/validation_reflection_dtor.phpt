--TEST--
Validation destructor survives an unconstructed object (NULL ptr guard)
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* An object created without running __construct has ptr.validation == NULL.
 * The destructor used to dereference value_list before the NULL check and
 * crash; it must now no-op safely. */
$r = new ReflectionClass('Vtiful\Kernel\Validation');
$v = $r->newInstanceWithoutConstructor();
unset($v);
echo "unconstructed dtor ok\n";

/* A normally-constructed, fully-populated Validation also destructs cleanly. */
$v2 = new \Vtiful\Kernel\Validation();
$v2->validationType(\Vtiful\Kernel\Validation::TYPE_CUSTOM_FORMULA)
   ->valueFormula('=A1>0')
   ->inputTitle('t')
   ->errorMessage('e');
unset($v2);
echo "constructed dtor ok\n";
?>
--EXPECT--
unconstructed dtor ok
constructed dtor ok
