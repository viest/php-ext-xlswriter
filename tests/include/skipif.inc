<?php

/**
 * Skip Disable Reader
 *
 * @return void
 *
 * @author viest
 */
function skip_disable_reader() {
    if (!method_exists('\Vtiful\Kernel\Excel', 'openFile')) {
        print "skip";
    }
}

(function(){
    if (!extension_loaded("xlswriter")) {
        print "skip";
    }
})();

