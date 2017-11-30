<?php
/**
 * Xls Test
 *
 * @author Janson
 * @create 2017-11-28
 */
require __DIR__ . '/../autoload.php';

$start = microtime(true);
$memory = memory_get_usage();

$reader = EC\PHPExcel\Excel::load('files/03.xls', function(EC\PHPExcel\Reader\Xls $reader) {
    //$reader->setRowLimit(1);
    //$reader->setColumnLimit(10);

    //$reader->setSheetIndex(1);
});

$reader->seek(50);

$reader->seek(1);
$current = $reader->current();

$count = $reader->count();
//$sheets = $reader->sheets();

$time = microtime(true) - $start;
$use = memory_get_usage() - $memory;

var_dump($current, $count, $time, $use/1024/1024);
