<?php
/**
 * Csv test
 *
 * @author Janson
 * @create 2017-11-25
 */
require __DIR__ . '/../autoload.php';

$start = microtime(true);
$memory = memory_get_usage();

$reader = EC\PHPExcel\Excel::load('files/03.csv', function(EC\PHPExcel\Reader\Csv $reader) {
    $reader->setRowLimit(10);
    $reader->setColumnLimit(10);

    $reader->setInputEncoding('UTF-8');
    $reader->setDelimiter("\t");
});


$reader->seek(50);

$reader->seek(2);
$current = $reader->current();

$count = $reader->count();


$time = microtime(true) - $start;
$use = memory_get_usage() - $memory;
var_dump($current, $time, $use/1024/1024);
