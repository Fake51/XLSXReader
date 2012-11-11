<?php
/**
 * resource test to see memory and time consumption
 *
 * PHP Version 5.3+
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */

if (defined('TESTING')) {
    return;
}

require __DIR__ . '/../src/xlsxreader.php';

$start_time = microtime(true);
$start_mem  = memory_get_usage(true);

$reader = new XLSXReader('test3.xlsx');
$reader->setOption('read_minified', true);
$default_reader_mem = memory_get_usage(true);
foreach ($reader->getSheets() as $sheet) {
    foreach ($sheet->getRowIterator() as $idx => $row) {
        echo "Index: " . $idx . PHP_EOL;
        foreach ($row->getCellIterator() as $cell_idx => $cell) {
            echo "- " . $cell_idx . ': ' . $cell . PHP_EOL;
        }
    }
}

$end_memory = memory_get_usage(true);
echo "Start memory: " . ($start_mem / 1024 / 1024) . 'MB' . PHP_EOL;
echo "Init memory: " . ($default_reader_mem  / 1024 / 1024) . 'MB'. PHP_EOL;
echo "Finished memory: " . ($end_memory  / 1024 / 1024) . 'MB'. PHP_EOL;
echo "Time used: " . ((microtime(true) - $start_time) / 1) . ' seconds' . PHP_EOL;
