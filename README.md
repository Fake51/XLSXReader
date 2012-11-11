Introduction
============
This project started as a one file one file library for reading
.xlsx files in PHP. The main reason for creating this was finding
PHPExcel and not being able to use it (memory constraints).

It has since grown to comprehend a namespaced version as well,
alongside the single-file version - this bit contributed by
Alex Kucherenko.

Usage
=====

Single-file version:

<?php

require xlsxreader.php;
$reader = new XLSXReader($zip_file_name);
$sheets = $reader->getSheets();

foreach ($sheets as $sheet) {
    foreach ($sheet->getRowIterator() as $row_index => $row) {
        foreach ($row->getCellIterator() as $cell_index => $cell) {
            // do stuff
        }
    }
}

Author
======
Peter Lind

Contributions
=============
Alex Kucherenko

License
=======
See the COPYRIGHT file
