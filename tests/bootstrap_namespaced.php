<?php
/**
 * .xlsx file reader library
 *
 * PHP Version 5.3+
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */

if (!defined('TESTING_NAMESPACED')) {
    define('TESTING_NAMESPACED', true);
    define('XLSXBASEDIR', __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . 'XLSX' . DIRECTORY_SEPARATOR);
    include XLSXBASEDIR . 'XLSXReaderCellIterator.php';
    include XLSXBASEDIR . 'XLSXReaderException.php';
    include XLSXBASEDIR . 'XLSXReaderFakeCellIterator.php';
    include XLSXBASEDIR . 'XLSXReader.php';
    include XLSXBASEDIR . 'XLSXReaderRowIterator.php';
    include XLSXBASEDIR . 'XLSXReaderSharedStrings.php';
    include XLSXBASEDIR . 'XLSXReaderSheet.php';
    include XLSXBASEDIR . 'XLSXReaderWorkBook.php';
}
