<?php
use \XLSX;
/**
 * .xlsx file reader library
 *
 * PHP Version 5.3+
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */

require 'bootstrap_namespaced.php';

/**
 * tests the workbook class
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class TestXLSXReaderWorkBookNamespaced extends PHPUnit_Framework_TestCase
{
    /**
     * checks that workbook class doesn't accept
     * bad params for the constructor
     *
     * @access public
     * @return void
     */
    public function testBadConstruct()
    {
        $zip = new ZipArchive();

        $this->setExpectedException('\XLSX\XLSXReaderException');
        $workbook = new XLSX\XLSXReaderWorkBook($zip, array());
    }
}
