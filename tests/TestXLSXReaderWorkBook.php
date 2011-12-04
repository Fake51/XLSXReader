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

require 'bootstrap.php';

/**
 * tests the workbook class
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class TestXLSXReaderWorkBook extends PHPUnit_Framework_TestCase
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

        $this->setExpectedException('XLSXReaderException');
        $workbook = new XLSXReaderWorkBook($zip, array());
    }
}
