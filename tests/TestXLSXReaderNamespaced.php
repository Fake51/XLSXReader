<?php
use \XLSX;

/**
 * .xlsx file reader library test
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
 * tests the base class
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class TestXLSXReaderNamespaced extends PHPUnit_Framework_TestCase
{
    /**
     * tests good and bad construction
     *
     * @access public
     * @return void
     */
    public function testConstruct()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');

        $this->setExpectedException('\\XLSX\\XLSXReaderException');
        $reader = new XLSX\XLSXReader('');
    }

    /**
     * test getting the sheets from a
     * spreadsheet file
     *
     * @access public
     * @return void
     */
    public function testGetSheets()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');

        $i = 0;
        foreach ($reader->getSheets() as $sheet) {
            $this->assertTrue($sheet instanceof XLSX\XLSXReaderSheet);
            $i++;
        }

        $this->assertTrue($i === 3);
    }

    /**
     * test getting the work book
     *
     * @access public
     * @return void
     */
    public function testGetWorkBook()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');

        $this->assertTrue($reader->getWorkBook() instanceof XLSX\XLSXReaderWorkBook);

        $reader_sheets = $reader->getSheets();
        $wb_sheets     = $reader->getWorkBook()->getSheets();
        $this->assertTrue($reader_sheets === $wb_sheets);
    }

    /**
     * test getting the shared strings
     *
     * @access public
     * @return void
     */
    public function testGetSharedStrings()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');

        $this->assertTrue($reader->getSharedStrings() instanceof XLSX\XLSXReaderSharedStrings);
    }

    /**
     * tests the various option related methods
     *
     * @access public
     * @return void
     */
    public function testOptions()
    {
        $reader  = new XLSX\XLSXReader('test.xlsx');
        $options = $reader->getAvailableOptions();
        $this->assertTrue(is_array($options));
        $this->assertTrue(in_array('read_minified', $options));

        $options = $reader->getOptions();
        $this->assertTrue(is_array($options));
        $this->assertTrue(isset($options['read_minified']));

        $this->assertTrue($reader->setOption('read_minified', true) === $reader);

        $options = $reader->getOptions();
        $this->assertTrue(is_array($options));
        $this->assertTrue($options['read_minified'] === true);
    }
}
