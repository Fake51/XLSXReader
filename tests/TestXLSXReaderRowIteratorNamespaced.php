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
class TestXLSXReaderRowIteratorNamespaced extends PHPUnit_Framework_TestCase
{
    /**
     * makes sure that the row iterator returns correct
     * number of rows and iterates over correct amount
     *
     * @access public
     * @return void
     */
    public function testIteration()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();

        foreach ($sheets as $sheet) {
            $this->assertTrue($sheet->getRowCount() === count($sheet->getRowIterator()));

            $count = count($sheet->getRowIterator());
            foreach ($sheet->getRowIterator() as $row) {
                $count--;
            }

            $this->assertTrue($count === 0);
        }
    }

    /**
     * makes sure that the row iterator returns correct
     * number of rows and iterates over correct amount
     *
     * @access public
     * @return void
     */
    public function testIndexing()
    {
        $reader = new XLSX\XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();

        $i = 0;
        foreach ($sheets[0]->getRowIterator() as $idx => $row) {
            $this->assertTrue($idx === $i++);
        }

        $reader = new XLSX\XLSXReader('test.xlsx');
        $reader->setOption('indexing', 'spreadsheet');
        $sheets = $reader->getSheets();

        $i = 1;
        foreach ($sheets[0]->getRowIterator() as $idx => $row) {
            $this->assertTrue($idx === $i++);
        }
    }
}
