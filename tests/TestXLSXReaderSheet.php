<?php
/**
 * .xlsx file reader library test
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
 * tests the sheet class
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class TestXLSXReaderSheet extends PHPUnit_Framework_TestCase
{
    /**
     * checks to make sure sheets report
     * properly in terms of name and position
     *
     * @access public
     * @return void
     */
    public function testBaseValues()
    {
        $reader = new XLSXReader('test.xlsx');

        $sheets = $reader->getSheets();
        $this->assertTrue(count($sheets) === 3);
        $this->assertTrue($sheets[0]->getName() === 'hmmm');
        $this->assertTrue($sheets[0]->getPosition() === 1);
        $this->assertTrue($sheets[1]->getName() === 'Sheet2');
        $this->assertTrue($sheets[1]->getPosition() === 2);
        $this->assertTrue($sheets[2]->getName() === 'Sheet3');
        $this->assertTrue($sheets[2]->getPosition() === 3);
    }

    /**
     * checks that the sheet returns a row iterator
     * as it should
     *
     * @access public
     * @return void
     */
    public function testGetIterator()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowIterator() instanceof XLSXReaderRowIterator);
        $this->assertTrue($sheets[1]->getRowIterator() instanceof XLSXReaderRowIterator);
        $this->assertTrue($sheets[2]->getRowIterator() instanceof XLSXReaderRowIterator);
    }

    /**
     * checks that the sheets are aware of their dimensions
     *
     * @access public
     * @return void
     */
    public function testGetRowCount()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowCount() === 2);
        $this->assertTrue($sheets[1]->getRowCount() === 0);
        $this->assertTrue($sheets[2]->getRowCount() === 0);

        $reader = new XLSXReader('test2.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowCount() === 5, $sheets[0]->getRowCount());
        $this->assertTrue($sheets[1]->getRowCount() === 0);
        $this->assertTrue($sheets[2]->getRowCount() === 0);

        $reader = new XLSXReader('test2.xlsx');
        $reader->setOption('read_minified', true);
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowCount() === 3, $sheets[0]->getRowCount());
        $this->assertTrue($sheets[1]->getRowCount() === 0);
        $this->assertTrue($sheets[2]->getRowCount() === 0);
    }

    /**
     * checks that the sheets are aware of their dimensions
     *
     * @access public
     * @return void
     */
    public function testGetCellCount()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getCellCount() === 2);
        $this->assertTrue($sheets[1]->getCellCount() === 0);
        $this->assertTrue($sheets[2]->getCellCount() === 0);

        $reader = new XLSXReader('test2.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getCellCount() === 4);
        $this->assertTrue($sheets[1]->getCellCount() === 0);
        $this->assertTrue($sheets[2]->getCellCount() === 0);

        $reader = new XLSXReader('test2.xlsx');
        $reader->setOption('read_minified', true);
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getCellCount() === 2);
        $this->assertTrue($sheets[1]->getCellCount() === 0);
        $this->assertTrue($sheets[2]->getCellCount() === 0);
    }

    /**
     * tests the start position methods
     *
     * @access public
     * @return void
     */
    public function testGetStart()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowStart() === 1);
        $this->assertTrue($sheets[0]->getCellStart() === 1);
        $this->assertNull($sheets[1]->getRowStart());
        $this->assertNull($sheets[1]->getCellStart());
        $this->assertNull($sheets[2]->getRowStart());
        $this->assertNull($sheets[2]->getCellStart());

        $reader = new XLSXReader('test2.xlsx');
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowStart() === 1);
        $this->assertTrue($sheets[0]->getCellStart() === 1);
        $this->assertNull($sheets[1]->getRowStart());
        $this->assertNull($sheets[1]->getCellStart());
        $this->assertNull($sheets[2]->getRowStart());
        $this->assertNull($sheets[2]->getCellStart());

        $reader = new XLSXReader('test2.xlsx');
        $reader->setOption('read_minified', true);
        $sheets = $reader->getSheets();
        $this->assertTrue($sheets[0]->getRowStart() === 3);
        $this->assertTrue($sheets[0]->getCellStart() === 3);
        $this->assertNull($sheets[1]->getRowStart());
        $this->assertNull($sheets[1]->getCellStart());
        $this->assertNull($sheets[2]->getRowStart());
        $this->assertNull($sheets[2]->getCellStart());
    }
}
