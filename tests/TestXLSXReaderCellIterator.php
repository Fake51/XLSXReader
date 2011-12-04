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
 * tests the base class
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class TestXLSXReaderCellIterator extends PHPUnit_Framework_TestCase
{
    /**
     * tests the utility method converting dimensions to positions
     *
     * @access public
     * @return void
     */
    public function testConvertCellIndex()
    {
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('A1') === 1);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('a1') === 1);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('b1') === 2);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('z1') === 26);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('AA1') === 27);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('Az1') === 52);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('ba1') === 53);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('zz1') === 702);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('aaa1') === 703);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('aaz1') === 728);
        $this->assertTrue(XLSXReaderCellIterator::convertFromCellPosition('zzz1') === 18278);
    }

    /**
     * tests the utility method converting positions to spreadsheet indexing
     *
     * @access public
     * @return void
     */
    public function testConvertCellToPosition()
    {
        // feature not implemented yet
        return;
        $tests = array(
            'A1' => array(1, 1),
            'B1' => array(1, 2),
            'Z1' => array(1, 26),
            'AA1' => array(1, 27),
            'AZ1' => array(1, 52),
            'BA1' => array(1, 53),
            'ZZ1' => array(1, 702),
            'AAA1' => array(1, 703),
            'AAZ1' => array(1, 728),
            'ZZZ1' => array(1, 18278),
        );

        foreach ($tests as $outcome => $test) {
            $this->assertTrue(XLSXReaderCellIterator::convertToCellPosition($test[0], $test[1]) === $outcome, XLSXReaderCellIterator::convertToCellPosition($test[0], $test[1]) . ' does not match expected ' . $outcome);
        }
    }

    /**
     * tests the counting of cells
     *
     * @access public
     * @return void
     */
    public function testCounts()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();

        foreach ($sheets as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $this->assertTrue($sheet->getCellCount() === count($row));
            }
        }
    }

    /**
     * checks that sheets contain the proper content
     * and amount of cells
     *
     * @access public
     * @return void
     */
    public function testIteration()
    {
        $reader = new XLSXReader('test.xlsx');
        $sheets = $reader->getSheets();

        $data = array(
            array(
                'hans',
                '1',
            ),
            array(
                'wow',
                'hans',
            ),
        );

        foreach ($sheets[0]->getRowIterator() as $row_idx => $row) {
            $iterator = $row->getCellIterator();

            foreach ($iterator as $cell_idx => $cell) {
                $this->assertTrue($data[$row_idx][$cell_idx] == $cell, 'Returned: ' . $cell . ' from row: ' . $row_idx . ' cell: ' . $cell_idx . '. Should be ' . $data[$row_idx][$cell_idx]);
            }
        }
    }

    /**
     * checks that sheets contain the proper content
     * and amount of cells
     *
     * @access public
     * @return void
     */
    public function testIterationBigger()
    {
        $reader = new XLSXReader('test2.xlsx');
        $sheets = $reader->getSheets();

        $data = array(
            array(
                null,
                null,
                null,
                null,
            ),
            array(
                null,
                null,
                null,
                null,
            ),
            array(
                null,
                null,
                1,
                'hans',
            ),
            array(
                null,
                null,
                null,
                'hmm',
            ),
            array(
                null,
                null,
                null,
                'hmm',
            ),
        );

        $i = 0;
        foreach ($sheets[0]->getRowIterator() as $row_idx => $row) {
            $iterator = $row->getCellIterator();

            foreach ($iterator as $cell_idx => $cell) {
                $this->assertTrue($data[$row_idx][$cell_idx] == $cell, 'Returned: ' . $cell . ' from row: ' . $row_idx . ', cell: ' . $cell_idx);
            }
        }
    }

    /**
     * checks that sheets contain the proper content
     * and amount of cells from a minified reader
     *
     * @access public
     * @return void
     */
    public function testIterationMinified()
    {
        $reader = new XLSXReader('test2.xlsx');
        $reader->setOption('read_minified', true);
        $sheets = $reader->getSheets();

        $data = array(
            array(
                1,
                'hans',
            ),
            array(
                null,
                'hmm',
            ),
            array(
                null,
                'hmm',
            ),
        );

        $i = 0;
        foreach ($sheets[0]->getRowIterator() as $row_idx => $row) {
            $iterator = $row->getCellIterator();

            foreach ($iterator as $cell_idx => $cell) {
                $this->assertTrue($data[$row_idx][$cell_idx] == $cell, 'Returned: ' . $cell . ' from row: ' . $row_idx . ', cell: ' . $cell_idx);
            }
        }
    }

    /**
     * checks that sheets contain the proper content
     * and amount of cells from a minified reader
     *
     * @access public
     * @return void
     */
    public function testIterationMinified2()
    {
        $reader = new XLSXReader('test3.xlsx');
        $reader->setOption('read_minified', true);
        $sheets = $reader->getSheets();

        $data = array(
            array(
                1,
                2,
                3,
                null,
                4,
            ),
        );

        $i = 0;
        foreach ($sheets[0]->getRowIterator() as $row_idx => $row) {
            $iterator = $row->getCellIterator();

            foreach ($iterator as $cell_idx => $cell) {
                $this->assertTrue($data[$row_idx][$cell_idx] == $cell, 'Returned: ' . $cell . ' from row: ' . $row_idx . ', cell: ' . $cell_idx);
            }
        }
    }

    /**
     * tests the indexing, to make sure proper spreadsheet
     * indexing works
     *
     * @access public
     * @return void
     */
    public function testSpreadsheetIndexing()
    {
        $reader = new XLSXReader('test.xlsx');
        $reader->setOption('indexing', 'spreadsheet');
        $sheets = $reader->getSheets();

        $data = array(
            '',
            array(
                'A1',
                'B1',
            ),
            array(
                'A2',
                'B2',
            ),
        );

        foreach ($sheets[0]->getRowIterator() as $row_idx => $row) {
            $iterator = $row->getCellIterator();

            $i = 0;
            foreach ($iterator as $cell_idx => $cell) {
                $this->assertTrue($data[$row_idx][$i++] === $cell_idx, "{$data[$row_idx][$i - 1]} != {$cell_idx}");
            }
        }
    }

}
