<?php

namespace XLSX;
/**
 * .xlsx file reader library
 *
 * PHP Version 5.3+
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../../COPYRIGHT FreeBSD license
 * @version  1.1
 * @link     http://plind.dk/xlsxreader
 */

/**
 * an empty row of cells
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class XLSXReaderFakeCellIterator extends XLSXReaderCellIterator
{
    /**
     * current position of iterator
     *
     * @var int
     */
    protected $position;

    /**
     * public constructor
     *
     * @param null $xml Xml element describing the cell
     * @param XLSXReaderSharedStrings $shared_strings Xml element containing shared strings
     * @param int $cell_start Start position of valid cells in sheet
     * @param int $cell_count Number of cells in sheet
     * @param array $options Array of library options
     *
     * @access public
     * @return void
     */
    public function __construct($xml, $shared_strings, $cell_start, $cell_count, $options)
    {
        $this->cell_count = intval($cell_count);
    }

    /**
     * returns the cell contents at the current index
     *
     * @access public
     * @return mixed
     */
    public function current()
    {
        return null;
    }
}
