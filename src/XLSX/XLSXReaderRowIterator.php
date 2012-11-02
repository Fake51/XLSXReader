<?php

namespace XLSX;

use Iterator;
use SimpleXMLElement;
use Countable;

/**
* row iterator - allows for iterating over rows
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
*/
class XLSXReaderRowIterator implements Iterator, countable
{
    /**
* xml doc to iterate over
*
* @var SimpleXMLElement
*/
    protected $xml;

    /**
* current position of iterator
*
* @var int
*/
    protected $position;

    /**
* number of rows in sheet
*
* @var int
*/
    protected $row_count = 0;

    /**
* number of cells in sheet
*
* @var int
*/
    protected $cell_count = 0;

    /**
* start of rows in sheet
*
* @var int
*/
    protected $row_start;

    /**
* start of cells in sheet
*
* @var int
*/
    protected $cell_start;

    /**
* shared strings xml object
*
* @var SimpleXMLElement
*/
    protected $shared_strings;

    /**
* array of options for the library
*
* @var array
*/
    protected $options;

    /**
* keeps track of which rows are valid
*
* @var array
*/
    protected $valid_positions;

    /**
* integer added to key for
* array indexing
*
* @var int
*/
    protected $key_addition = 0;

    /**
* public constructor
*
* @param SimpleXMLElement $xml Simplexml element to iterator over
* @param XMLSXReaderSharedStrings $shared_strings Xml element containing shared strings
* @param array $dimensions Array with info about dimensions of sheet
* @param array $options Array of library options
*
* @access public
* @return void
*/
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, array $dimensions, array $options)
    {
        $this->xml = $xml;
        $this->shared_strings = $shared_strings;
        $this->options = $options;
        $this->fillDimensions($dimensions);
        $this->position = 0;

        $this->checkValidPositions();

        if (!empty($options['indexing']) && in_array($options['indexing'], array('spreadsheet', 'spreadsheet-numerical'))) {
            $this->key_addition = 1;
        }
    }

    /**
* checks that dimensions are valid
*
* @param array $dimensions Array of dimensions f sheet
*
* @throws XLSXReaderException
* @access protected
* @return void
*/
    protected function fillDimensions(array $dimensions)
    {
        if (!in_array('row_start', array_keys($dimensions))) {
            throw new XLSXReaderException('Lacking row start from dimensions');
        }

        if (!in_array('cell_start', array_keys($dimensions))) {
            throw new XLSXReaderException('Lacking cell start from dimensions');
        }

        if (!isset($dimensions['row_count'])) {
            throw new XLSXReaderException('Lacking row count from dimensions');
        }

        if (!isset($dimensions['cell_count'])) {
            throw new XLSXReaderException('Lacking cell count from dimensions');
        }

        $this->cell_count = intval($dimensions['cell_count']);
        $this->cell_start = intval($dimensions['cell_start']);
        $this->row_count = intval($dimensions['row_count']);
        $this->row_start = intval($dimensions['row_start']);
    }

    /**
* checks which positions are valid in
* terms of rows
*
* @access protected
* @return void
*/
    protected function checkValidPositions()
    {
        $pos = $this->row_start - 1;
        $idx = $idx2 = 0;
        foreach ($this->xml->sheetData->row as $row) {
            while (($pos + 1) < (int) $row['r']) {
                $this->valid_positions[$idx++] = false;
                $pos++;
            }

            $this->valid_positions[$idx++] = $idx2++;
            $pos++;
        }
    }

    /**
* returns number of rows in sheet
*
* @access public
* @return int
*/
    public function count()
    {
        return $this->row_count;
    }

    /**
* returns the currently focused element of the rows
*
* @access public
* @return XLSXReaderRowIterator
*/
    public function current()
    {
        return $this;
    }

    /**
* returns a cell iterator for the row
*
* @access public
* @return XLSXReaderCellIterator
*/
    public function getCellIterator()
    {
        $class = 'XLSXReaderFakeCellIterator';
        if ($this->valid_positions[$this->position] !== false) {
            $class = 'XLSXReaderCellIterator';
        }

        return new $class(
            $this->xml->sheetData->row[$this->valid_positions[$this->position]],
            $this->shared_strings,
            $this->cell_start,
            $this->cell_count,
            $this->options
        );
    }

    /**
* checks if the current position is valid
*
* @access public
* @return bool
*/
    public function valid()
    {
        return $this->position > -1 && $this->position < $this->row_count;
    }

    /**
* increases the position by 1
*
* @access public
* @return void
*/
    public function next()
    {
        $this->position++;
    }

    /**
* rewinds the iterator position
*
* @access public
* @return void
*/
    public function rewind()
    {
        $this->position = 0;
    }

    /**
* returns the current position in the row-set
*
* @access public
* @return int
*/
    public function key()
    {
        if (!empty($this->options['indexing']) && $this->options['indexing'] === 'spreadsheet') {
            return (int) $this->xml->sheetData->row[$this->valid_positions[$this->position]]['r'];
        }
        return $this->position + $this->key_addition;
    }
}

