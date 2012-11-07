<?php

namespace XLSX;
use SimpleXMLElement;


/**
* represent a single spreadsheet in the file
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
*/
class XLSXReaderSheet
{
    /**
* xml of the sheet file
*
* @var SimpleXMLElement
*/
    protected $xml;

    /**
* name of the sheet
*
* @var string
*/
    protected $name;

    /**
* index of the sheet
*
* @var int
*/
    protected $index;

    /**
* tracks the number of rows in the sheet
*
* @var int
*/
    protected $row_count;

    /**
* where the rows start in the sheet
*
* @var int
*/
    protected $row_start = null;

    /**
* where the cells start in the sheet
*
* @var int
*/
    protected $cell_start = null;

    /**
* tracks the number of cells in the sheet
*
* @var int
*/
    protected $cell_count;

    /**
* XLSXReaderSharedStrings instance
*
* @var XLSXReaderSharedStrings
*/
    protected $shared_strings;

    /**
* options for the library
*
* @var array
*/
    protected $options;

    /**
* public constructor
*
* @param SimpleXMLElement $xml Simplexml element of the sheet
* @param XLSXReaderSharedStrings $shared_strings Shared strings object
* @param string $name Name of the spreadsheet
* @param int $index Index position of sheet
* @param array $options Options for the library
*
* @access public
* @return void
*/
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, $name, $index, array $options)
    {
        $this->xml = $xml;
        $this->name = $name;
        $this->index = $index;
        $this->shared_strings = $shared_strings;
        $this->options = $options;
    }

    /**
* returns the name of the sheet
*
* @access public
* @return string
*/
    public function getName()
    {
        return $this->name;
    }

    /**
* returns the index of the sheet
*
* @access public
* @return int
*/
    public function getPosition()
    {
        return $this->index;
    }

    /**
* returns an iterable and array-accesible
* object for iterating over the spreadsheet
* rows
*
* @access public
* @return XLSXReaderRowIterator
*/
    public function getRowIterator()
    {
        return new XLSXReaderRowIterator(
            $this->xml,
            $this->shared_strings,
            array(
                'row_start' => $this->getRowStart(),
                'cell_start' => $this->getCellStart(),
                'row_count' => $this->getRowCount(),
                'cell_count' => $this->getCellCount(),
            ),
            $this->options
        );
    }

    /**
* returns the row count of the sheet
*
* @access public
* @return int
*/
    public function getRowCount()
    {
        if (!isset($this->row_count)) {
            if (isset($this->xml->dimension) && strpos((string) $this->xml->dimension['ref'], ':')) {
                preg_match('/^[A-Z]+(\d+):[A-Z]+(\d+)$/', (string) $this->xml->dimension['ref'], $matches);
                $start = empty($this->options['read_minified']) ? 0 : intval($matches[1]) - 1;
                $this->row_count = intval($matches[2]) - $start;
                $this->row_start = intval($start + 1);
            } else {
                $this->row_count = 0;
            }
        }

        return $this->row_count;
    }

    /**
* returns number of cells in sheet
*
* @access public
* @return int
*/
    public function getCellCount()
    {
        if (!isset($this->cell_count)) {
            if (isset($this->xml->dimension)) {
                $parts = explode(':', (string) $this->xml->dimension['ref']);

                if (count($parts) === 1) {
                    $this->cell_count = 0;
                } else {
                    $start = empty($this->options['read_minified']) ? 0 : XLSXReaderCellIterator::convertFromCellPosition($parts[0]) - 1;
                    $this->cell_count = XLSXReaderCellIterator::convertFromCellPosition($parts[1]) - $start;
                    $this->cell_start = intval($start + 1);
                }
            } else {
                $this->cell_count = 0;
            }
        }

        return $this->cell_count;
    }

    /**
* returns the vertical position of the first
* row in the sheet
*
* @access public
* @return int
*/
    public function getRowStart()
    {
        $this->getRowCount();
        return $this->row_start;
    }

    /**
* returns the horizontal position of the first
* cell in the sheet
*
* @access public
* @return int
*/
    public function getCellStart()
    {
        $this->getCellCount();
        return $this->cell_start;
    }
}