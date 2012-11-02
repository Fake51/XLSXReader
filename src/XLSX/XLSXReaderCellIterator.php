<?php

namespace XLSX;

use Iterator;
use SimpleXMLElement;
use Countable;

/**
* represents a row of cells in a sheet
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
*/
class XLSXReaderCellIterator implements Iterator, Countable
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
* alphabet range used for cell indexing
*
* @var array
*/
    protected static $alpha_range;

    /**
* flipped alphabet range used for cell indexing
*
* @var array
*/
    protected static $alpha_range_flipped;

    /**
* number of cells in this sheet
*
* @var int
*/
    protected $cell_count = 0;

    /**
* start of cells in this sheet
*
* @var int
*/
    protected $cell_start;

    /**
* shared strings xml document
*
* @var SimpleXMLElement
*/
    protected $shared_strings;

    /**
* data mapped from xml document
*
* @var array
*/
    protected $map_data;

    /**
* array of options for the library
*
* @var array
*/
    protected $options;

    /**
* integer added to key for
* array indexing
*
* @var int
*/
    protected $key_addition = 0;

    /**
* whether to use spreadsheet indexing
* for array keys
*
* @var bool
*/
    protected $use_spreadsheet_indexing = false;

    /**
* public constructor
*
* @param SimpleXMLElement $xml Xml element describing the cell
* @param XLSXReaderSharedStrings $shared_strings Xml element containing shared strings
* @param int $cell_start Start position of valid cells in sheet
* @param int $cell_count Number of cells in sheet
* @param array $options Array of library options
*
* @access public
* @return void
*/
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, $cell_start, $cell_count, array $options)
    {
        $this->xml = $xml;
        $this->shared_strings = $shared_strings;
        $this->cell_count = intval($cell_count);
        $this->cell_start = intval($cell_start);
        $this->options = $options;

        $this->position = 0;

        if (!empty($options['indexing'])) {
            if (in_array($options['indexing'], array('spreadsheet', 'spreadsheet-numerical'))) {
                $this->key_addition = 1;
            }

            if ($options['indexing'] === 'spreadsheet') {
                $this->use_spreadsheet_indexing = true;
            }
        }
    }

    /**
* returns number of cells in sheet
*
* @access public
* @return int
*/
    public function count()
    {
        return $this->cell_count;
    }

    /**
* maps out data for a row of cells
*
* @access protected
* @return void
*/
    protected function mapData()
    {
        foreach ($this->xml->c as $cell) {
            if (!isset($cell['r'])) {
                continue;
            }

            $xml_position = self::convertFromCellPosition((string) $cell['r']) - $this->cell_start;

            if (isset($cell['t']) && (string) $cell['t'] == 's') {
                $this->map_data[$xml_position] = array('value' => (string) $this->shared_strings->getItem((int) $cell->v), 'position' => (string) $cell['r']);
            } else {
                $this->map_data[$xml_position] = array('value' => (string) $cell->v, 'position' => (string) $cell['r']);
            }
        }

        for ($i = 0; $i < $this->cell_count; $i++) {
            if (!isset($this->map_data[$i])) {
                $this->map_data[$i] = null;
            }
        }

        ksort($this->map_data);
    }

    /**
* returns the cell contents at the current index
*
* @access public
* @return mixed
*/
    public function current()
    {
        if (!isset($this->map_data)) {
            $this->mapData();
        }

        if (isset($this->map_data[$this->position])) {
            return $this->map_data[$this->position]['value'];
        }

        return null;
    }

    /**
* returns the current position in the row
*
* @access public
* @return int
*/
    public function key()
    {
        if ($this->use_spreadsheet_indexing) {
            return $this->map_data[$this->position]['position'];
        }

        return $this->position + $this->key_addition;
    }

    /**
* checks if the current cell position is valid
*
* @access public
* @return bool
*/
    public function valid()
    {
        return $this->position > -1 && $this->position < $this->cell_count;
    }

    /**
* advances the cell pointer
*
* @access public
* @return void
*/
    public function next()
    {
        $this->position++;
    }

    /**
* resets the cell pointer
*
* @access public
* @return void
*/
    public function rewind()
    {
        $this->position = 0;
    }

    /**
* converts the cell position to a meaningful position
*
* @param string $index String index to convert to position
*
* @throws XLSXReaderException
* @access public
* @return int
*/
    public static function convertFromCellPosition($index)
    {
        if (!is_string($index)) {
            throw new XLSXReaderException('Param for convertCellIndex is not a string');
        }

        if (empty(self::$alpha_range)) {
            self::$alpha_range = array_flip(range('a', 'z'));
        }

        $alpha_part = strtolower(preg_replace('/^([^0-9]+).*$/i', '$1', $index));

        $return = 0;
        foreach (str_split(strtolower($alpha_part)) as $letter) {
            if (!isset(self::$alpha_range[$letter])) {
                throw new XLSXReaderException('Letter outside range: ' . $letter);
            }

            $return *= 26;
            $return += self::$alpha_range[$letter] + 1;
        }

        return $return;
    }

    /**
* converts row number and cell position
* to spreadsheet position, e.g. A1
*
* @param int $row_position Row position
* @param int $cell_position Cell position
*
* @throws XLSXReaderException
* @access public
* @return string
*/
    public static function convertToCellPosition($row_position, $cell_position)
    {
        throw new XLSXReaderException('Not implemented yet');

        if (!intval($row_position)) {
            throw new XLSXReaderException('Row position is not a number');
        }

        if (!intval($cell_position)) {
            throw new XLSXReaderException('Cell position is not a number');
        }

        if (empty(self::$alpha_range_flipped)) {
            self::$alpha_range_flipped = range('A', 'Z');
        }

        $cell_index = $cell_position - 1;
        $letter_end = $cell_index % 26;
        $cell_index -= $letter_end;
        $letters = array(self::$alpha_range_flipped[$letter_end]);
        while ($temp = floor($cell_index / 26)) {
            $letter = $temp % 26;
            $cell_index = $temp - $letter;
            array_unshift($letters, self::$alpha_range_flipped[$letter - 1]);
        }

        return implode('', $letters) . $row_position;
    }
}