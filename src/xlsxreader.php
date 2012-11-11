<?php
/**
 * .xlsx file reader library
 *
 * PHP Version 5.3+
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @version  1.1
 * @link     http://plind.dk/xlsxreader
 */

/**
 * exception class for library
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class XLSXReaderException extends Exception
{
}

/**
 * base class, gives access to the .xlsx file
 * contents in terms of the worksheets
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class XLSXReader
{
    /**
     * name of .xlsx file
     *
     * @var string
     */
    protected $filename;

    /**
     * ZipArchive of spreadsheet
     *
     * @var ZipArchive
     */
    protected $zip;

    /**
     * XLSXReaderWorkBook instance
     *
     * @var XLSXReaderWorkBook
     */
    protected $workbook;

    /**
     * default options of the library
     *
     * @var array
     */
    protected $default_options = array(
        'read_minified' => false,
        'indexing'      => 'php-array',
    );

    /**
     * manually set options
     *
     * @var array
     */
    protected $options = array();

    /**
     * public constructor
     *
     * @param string $filename Name of .xlsx file to read
     *
     * @throws XLSXReaderException
     * @access public
     * @return void
     */
    public function __construct($filename)
    {
        if (!is_file($filename)) {
            throw new XLSXReaderException($filename . ' does not appear to be a valid file');
        }

        $this->filename = $filename;

        if (!extension_loaded('zip') || !class_exists('ZipArchive')) {
            throw new XLSXReaderException('Zip extension is not loaded');
        }

        if (!extension_loaded('SimpleXML')) {
            throw new XLSXReaderException('XLSXReader relies upon SimpleXML to work - it is not present');
        }

        $this->zip = new ZipArchive();
        $this->zip->open($filename);
    }

    /**
     * sets the value for an option
     *
     * @param string $option Name of option to set
     * @param mixed  $value  Value to set for the option
     *
     * @throws XLSXReaderException
     * @access public
     * @return $this
     */
    public function setOption($option, $value)
    {
        if (!in_array($option, $this->getAvailableOptions())) {
            throw new XLSXReaderException('No such option available');
        }

        $this->options[$option] = $value;
        return $this;
    }

    /**
     * returns the options set or defaults
     *
     * @access public
     * @return array
     */
    public function getOptions()
    {
        return array_merge($this->default_options, $this->options);
    }

    /**
     * returns array of recognized settings
     *
     * @access public
     * @return array
     */
    public function getAvailableOptions()
    {
        return array_keys($this->default_options);
    }

    /**
     * returns array of sheet objects
     * based on available sheets in spreadsheet
     * file
     *
     * @access public
     * @return array
     */
    public function getSheets()
    {
        return $this->getWorkBook()->getSheets();
    }

    /**
     * returns the shared strings instance
     *
     * @access public
     * @return XLSXReaderSharedString
     */
    public function getSharedStrings()
    {
        return $this->getWorkBook()->getSharedStrings();
    }

    /**
     * returns the workbook class
     * for the spreadsheet
     *
     * @access public
     * @return XLSXReaderWorkBook
     */
    public function getWorkBook()
    {
        if (empty($this->workbook)) {
            $this->workbook = new XLSXReaderWorkBook($this->zip, $this->getOptions());
        }

        return $this->workbook;
    }
}

/**
 * represents the workbook xml file
 * from the spreadsheet
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class XLSXReaderWorkBook
{
    const BASE_PREFIX         = 'xl/';
    const WORKBOOK_PATH       = 'workbook.xml';
    const WORKBOOK_RELS_PATH  = '_rels/workbook.xml.rels';
    const SHARED_STRINGS_PATH = 'sharedStrings.xml';

    /**
     * xml of the workbook file
     *
     * @var SimpleXMLElement
     */
    protected $xml;

    /**
     * xml of the workbook relations file
     *
     * @var SimpleXMLElement
     */
    protected $xml_rels;

    /**
     * array of information on sheets in spreadsheet
     *
     * @var array
     */
    protected $sheet_info;

    /**
     * array of spreadsheet sheet objects
     *
     * @var array
     */
    protected $sheets;

    /**
     * shared strings object
     *
     * @var XLSXReaderSharedStrings
     */
    protected $shared_strings;

    /**
     * ziparchive of the spreadsheet
     *
     * @var ZipArchive
     */
    protected $zip;

    /**
     * options for the library
     *
     * @var array
     */
    protected $options;

    /**
     * contains the relationships of the workbook
     *
     * @var array
     */
    protected $relationships = array();

    /**
     * public constructor
     *
     * @param ZipArchive $zip     ZipArchive instance containing spreadsheet
     * @param array      $options Options for the library
     *
     * @access public
     * @return void
     */
    public function __construct(ZipArchive $zip, array $options)
    {
        if (!$zip->numFiles) {
            throw new XLSXReaderException('Zip object does not represent actual zip file');
        }

        $this->zip = $zip;

        if (!($this->xml = simplexml_load_string($this->zip->getFromName(self::BASE_PREFIX . self::WORKBOOK_PATH)))) {
            throw new XLSXReaderException('Could not load workbook xml file');
        }

        if (!($this->xml_rels = simplexml_load_string($this->zip->getFromName(self::BASE_PREFIX . self::WORKBOOK_RELS_PATH)))) {
            throw new XLSXReaderException('Could not load workbook relations xml file');
        }

        $this->options = $options;

        $this->extractRelationInformation();
        $this->extractSheetInformation();
        $this->loadSharedStrings();
        $this->loadSheets();
    }

    /**
     * runs through the relations xml to extract information
     * about workbook relations - including worksheet names
     *
     * @access protected
     * @return void
     */
    protected function extractRelationInformation()
    {
        if (!$this->xml_rels->Relationship) {
            throw new XLSXReaderException('Could not load workbook relations from xml file');
        }

        foreach ($this->xml_rels->Relationship as $relationship) {
            $id                       = (string) $relationship['Id'];
            $this->relationships[$id] = array(
                'id'     => $id,
                'type'   => (string) $relationship['Type'],
                'target' => (string) $relationship['Target'],
            );
        }
    }

    /**
     * grabs the sheet information from
     * the xml file
     *
     * @throws XLSXReaderException
     * @access protected
     * @return void
     */
    protected function extractSheetInformation()
    {
        $this->sheets = array();

        if (!isset($this->xml->sheets, $this->xml->sheets->sheet)) {
            throw new XLSXReaderException('Workbook is malformed, cannot read information');
        }

        $namespaces = $this->xml->getNameSpaces(true);
        if (!isset($namespaces['r'])) {
            throw new XLSXReaderException('Workbook lacks relationship namespace, cannot load worksheets');
        }

        foreach ($this->xml->sheets->sheet as $sheet) {
            $attributes = $sheet->attributes($namespaces['r']);
            if (!$attributes->id) {
                throw new XLSXReaderException('Workbook sheet info malformed, no relational ID');
            }

            $sheet_id = (string) $attributes->id;
            if (!isset($this->relationships[$sheet_id])) {
                throw new XLSXReaderException('Workbook sheet info invalid, refers to non-existing relation');
            }

            if (!isset($sheet['name'])) {
                throw new XLSXReaderException('Workbook is malformed, cannot read information');
            }

            $this->sheet_info[$sheet_id] = (string) $sheet['name'];
        }
    }

    /**
     * finds the sharedStrings xml file
     * and loads it as simplexml
     *
     * @throws XLSXReaderException
     * @access protected
     * @return void
     */
    protected function loadSharedStrings()
    {
        $xml = null;
        if ($xml_string = $this->zip->getFromName(self::BASE_PREFIX . self::SHARED_STRINGS_PATH)) {
            $xml = simplexml_load_string($xml_string);
        }

        $this->shared_strings = new XLSXReaderSharedStrings($xml);
    }

    /**
     * load sheets as xml strings and turns them
     * into simplexml objects
     *
     * @throws XLSXReaderException
     * @access protected
     * @return void
     */
    protected function loadSheets()
    {
        $index = 1;
        foreach ($this->sheet_info as $id => $name) {
            if (!isset($this->relationships[$id])) {
                throw new XLSXReaderException('Spreadsheet is malformed - refers to non-existing relationship');
            }

            if (!($xml_string = $this->zip->getFromName(self::BASE_PREFIX . $this->relationships[$id]['target']))) {
                throw new XLSXReaderException('Spreadsheet is malformed - cannot load worksheets');
            }

            $this->sheets[$index] = new XLSXReaderSheet(simplexml_load_string($xml_string), $this->shared_strings, $name, $index, $this->options);
            $index++;
        }
    }

    /**
     * returns array with id and names of sheets
     * in the spreadsheet
     *
     * @access public
     * @return array
     */
    public function getSheetNames()
    {
        return $this->sheet_info;
    }

    /**
     * returns array of sheets
     *
     * @access public
     * @return array
     */
    public function getSheets()
    {
        return array_values($this->sheets);
    }

    /**
     * returns the shared strings object
     *
     * @access public
     * @return XLSXReaderSharedStrings
     */
    public function getSharedStrings()
    {
        return $this->shared_strings;
    }
}

/**
 * represent a single spreadsheet in the file
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
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
     * @param SimpleXMLElement        $xml            Simplexml element of the sheet
     * @param XLSXReaderSharedStrings $shared_strings Shared strings object
     * @param string                  $name           Name of the spreadsheet
     * @param int                     $index          Index position of sheet
     * @param array                   $options        Options for the library
     *
     * @access public
     * @return void
     */
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, $name, $index, array $options)
    {
        $this->xml             = $xml;
        $this->name            = $name;
        $this->index           = $index;
        $this->shared_strings  = $shared_strings;
        $this->options         = $options;
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
                'row_start'  => $this->getRowStart(),
                'cell_start' => $this->getCellStart(),
                'row_count'  => $this->getRowCount(),
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
                $start           = empty($this->options['read_minified']) ? 0 : intval($matches[1]) - 1;
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
                    $start            = empty($this->options['read_minified']) ? 0 : XLSXReaderCellIterator::convertFromCellPosition($parts[0]) - 1;
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

/**
 * represents the sharedStrings file of the spreadsheet
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class XLSXReaderSharedStrings
{
    /**
     * xml of the sharedStrings file
     *
     * @var SimpleXMLElement
     */
    protected $xml;

    /**
     * public constructor
     *
     * @param SimpleXMLElement $xml Simplexml element of the sheet
     *
     * @access public
     * @return void
     */
    public function __construct(SimpleXMLElement $xml = null)
    {
        $this->xml = $xml;
    }

    /**
     * returns the item from the given position
     *
     * @param int $index Index to fetch from
     *
     * @throws XLSXReaderException
     * @access public
     * @return string
     */
    public function getItem($index)
    {
        if ($this->xml && isset($this->xml->si[$index])) {
            return (string) $this->xml->si[$index]->t;
        }

        throw new XLSXReaderException('No shared string at index: ' . $index);
    }
}

/**
 * row iterator - allows for iterating over rows
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
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
     * @param SimpleXMLElement         $xml            Simplexml element to iterator over
     * @param XMLSXReaderSharedStrings $shared_strings Xml element containing shared strings
     * @param array                    $dimensions     Array with info about dimensions of sheet
     * @param array                    $options        Array of library options
     *
     * @access public
     * @return void
     */
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, array $dimensions, array $options)
    {
        $this->xml            = $xml;
        $this->shared_strings = $shared_strings;
        $this->options        = $options;
        $this->fillDimensions($dimensions);
        $this->position       = 0;

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
        $this->row_count  = intval($dimensions['row_count']);
        $this->row_start  = intval($dimensions['row_start']);
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

/**
 * represents a row of cells in a sheet
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
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
     * @param SimpleXMLElement        $xml            Xml element describing the cell
     * @param XLSXReaderSharedStrings $shared_strings Xml element containing shared strings
     * @param int                     $cell_start     Start position of valid cells in sheet
     * @param int                     $cell_count     Number of cells in sheet
     * @param array                   $options        Array of library options
     *
     * @access public
     * @return void
     */
    public function __construct(SimpleXMLElement $xml, XLSXReaderSharedStrings $shared_strings, $cell_start, $cell_count, array $options)
    {
        $this->xml            = $xml;
        $this->shared_strings = $shared_strings;
        $this->cell_count     = intval($cell_count);
        $this->cell_start     = intval($cell_start);
        $this->options        = $options;

        $this->position       = 0;

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
     * @param int $row_position  Row position
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

        $cell_index  = $cell_position - 1;
        $letter_end  = $cell_index % 26;
        $cell_index -= $letter_end;
        $letters     = array(self::$alpha_range_flipped[$letter_end]);
        while ($temp = floor($cell_index / 26)) {
            $letter     = $temp % 26;
            $cell_index = $temp - $letter;
            array_unshift($letters, self::$alpha_range_flipped[$letter - 1]);
        }

        return implode('', $letters) . $row_position;
    }
}

/**
 * an empty row of cells
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../COPYRIGHT FreeBSD license
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
     * @param null                    $xml            Xml element describing the cell
     * @param XLSXReaderSharedStrings $shared_strings Xml element containing shared strings
     * @param int                     $cell_start     Start position of valid cells in sheet
     * @param int                     $cell_count     Number of cells in sheet
     * @param array                   $options        Array of library options
     *
     * @access public
     * @return void
     */
    public function __construct($xml, $shared_strings, $cell_start, $cell_count, $options)
    {
        $this->cell_count     = intval($cell_count);
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
