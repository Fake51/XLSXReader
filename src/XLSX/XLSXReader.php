<?php

namespace XLSX;

use ZipArchive;

/**
* base class, gives access to the .xlsx file
* contents in terms of the worksheets
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
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
        'indexing' => 'php-array',
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
* @param mixed $value Value to set for the option
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
* @return XLSXReaderSharedStrings
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