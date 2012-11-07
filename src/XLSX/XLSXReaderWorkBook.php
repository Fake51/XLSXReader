<?php

namespace XLSX;

use ZipArchive;
use SimpleXMLElement;

/**
* represents the workbook xml file
* from the spreadsheet
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
*/
class XLSXReaderWorkBook
{
    const WORKBOOK_PATH = 'xl/workbook.xml';
    const SHARED_STRINGS_PATH = 'xl/sharedStrings.xml';
    const WORKSHEETS_PATH = 'xl/worksheets/';

    /**
* xml of the workbook file
*
* @var SimpleXMLElement
*/
    protected $xml;

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
* public constructor
*
* @param ZipArchive $zip ZipArchive instance containing spreadsheet
* @param array $options Options for the library
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

        if (!($this->xml = simplexml_load_string($this->zip->getFromName(self::WORKBOOK_PATH)))) {
            throw new XLSXReaderException('Could not load workboox xml file');
        }

        $this->options = $options;

        $this->extractSheetInformation();
        $this->loadSharedStrings();
        $this->loadSheets();
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

        foreach ($this->xml->sheets->sheet as $sheet) {
            if (!isset($sheet['sheetId'], $sheet['name'])) {
                throw new XLSXReaderException('Workbook is malformed, cannot read information');
            }

            $this->sheet_info[(int) $sheet['sheetId']] = (string) $sheet['name'];
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
        if ($xml_string = $this->zip->getFromName(self::SHARED_STRINGS_PATH)) {
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
        foreach ($this->sheet_info as $id => $name) {
            if (!($xml_string = $this->zip->getFromName(self::WORKSHEETS_PATH . 'sheet' . $id . '.xml'))) {
                throw new XLSXReaderException('Spreadsheet is malformed - cannot load worksheets');
            }

            $this->sheets[$id] = new XLSXReaderSheet(simplexml_load_string($xml_string), $this->shared_strings, $name, $id, $this->options);
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