<?php

namespace XLSX;

use ZipArchive;
use SimpleXMLElement;
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
 * represents the workbook xml file
 * from the spreadsheet
 *
 * @category XLSXReader
 * @package  XLSXReader
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ../../COPYRIGHT FreeBSD license
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
            throw new XLSXReaderException('Could not load workboox xml file');
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
