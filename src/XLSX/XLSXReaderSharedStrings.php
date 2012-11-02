<?php

namespace XLSX;
use SimpleXMLElement;

/**
* represents the sharedStrings file of the spreadsheet
*
* @category XLSXReader
* @package XLSXReader
* @author Peter Lind <peter.e.lind@gmail.com>
* @license ./COPYRIGHT FreeBSD license
* @link http://plind.dk/xlsxreader
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