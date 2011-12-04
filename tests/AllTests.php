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
require 'suitebase.php';

/**
 * runs all tests
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class AllTests extends SuiteBase
{
    protected static $folder = '.';

    /**
     * called by the phpunit framework
     *
     * @access public
     * @return PHPUnit_Framework_TestSuite
     */
    public static function suite()
    {
        return parent::init();
    }
}
