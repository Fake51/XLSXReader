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

/**
 * utility class for including other tests
 *
 * @category XLSXReader
 * @package  Tests
 * @author   Peter Lind <peter.e.lind@gmail.com>
 * @license  ./COPYRIGHT FreeBSD license
 * @link     http://plind.dk/xlsxreader
 */
class SuiteBase extends PHPUnit_Framework_TestSuite
{

    /**
     * includes all .php files in the $folder
     * directory of the child in the unit test
     *
     * @access public
     * @return SuiteBase
     */
    public static function init()
    {
        $c = get_called_class();
        $suite = new $c;
        foreach (new DirectoryIterator(static::$folder) as $file) {
            if ($file->isDot() || substr($file->getFilename(), -4) !== '.php') {
                continue;
            }

            $suite->addTestFile(static::$folder .DIRECTORY_SEPARATOR .  $file->getFilename());
        }

        return $suite;
    }

}
