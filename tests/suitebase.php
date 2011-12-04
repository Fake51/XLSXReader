<?php

class SuiteBase extends PHPUnit_Framework_TestSuite
{

    public static function init()
    {
        $c = get_called_class();
        $suite = new $c;
        foreach (new DirectoryIterator(static::$folder) as $file)
        {
            if ($file->isDot() || substr($file->getFilename(), -4) !== '.php') continue;
            $suite->addTestFile(static::$folder .DIRECTORY_SEPARATOR .  $file->getFilename());
        }
        return $suite;
    }

}
