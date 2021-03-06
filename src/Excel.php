<?php
/**
 * PHP Excel
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel;

use EC\PHPExcel\Exception\ReaderException;

class Excel {
    /**
     * Load a file
     *
     * @param string $file
     * @param callback|null $callback
     * @param string|null $encoding
     * @param string $ext
     *
     * @throws ReaderException
     * @return \EC\PHPExcel\Reader\BaseReader
     */
    public static function load($file, $callback = null, $encoding = null, $ext = '') {
        set_error_handler(function($errno, $errstr, $errfile, $errline) {
            // no-op
        }, E_ALL ^ E_ERROR);

        $ext = $ext ?: strtolower(pathinfo($file, PATHINFO_EXTENSION));

        $format = self::getFormatByExtension($ext);

        if (empty($format)) {
            throw new ReaderException("Could not identify file format for file [$file] with extension [$ext]");
        }

        $class = __NAMESPACE__ . '\\Reader\\' . $format;
        $reader = new $class;

        if ($callback) {
            if ($callback instanceof \Closure) {
                // Do the callback
                call_user_func($callback, $reader);
            } elseif (is_string($callback)) {
                // Set the encoding
                $encoding = $callback;
            }
        }

        if ($encoding && method_exists($reader, 'setInputEncoding')) {
            $reader->setInputEncoding($encoding);
        }

        return $reader->load($file);
    }

    /**
     * Identify file format
     *
     * @param string $ext
     * @return string
     */
    protected static function getFormatByExtension($ext) {
        $formart = '';

        switch ($ext) {
            /*
            |--------------------------------------------------------------------------
            | Excel 2007
            |--------------------------------------------------------------------------
            */
            case 'xlsx':
            case 'xlsm':
            case 'xltx':
            case 'xltm':
                $formart = 'Xlsx';
                break;

            /*
            |--------------------------------------------------------------------------
            | Excel5
            |--------------------------------------------------------------------------
            */
            case 'xls':
            case 'xlt':
                $formart = 'Xls';
                break;

            /*
            |--------------------------------------------------------------------------
            | CSV
            |--------------------------------------------------------------------------
            */
            case 'csv':
            case 'txt':
                $formart = 'Csv';
                break;
        }

        return $formart;
    }
}
