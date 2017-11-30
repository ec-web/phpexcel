<?php
/**
 * XLSX Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel\Reader;

use EC\PHPExcel\Exception\ReaderException;
use EC\PHPExcel\Parser\Excel2007;

class Xlsx extends BaseReader {
    /**
     * Xls parser
     *
     * @var Excel2007
     */
    protected $parser;

    /**
     * Loads Excel from file
     *
     * @param string $file
     *
     * @return $this
     */
    public function load($file) {
        $this->parser = new Excel2007();
        $this->parser->loadZip($file);

        $this->makeGenerator();

        return $this;
    }

    /**
     * Count elements of an object
     *
     * @return int
     */
    public function count() {

    }

    /**
     * Get the work sheets info
     *
     * @return array
     */
    public function sheets() {

    }

    /**
     * Make the generator
     */
    protected function makeGenerator() {

    }

    /**
     * Set sheet index
     *
     * @param int $index
     * @return $this
     */
    public function setSheetIndex($index = 0) {

    }

    /**
     * Can the current Reader read the file?
     *
     * @param string $file
     *
     * @return bool
     */
    public function canRead($file) {
        try {
            $parser = new Excel2007();

            // open file
            $parser->openFile($file);
        } catch (\Exception $e) {
            return false;
        }

        return true;
    }
}
