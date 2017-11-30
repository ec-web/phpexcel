<?php
/**
 * XLS Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel\Reader;

use EC\PHPExcel\Parser\Excel5;
use EC\PHPExcel\Parser\OLERead;

class Xls extends BaseReader {
    /**
     * Xls parser
     *
     * @var Excel5
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
        $this->parser = new Excel5();
        $this->parser->loadOLE($file);

        $this->makeGenerator();

        return $this;
    }

    /**
     * Count elements of the selected sheet
     *
     * @return int
     */
    public function count() {
        $sheets = $this->sheets();

        return $sheets[$this->parser->getSheetIndex()]['totalRows'] ?? 0;
    }

    /**
     * Get the work sheets info
     *
     * @return array
     */
    public function sheets() {
        return $this->parser->parseWorksheetInfo();
    }

    /**
     * Make the generator
     */
    protected function makeGenerator() {
        $this->generator = call_user_func(function() {
            $line = 0;
            $count = $this->count();

            while (++$line <= $count) {
                if ($this->rowLimit > 0 && $line > $this->rowLimit) {
                    break;
                }

                $row = $this->parser->getCell($line - 1, $this->columnLimit);

                if (!$this->readEmptyCells && (empty($row) || trim(implode('', $row)) === '')) {
                    continue;
                }

                yield $row;
            }
        });
    }

    /**
     * Set sheet index
     *
     * @param int $index
     * @return $this
     */
    public function setSheetIndex($index = 0) {
        if ($index != $this->parser->getSheetIndex()) {
            $this->parser->setSheetIndex($index);

            $this->rewind();
        }

        return $this;
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
            // Use ParseXL for the hard work.
            $ole = new OLERead();

            // open file
            $ole->openFile($file);
        } catch (\Exception $e) {
            return false;
        }

        return true;
    }
}
