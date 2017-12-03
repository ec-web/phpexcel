<?php
/**
 * Xls Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel\Reader;

use EC\PHPExcel\Parser\Excel5;
use EC\PHPExcel\Parser\Excel5\OLERead;

class Xls extends BaseReader {
    /**
     * Xls parser
     *
     * @var Excel5
     */
    protected $parser;

    /**
     * File row and column count
     *
     * @var array
     */
    protected $count;

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
     * @param bool $all
     * @return int|array
     */
    public function count($all = false) {
        if ($this->count === null) {
            $sheet = $this->sheets()[$this->parser->getSheetIndex()] ?? [];
            $row = $sheet['totalRows'] ?? 0;
            $column = $sheet['totalColumns'] ?? 0;

            $this->count = [
                $this->rowLimit > 0 && $row > $this->rowLimit ? $this->rowLimit : $row,
                $this->columnLimit > 0 && $column > $this->columnLimit ? $this->columnLimit : $column
            ];
        }

        return $all ? $this->count : $this->count[0];
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
            list($rowLimit, $columnLimit) = $this->count(true);

            while (++$line <= $rowLimit) {
                $row = $this->parser->getRow($line - 1, $columnLimit);

                if (!$this->readEmptyCells && (empty($row) || trim(implode('', $row)) === '')) {
                    continue;
                }

                // Fill the empty cell
                for ($i = 0; $i < $columnLimit; $i++) {
                    if (!isset($row[$i])) {
                        $row[$i] = '';
                    }
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
    public function setSheetIndex($index) {
        if ($index != $this->parser->getSheetIndex()) {
            $this->parser->setSheetIndex($index);

            $this->count = null;
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
