<?php
/**
 * Xlsx Reader
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel\Reader;

use EC\PHPExcel\Parser\Excel2007;

class Xlsx extends BaseReader {
    /**
     * Xls parser
     *
     * @var Excel2007
     */
    protected $parser;

    /**
     * File row、column count
     *
     * @var array|int
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
        $this->parser = new Excel2007();
        $this->parser->loadZip($file);

        $this->generator = $this->makeGenerator();

        return $this;
    }

    /**
     * Count elements of an object
     *
     * @param bool $all
     * @return int|array
     */
    public function count($all = false) {
        if ($this->count === null) {
            $row = $column = 0;
            if ($sheet = $this->sheets($this->parser->getSheetIndex())) {
                $row = $sheet['totalRows'];
                $column = $sheet['totalColumns'];
            }

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
     * @param int $index
     * @return array
     */
    public function sheets($index = null) {
        $sheets = $this->parser->parseWorksheetInfo();

        if ($index !== null) {
            return $sheets[$index] ?? [];
        }

        return $sheets;
    }

    /**
     * Make the generator
     *
     * @return \Generator
     */
    protected function makeGenerator() {
        $line = 0;
        list($rowLimit, $columnLimit) = $this->count(true);

        while ($line < $rowLimit) {
            $row = $this->parser->getRow($line++, $columnLimit);

            if ($this->ignoreEmpty && (empty($row) || trim(implode('', $row)) === '')) {
                continue;
            }

            yield $row;
        }
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
            $parser = new Excel2007();

            // open file
            $parser->openFile($file);
        } catch (\Exception $e) {
            return false;
        }

        return true;
    }

    /**
     * Release parser and generator
     */
    public function __destruct() {
        $this->parser = null;
        $this->generator = null;
    }
}
