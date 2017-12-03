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
            list($rowLimit, $columnLimit) = $this->count(true);
            $this->parser->worksheet = null;

            while ($rowLimit--) {
                $row = $this->parser->getRow($columnLimit);

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
}
