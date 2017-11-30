<?php
/**
 * Reader Abstract
 *
 * @author Janson
 * @create 2017-11-23
 */
namespace EC\PHPExcel\Reader;

use EC\PHPExcel\Contract\ReaderInterface;
use EC\PHPExcel\Exception\ReaderException;

abstract class BaseReader implements ReaderInterface {
    /**
     * Generator
     *
     * @var \Generator
     */
    protected $generator;

    /**
     * File row count
     *
     * @var int
     */
    protected $count;

    /**
     * Max row number
     *
     * @var int
     */
    protected $rowLimit;

    /**
     * Max column number
     *
     * @var int
     */
    protected $columnLimit;

    /**
     * Read empty cells?
     * Identifies whether the Reader should read data values for cells all cells, or should ignore cells containing
     *         null value or empty string
     *
     * @var boolean
     */
    protected $readEmptyCells = true;

    /**
     * Return the current element
     *
     * @return array
     */
    public function current() {
        return $this->generator->current();
    }

    /**
     * Move forward to next element
     */
    public function next() {
        $this->generator->next();
    }

    /**
     * Return the key of the current element
     *
     * @return int
     */
    public function key() {
        return $this->generator->key();
    }

    /**
     * Checks if current position is valid
     *
     * @return bool
     */
    public function valid() {
        return $this->generator->valid();
    }

    /**
     * Rewind the Iterator to the first element
     */
    public function rewind() {
        $this->makeGenerator();
    }

    /**
     * Make the generator
     */
    protected function makeGenerator() {

    }

    /**
     * Set read empty cells
     *        Set to true (the default) to advise the Reader read data values for all cells, irrespective of value.
     *        Set to false to advise the Reader to ignore cells containing a null value or an empty string.
     *
     * @param bool $readEmpty
     *
     * @return $this
     */
    public function setReadEmptyCells($readEmpty = true) {
        $this->readEmptyCells = $readEmpty;

        return $this;
    }

    /**
     * Set row limit
     *
     * @param int $limit
     * @return $this
     */
    public function setRowLimit($limit = null) {
        $this->rowLimit = $limit;

        return $this;
    }

    /**
     * Get row limit
     *
     * @return int
     */
    public function getRowLimit() {
        return $this->rowLimit;
    }

    /**
     * Set column limit
     *
     * @param int $limit
     * @return $this
     */
    public function setColumnLimit($limit = null) {
        $this->columnLimit = $limit;

        return $this;
    }

    /**
     * Takes a row and traverses the file to that row
     *
     * @param int $row
     */
    public function seek($row) {
        if ($row <= 0) {
            throw new \InvalidArgumentException("Row $row is invalid");
        }

        $key = $this->key();

        if ($key !== --$row) {
            if ($row < $key || is_null($key) || $row == 0) {
                $this->rewind();
            }

            while ($this->valid() && $row > $this->key()) {
                $this->next();
            }
        }
    }

    /**
     * Get column limit
     *
     * @return int
     */
    public function getColumnLimit() {
        return $this->columnLimit;
    }
}
