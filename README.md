# phpexcel
A lightweight PHP library for reading spreadsheet files
  - Based on Generator、SeekableIterator and Countable
  - Support for reading by line

### Requirements

  - PHP 7.0 or higher

### Installation

    composer require ecweb/phpexcel

## Usage

### csv

```
// Simple setting 
$reader = EC\PHPExcel\Excel::load('files/02.csv', 'GBK');

// Flexible setting
$reader = EC\PHPExcel\Excel::load('files/01.csv', function(EC\PHPExcel\Reader\Csv $reader) {
    // Set row limit
    $reader->setRowLimit(10);

    // Set column limit
    $reader->setColumnLimit(10);

    // Ignore emoty row
    $reader->ignoreEmptyRow(true);

    // Set encoding
    //$reader->setInputEncoding('GBK');

    // Set delimiter
    $reader->setDelimiter("\t");
}, 'GBK');

// skip to row 50 
$reader->seek(50);

// Get the current row data
$current = $reader->current();

// Get row count
$count = $reader->count();
```

### xls

```
$reader = EC\PHPExcel\Excel::load('files/01.xls', function(EC\PHPExcel\Reader\Xls $reader) {
    // Set row limit
    $reader->setRowLimit(10);

    // Set column limit
    $reader->setColumnLimit(10);

    // Ignore emoty row
    $reader->ignoreEmptyRow(true);

    // Select sheet index
    $reader->setSheetIndex(1);
});

// skip to row 50 
$reader->seek(50);

// Get the current row data
$current = $reader->current();

// Get row count
$count = $reader->count();

// Get all sheets info
$sheets = $reader->sheets();
```

### xlsx
```
$reader = EC\PHPExcel\Excel::load('files/01.xlsx', function(EC\PHPExcel\Reader\Xlsx $reader) {
    // Set row limit
    $reader->setRowLimit(10);

    // Set column limit
    $reader->setColumnLimit(10);

    // Ignore emoty row
    $reader->ignoreEmptyRow(true);

    // Select sheet index
    $reader->setSheetIndex(0);
});

// skip to row 50 
$reader->seek(50);

// Get the current row data
$current = $reader->current();

// Get row count
$count = $reader->count();

// Get all sheets info
$sheets = $reader->sheets();
```
