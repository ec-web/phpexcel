<?php
/**
 * Excel2007 Parser
 *
 * @author Janson
 * @create 2017-11-29
 */
namespace EC\PHPExcel\Parser;

use EC\PHPExcel\Exception\ParserException;
use EC\PHPExcel\Exception\ReaderException;

class Excel2007 {
    /**
     * Number of shared strings that can be reasonably cached, i.e., that aren't read from file but stored in memory.
     * If the total number of shared strings is higher than this, caching is not used. If this value is null, shared
     * strings are cached regardless of amount. With large shared string caches there are huge performance gains,
     * however a lot of memory could be used which can be a problem, especially on shared hosting.
     */
    const SHARED_STRING_CACHE_LIMIT = 50000;

    /**
     * SimpleXMLElement XML object for the workbook XML file
     *
     * @var \SimpleXMLElement
     */
    protected $workbookXML;

    /**
     * SimpleXMLElement XML object for the style XML file
     *
     * @var \SimpleXMLElement
     */
    protected $stylesXML;

    /**
     * XMLReader for the shared strings XML file
     *
     * @var \XMLReader
     */
    protected $sharedStringsReader;

    /**
     * Shared strings XML file
     *
     * @var string
     */
    protected $sharedStringsFile;

    /**
     * Shared strings XML file count
     *
     * @var int
     */
    protected $sharedStringsCount;

    /**
     * Shared strings XML file cache
     *
     * @var array
     */
    protected $sharedStringCache = [];

    /**
     * Style XML file styles
     *
     * @var array
     */
    protected $styles = [];

    /**
     * Style XML file formats
     *
     * @var array
     */
    protected $formats = [];

    /**
     * Temp dir
     *
     * @var string
     */
    protected $tempDir;

    /**
     * Temp files
     *
     * @var array
     */
    protected $tempFiles = [];

    /**
     * Worksheets
     *
     * @var array
     */
    protected $sheets;

    /**
     * @var \DateTime
     */
    private static $baseDate;
    private static $decimalSeparator = '.';
    private static $thousandSeparator = ',';
    private static $currencyCode = '';
    private static $runtimeInfo = ['GMPSupported' => false];

    /**
     * Default options for libxml loader
     *
     * @var int
     */
    private static $libXmlLoaderOptions;

    /**
     * Use ZipArchive reader to extract the relevant data streams from the ZipArchive file
     *
     * @param string $file
     */
    public function loadZip($file) {
        $zip = $this->openFile($file);

        // Getting the general workbook information
        if ($zip->locateName('xl/workbook.xml') !== false) {
            $this->workbookXML = new \SimpleXMLElement($zip->getFromName('xl/workbook.xml'));
        }

        // Extracting the XMLs from the XLSX zip file
        if ($zip->locateName('xl/sharedStrings.xml') !== false) {
            $this->sharedStringsFile = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';

            $zip->extractTo($this->tempDir, 'xl/sharedStrings.xml');
            $this->tempFiles[] = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';

            if (is_readable($this->sharedStringsFile)) {
                $this->sharedStringsReader = new \XMLReader();
                $this->sharedStringsReader->open($this->sharedStringsFile);
                $this->prepareSharedStringCache();
            }
        }

        $this->parseWorksheetInfo();

        foreach ($this->sheets as $index => $name) {
            if ($zip->locateName('xl/worksheets/sheet' . $index . '.xml') !== false) {
                $zip->extractTo($this->tempDir, 'xl/worksheets/sheet' . $index . '.xml');

                $this->tempFiles[] = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets'
                    . DIRECTORY_SEPARATOR . 'sheet' . $index . '.xml';
            }
        }

        $this->ChangeSheet(0);

        // If worksheet is present and is OK, parse the styles already
        if ($zip->locateName('xl/styles.xml') !== false) {
            $this->stylesXML = new \SimpleXMLElement($zip->getFromName('xl/styles.xml'));
            if ($this->stylesXML && $this->stylesXML->cellXfs && $this->stylesXML->cellXfs->xf) {
                foreach ($this->stylesXML->cellXfs->xf as $xf) {
                    // Format #0 is a special case
                    // it is the "General" format that is applied regardless of applyNumberFormat
                    if ($xf->attributes()->applyNumberFormat || (0 == (int)$xf->attributes()->numFmtId)) {
                        // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                        $this->styles[] = (int)$xf->attributes()->numFmtId;
                    } else {
                        // 0 for "General" format
                        $this->styles[] = 0;
                    }
                }
            }

            if ($this->stylesXML->numFmts && $this->stylesXML->numFmts->numFmt) {
                foreach ($this->stylesXML->numFmts->numFmt as $numFmt) {
                    $this->formats[(int)$numFmt->attributes()->numFmtId] = (string)$numFmt->attributes()->formatcode;
                }
            }

            unset($this->stylesXML);
        }

        $zip->close();

        // Setting base date
        if (!self::$baseDate) {
            self::$baseDate = new \DateTime;
            self::$baseDate->setTimezone(new \DateTimeZone('UTC'));
            self::$baseDate->setDate(1900, 1, 0);
            self::$baseDate->setTime(0, 0, 0);
        }

        // Decimal and thousand separators
        if (!self::$decimalSeparator && !self::$thousandSeparator && !self::$currencyCode) {
            $locale = localeconv();

            self::$decimalSeparator = $locale['decimal_point'];
            self::$thousandSeparator = $locale['thousands_sep'];
            self::$currencyCode = $locale['int_curr_symbol'];
        }

        if (function_exists('gmp_gcd')) {
            self::$runtimeInfo['GMPSupported'] = true;
        }
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return array
     */
    public function parseWorksheetInfo() {
        if ($this->sheets === null) {
            $this->sheets = [];

            foreach ($this->workbookXML->sheets->sheet as $sheet) {
                $info = array(
                    'worksheetName' => (string)$sheet["name"],
                    'lastColumnLetter' => 'A',
                    'lastColumnIndex' => 0,
                    'totalRows' => 0,
                    'totalColumns' => 0,
                );

                $fileWorksheet = $worksheets[(string) self::array_item($eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];

                $xml = new XMLReader();
                $res = $xml->xml($this->securityScanFile('zip://'.PHPExcel_Shared_File::realpath($pFilename).'#'."$dir/$fileWorksheet"), null, PHPExcel_Settings::getLibXmlLoaderOptions());
                $xml->setParserProperty(2,true);

                $currCells = 0;
                while ($xml->read()) {
                    if ($xml->name == 'row' && $xml->nodeType == XMLReader::ELEMENT) {
                        $row = $xml->getAttribute('r');
                        $tmpInfo['totalRows'] = $row;
                        $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'],$currCells);
                        $currCells = 0;
                    } elseif ($xml->name == 'c' && $xml->nodeType == XMLReader::ELEMENT) {
                        $currCells++;
                    }
                }
                $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'],$currCells);
                $xml->close();

                $tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
                $tmpInfo['lastColumnLetter'] = Format::stringFromColumnIndex($tmpInfo['lastColumnIndex']);


                $attributes = $sheet->attributes('r', true);

                $sheetID = 0;
                foreach ($attributes as $name => $value) {
                    if ($name == 'id') {
                        $sheetID = (int)str_replace('rId', '', (string)$value);
                        break;
                    }
                }

                if ($sheetID) {
                    $this->sheets[$sheetID] = (string)$sheet['name'];
                }
            }

            ksort($this->sheets);
        }

        return $this->sheets ? array_values($this->sheets) : [];
    }

    /**
     * Open file for reading
     *
     * @param string $file
     *
     * @throws ParserException|ReaderException
     * @return \ZipArchive
     */
    public function openFile($file) {
        // Check if file exists
        if (!file_exists($file) || !is_readable($file)) {
            throw new ReaderException("Could not open file [$file] for reading! File does not exist.");
        }

        $zip = new \ZipArchive();
        $status = $zip->open($file);

        $xl = false;
        if ($status === true) {
            // check if it is an OOXML archive
            $rels = simplexml_load_string(
                $this->securityScan($zip->getFromName('rels/.rels')), 'SimpleXMLElement', self::getLibXmlLoaderOptions()
            );

            if ($rels !== false) {
                foreach ($rels->Relationship as $rel) {
                    switch ($rel["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                            if (basename($rel["Target"]) == 'workbook.xml') {
                                $xl = true;
                            }

                            break;
                    }
                }
            }
        }

        if ($status !== true || $xl === false) {
            throw new ParserException(
                "The file [$file] is not recognised as a zip archive: " . $zip->getStatusString()
            );
        }

        return $zip;
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param  string $xml
     *
     * @throws ReaderException
     * @return string
     */
    protected function securityScan($xml) {
        $pattern = sprintf('/\\0?%s\\0?/', implode('\\0?', str_split('<!DOCTYPE')));

        if (preg_match($pattern, $xml)) {
            throw new ReaderException(
                'Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks'
            );
        }

        return $xml;
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param  string $filestream
     *
     * @return string
     */
    protected function securityScanFile($filestream) {
        return $this->securityScan(file_get_contents($filestream));
    }

    /**
     * Set default options for libxml loader
     *
     * @param int $options Default options for libxml loader
     */
    public static function setLibXmlLoaderOptions($options = null) {
        if (is_null($options) && defined(LIBXML_DTDLOAD)) {
            $options = LIBXML_DTDLOAD | LIBXML_DTDATTR;
        }

        if (version_compare(PHP_VERSION, '5.2.11') >= 0) {
            @libxml_disable_entity_loader($options == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        }

        self::$libXmlLoaderOptions = $options;
    }

    /**
     * Get default options for libxml loader.
     * Defaults to LIBXML_DTDLOAD | LIBXML_DTDATTR when not set explicitly.
     *
     * @return int Default options for libxml loader
     */
    public static function getLibXmlLoaderOptions() {
        if (is_null(self::$libXmlLoaderOptions) && defined(LIBXML_DTDLOAD)) {
            self::setLibXmlLoaderOptions(LIBXML_DTDLOAD | LIBXML_DTDATTR);
        }

        if (version_compare(PHP_VERSION, '5.2.11') >= 0) {
            @libxml_disable_entity_loader(self::$libXmlLoaderOptions == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        }

        return self::$libXmlLoaderOptions;
    }

    /**
     * Set temp dir
     *
     * @param string $dir
     * @return $this
     */
    public function setTempDir($dir) {
        $this->tempDir = $dir;

        return $this;
    }

    /**
     * Creating shared string cache if the number of shared strings is acceptably low
     * (or there is no limit on the amount)
     *
     * @return bool
     */
    private function PrepareSharedStringCache() {
        while ($this->sharedStringsReader->read()) {
            if ($this->sharedStringsReader->name == 'sst') {
                $this->sharedStringsCount = $this->sharedStringsReader->getAttribute('count');
                break;
            }
        }

        if (!$this->sharedStringsCount || (self::SHARED_STRING_CACHE_LIMIT < $this->sharedStringsCount
                && self::SHARED_STRING_CACHE_LIMIT !== null)) {

            return false;
        }

        $cacheIndex = 0;
        $cacheValue = '';
        while ($this->sharedStringsReader->read()) {
            switch ($this->sharedStringsReader->name) {
                case 'si':
                    if ($this->sharedStringsReader->nodeType == \XMLReader::END_ELEMENT) {
                        $this->sharedStringCache[$cacheIndex] = $cacheValue;
                        $cacheIndex++;
                        $cacheValue = '';
                    }
                    break;

                case 't':
                    if ($this->sharedStringsReader->nodeType == \XMLReader::END_ELEMENT) {
                        continue;
                    }

                    $cacheValue .= $this->sharedStringsReader->readString();
                    break;
            }
        }

        $this->sharedStringsReader->close();

        return true;
    }

    /**
     * Formats the value according to the index
     *
     * @param string $value
     * @param int $index Format index
     *
     * @return string Formatted cell value
     */
    private function formatValue($value, $index) {
        if (!is_numeric($value)) {
            return $value;
        }

        if (isset($this->styles[$index]) && $this->styles[$index] !== false) {
            $index = $this->styles[$index];
        } else {
            return $value;
        }

        // A special case for the "General" format
        if ($index == 0) {
            return $this->generalFormat($value);
        }

        $format = $this->parsedFormatCache[$index] ?? [];

        if (empty($format)) {
            $format = [
                'code' => false, 'type' => false, 'scale' => 1, 'thousands' => false, 'currency' => false
            ];

            if (isset(Format::$buildInFormats[$index])) {
                $format['code'] = Format::$buildInFormats[$index];
            } elseif (isset($this->formats[$index])) {
                $format['code'] = str_replace('"', '', $this->formats[$index]);
            }

            // Format code found, now parsing the format
            if ($format['code']) {
                $sections = explode(';', $format['code']);
                $format['code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['code'] = $sections[1];
                        }

                        $value = abs($value);
                        break;

                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['code'] = $sections[1];
                        } elseif ($value == 0) {
                            $format['code'] = $sections[2];
                        }

                        $value = abs($value);
                        break;
                }
            }

            // Stripping colors
            $format['code'] = trim(preg_replace('/^\\[[a-zA-Z]+\\]/', '', $format['code']));

            // Percentages
            if (substr($format['code'], -1) == '%') {
                $format['type'] = 'Percentage';
            } elseif (preg_match('/(\[\$[A-Z]*-[0-9A-F]*\])*[hmsdy]/i', $format['code'])) {
                $format['type'] = 'DateTime';
                $format['code'] = trim(preg_replace('/^(\[\$[A-Z]*-[0-9A-F]*\])/i', '', $format['code']));
                $format['code'] = strtolower($format['code']);
                $format['code'] = strtr($format['code'], Format::$dateFormatReplacements);

                if (strpos($format['code'], 'A') === false) {
                    $format['code'] = strtr($format['code'], Format::$dateFormatReplacements24);
                } else {
                    $format['code'] = strtr($format['code'], Format::$dateFormatReplacements12);
                }
            } elseif ($format['code'] == '[$EUR ]#,##0.00_-') {
                $format['type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['code'] = preg_replace('/_./', '', $format['code']);

                // Removing unnecessary escaping
                $format['code'] = preg_replace("/\\\\/", '', $format['code']);

                // Removing string quotes
                $format['code'] = str_replace(['"', '*'], '', $format['code']);

                // Removing thousands separator
                if (strpos($format['code'], '0,0') !== false || strpos($format['code'], '#,#') !== false) {
                    $format['thousands'] = true;
                }

                $format['code'] = str_replace(['0,0', '#,#'], ['00', '##'], $format['code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = [];

                if (preg_match('/(0|#)(,+)/', $format['code'], $matches)) {
                    $scale = pow(1000, strlen($matches[2]));

                    // Removing the commas
                    $format['code'] = preg_replace(['/0,+/', '/#,+/'], ['0', '#'], $format['code']);
                }

                $format['scale'] = $scale;
                if (preg_match('/#?.*\?\/\?/', $format['code'])) {
                    $format['type'] = 'Fraction';
                } else {
                    $format['code'] = str_replace('#', '', $format['code']);
                    $matches = [];

                    if (preg_match('/(0+)(\.?)(0*)/', preg_replace('/\[[^\]]+\]/', '', $format['code']), $matches)) {
                        list(, $integer, $decimalPoint, $decimal) = $matches;

                        $format['minWidth'] = strlen($integer) + strlen($decimalPoint) + strlen($decimal);
                        $format['decimals'] = $decimal;
                        $format['precision'] = strlen($format['decimals']);
                        $format['pattern'] = '%0' . $format['minWidth'] . '.' . $format['precision'] . 'f';
                    }
                }

                $matches = [];
                if (preg_match('/\[\$(.*)\]/u', $format['code'], $matches)) {
                    $currencyCode = explode('-', $matches[1]);
                    if ($currencyCode) {
                        $currencyCode = $currencyCode[0];
                    }
                    
                    if (!$currencyCode) {
                        $currencyCode = self::$currencyCode;
                    }
                    
                    $format['currency'] = $currencyCode;
                }

                $format['code'] = trim($format['code']);
            }

            $this->parsedFormatCache[$index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['code'] == '@') {
                return (string)$value;
            } elseif ($format['type'] == 'Percentage') { // Percentages
                if ($format['code'] === '0%') {
                    $value = round(100*$value, 0) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100*$value, 2));
                }
            } elseif ($format['type'] == 'DateTime') { // Dates and times
                $days = (int)$value;

                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int)$value);

                // Here time is converted to seconds
                // Some loss of precision will occur
                $seconds = $time ? (int)($time*86400) : 0;

                $value = clone self::$baseDate;
                $value->add(new \DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));

                $value = $value->format($format['code']);
            } elseif ($format['type'] == 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers
                if ($format['type'] == 'Fraction' && ($value != (int)$value)) {
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);

                    // Removing the integer part and decimal point
                    $decimal *= pow(10, strlen($decimal) - 2);
                    $decimalDivisor = pow(10, strlen($decimal));

                    if (self::$runtimeInfo['GMPSupported']) {
                        $GCD = gmp_strval(gmp_gcd($decimal, $decimalDivisor));
                    } else {
                        $GCD = self::GCD($decimal, $decimalDivisor);
                    }

                    $adjDecimal = $decimal/$GCD;
                    $adjDecimalDivisor = $decimalDivisor/$GCD;

                    if (strpos($format['code'], '0') !== false || strpos($format['code'], '#') !== false
                        || substr($format['code'], 0, 3) == '? ?') {

                        // The integer part is shown separately apart from the fraction
                        $value = ($value < 0 ? '-' : '') . $integer ? $integer . ' '
                            : '' . $adjDecimal . '/' . $adjDecimalDivisor;
                    } else {
                        // The fraction includes the integer part
                        $adjDecimal += $integer * $adjDecimalDivisor;
                        $value = ($value < 0 ? '-' : '') . $adjDecimal . '/' . $adjDecimalDivisor;
                    }
                } else {
                    // Scaling
                    $value = $value/$format['scale'];
                    if (!empty($format['minWidth']) && $format['decimals']) {
                        if ($format['thousands']) {
                            $value = number_format(
                                $value, $format['precision'], self::$decimalSeparator, self::$thousandSeparator
                            );

                            $value = preg_replace('/(0+)(\.?)(0*)/', $value, $format['code']);
                        } else {
                            if (preg_match('/[0#]E[+-]0/i', $format['code'])) {
                                // Scientific format
                                $value = sprintf('%5.2E', $value);
                            } else {
                                $value = sprintf($format['pattern'], $value);
                                $value = preg_replace('/(0+)(\.?)(0*)/', $value, $format['code']);
                            }
                        }
                    }
                }

                // currency/Accounting
                if ($format['currency']) {
                    $value = preg_replace('', $format['currency'], $value);
                }
            }
        }

        return $value;
    }
}
