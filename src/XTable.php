<?php
namespace Noctis;

use \PHPExcel;
use \PHPExcel_Style_Border;
use \PHPExcel_Style_Fill;
use \PHPExcel_Writer_Excel2007;
use \PHPExcel_Writer_Excel5;


/**
 * Wrapper for the PHPExcel library, used to generate Excel files (BIFF/Open XML formats)
 *
 * For now only the most basic options are available
 *
 * @author Łukasz Czejgis
 */
class XTable
{
    private $sheet;     // Sheet
    private $excel;     // Worksheet
    private $row;       // Row number (1..)
    private $coll;      // Column number (yes, number) (0..)
    private $maxColl;
    private $debug;     // Display additional information?
    private $rangeStart;
    private $rangeEnd;
    private $defaultFontSize;
    private $sheetNo;   // numer of the current sheet in the worksheet
    private $globalCellOptions;
    private $currentRowOptions;

    /**
     * Available parameters (all are optional):
     *   > start_row - row number from which we start to fill cells (default value: 1)
     *   > start_coll - column number from which we start to fill cells (default value: 0; 0 = A, 1 = B, 2 = C, etc.)
     *   > creator - worksheet author
     *   > modified_by - name of person who last modified the worksheet
     *   > title - worksheet title
     *   > subject - worksheet subject
     *   > description - worksheet description (default value: date and time when the document was created)
     *   > keywords - worksheet keywords
     *   > category - worksheet category
     *   > sheet_title - sheet name (sheet, not worksheet!)
     *
     * @param array $params Parameters
     */
    public function __construct($params = array())
    {
        $this->row  = ( isset($params['start_row'])  ? $params['start_row']  : 1 );
        $this->coll = ( isset($params['start_coll']) ? $params['start_coll'] : 0 );
        $this->maxColl = $this->coll;
        $this->debug = false;               // By default we want to be quiet be (otherwise the generated file can not be sent)
        $this->excel = new PHPExcel();
        $this->globalCellOptions = [];
        $this->currentRowOptions = [];

        if ( isset($params['creator']) ) {
            $this->excel->getProperties()->setCreator(trim($params['creator']));
        }

        if ( isset($params['modified_by']) ) {
            $this->excel->getProperties()->setLastModifiedBy(trim($params['modified_by']));
        }

        if ( isset($params['title']) ) {
            $this->excel->getProperties()->setTitle(trim($params['title']));
        }

        if ( isset($params['subject']) ) {
            $this->excel->getProperties()->setSubject(trim($params['subject']));
        }

        if ( isset($params['description']) ) {
            $this->excel->getProperties()->setDescription(trim($params['description']));
        } else {
            $this->excel->getProperties()->setDescription('Data/godzina wygenerowania: %data_godzina_wygenerowania%');
        }

        if ( isset($params['keywords']) && is_array($params['keywords']) && count($params['keywords']) > 0 ) {
            $this->excel->getProperties()->setKeywords(implode(', ', $params['keywords']));
        }

        if ( isset($params['category']) ) {
            $this->excel->getProperties()->setCategory(trim($params['category']));
        }

        $this->switchToSheet(0);

        if ( isset($params['sheet_title']) ) {
            $sheetTitle = trim($params['sheet_title']);

            if (mb_strlen($sheetTitle) > 31 ) {
                $sheetTitle = mb_substr($sheetTitle, 0, 31);
            }

            $this->sheet->setTitle($sheetTitle);
        }
    }

    /**
     * Turns on displaying additional information
     */
    public function debugOn()
    {
        $this->debug = true;
    }

    /**
     * Turns off displaying additional information
     */
    public function debugOff()
    {
        $this->debug = false;
    }

    /**
     * Inserts given value into the current cell and moves the internal pointer into the next cell on the right
     *
     * @param string $value Value
     * @param integer $colspan How many cells (including the current one) you wish to span (default value: 1 = no spanning)
     * @param array $options Optional parameters (@see applyCellOptions())
     *
     * @return XTable
     */
    public function addValue($value, $colspan = 1, $options = array())
    {
        if ( is_object($value) ) {
            $value = $value->__toString();
        }

        // Most of PHPExcel's methods operate on Excel coordinates (for example: A3, B7, AC10, etc.)
        // Because I operate on cell and rows numbers I have to convert given numbers into coordinates
        $cellCoords = $this->toCoords();

        $this->displayDebugMessage('Do komórki o wspolrzednych (r:'. $this->row .', c:'. $this->coll .'), aka. '. $cellCoords .' wpisuje wartosc "'. $value .'"<br>');

        $result = $this->sheet->setCellValueByColumnAndRow($this->coll, $this->row, $value);

        // If there is a newline char in the cell value (ALT+Enter) Excel automatically enables word wrap
        // That's why I do the exact same thing :)
        if ( strpos($value, "\n") !== false ) {
            $this->sheet->getStyle($cellCoords)->getAlignment()->setWrapText(true);
        }

        // Optional cell parametrization
        $this->applyCellOptions($cellCoords, $options);

        // If I should merge cells...
        if ( $colspan > 1 ) {
            $this->sheet->mergeCells($cellCoords .':'. $this->toCoords(null, $this->coll+$colspan-1));
        }

        // Increase the internal pointer
        $this->coll += $colspan;

        $this->bumpMaxColl();

        return $this;
    }

    /**
     * Skips given number of cells (moves the internal pointer by given amount)
     *
     * Default value: 1
     *
     * @param int $num How many columns should I skip (how far should I move the pointer)
     *
     * @return XTable
     */
    public function skip($num = 1)
    {
        if ( !is_numeric($num) || $num < 0 ) {
            $num = 1;
        }

        $this->coll += $num;

        $this->bumpMaxColl();

        return $this;
    }

    /**
     * Moves the internal pointer into the first column of the next row (equivalent of \n\r, ie. new line + carriage return) and resets the row formatting options
     *
     * return XTable
     */
    public function nextRow()
    {
        $this->row++;
        $this->coll = 0;
        $this->clearRowOptions();

        return $this;
    }

    /**
     * Sets the current row formatting options
     *
     * Using this will forget the currently set row formatting (if there was any)
     *
     * Available options:
     *   >
     *   > bgcolor - background color; format: RRGGBB (without # in front!)
     *   > height - row height
     *   > bold - should the text be bold by default (true) or not (false; default)?
     *
     * @param array $options Options
     *
     * @return XTable
     */
    public function setRowOptions(array $options = [])
    {
        $this->clearRowOptions();

        // The 'bgcolor' and 'bold' options are applied to row cells
        $this->currentRowOptions = $options;

        // The 'height' option is applied to the row
        $this->applyRowOptions($options);

        return $this;
    }

    /**
     * Converts given column number into its Excel coordinate (A, B, C, etc.)
     *
     * If no column number is given, the internal pointer's column will be used
     *
     * Constraint: only A-ZZ column range is currently supported (~700 columns)
     *
     * @param integer $coll Column number (optional)
     *
     * @return string Column coordinate
     */
    public function columnNumberToColumnName($coll = null)
    {
        // @TODO try using ExcelUtils::convertNumberToColumnName()?

        if ( is_null($coll) ) {
            $coll = $this->coll;
        }

        $firstDigit = floor($coll / 26);
        $secondDigit = $coll - ($firstDigit * 26);

        $collName = '';
        if ( $firstDigit > 0 ) {
            $collName .= chr(64+$firstDigit);
        }

        $collName .= chr(65+$secondDigit);

        return $collName;
    }

    /**
     * Converts given column and row numbers into its Excel coordinates (A1, B2, C5, etc.)
     *
     * If no column number is given, the internal pointer's column will be used
     *
     * If no row number is given, the internal pointer's row will be used
     *
     * @param integer $row Row number (optional)
     * @param integer $coll Column number (optional)
     *
     * @return string Cell coordinates
     */
    public function toCoords($row = null, $coll = null)
    {
        if ( is_null($row) ) {
            $row = $this->row;
        }

        return $this->columnNumberToColumnName($coll) . $row;
    }

    /**
     * Converts given Excel cell coordinates into cell and row numbers
     *
     * @return array Row number (row) and column number (coll)
     */
    public function toColumnAndRow($value)
    {
        return array(
            'coll' => (ord($value{0})-65),
            'row'  => $value{1}
        );
    }

    /**
     * Get the row number where the internal pointer is currently at
     *
     * @return integer Current row number (1..)
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * Get the column number where the internal pointer is currently at
     *
     * @return integer Current column number (0..)
     */
    public function getColl()
    {
        return $this->coll;
    }

    /**
     * Returns the current sheet
     *
     * @return PHPExcel_Worksheet Sheet
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * Returns the worksheet
     *
     * @return PHPExcel Worksheet
     */
    public function getExcel()
    {
        return $this->excel;
    }

    /**
     * Applies options to the given cell
     *
     * Options set here override the options set for the cell's row
     *
     * Available parameters:
     *   > bgcolor - background color; format: RRGGBB (without # in front!)
     *   > font-size - text size
     *   > bold - true if the text should be bold (default: false)
     *   > italic - true if the text should be italic (default: false)
     *   > underline - true if the text should be underlined (default: false)
     *   > strikethrough - true if the text should be struck through (default: false)
     *   > subscript - (default: false)
     *   > superscript - (default: false)
     *   > wrap - true if the text wrapping should be enabled (default: false)
     *   > text-align - horizontal cell contents align:
     *     - center
     *     - left
     *     - right
     *     - justify
     *     - general
     *     - centerContinous
     *   > vertical-align - vertical cell contents align:
     *     - bottom
     *     - center
     *     - justify
     *     - top
     *   > borders (array):
     *      > top, bottom, left, right (array):
     *          > border-style - given border style:
     *              - dashDot
     *              - dashDotDot
     *              - dashed
     *              - dotted
     *              - double
     *              - hair
     *              - medium
     *              - mediumDashDot
     *              - mediumDashDotDot
     *              - mediumDashed
     *              - none
     *              - slantDashDot
     *              - thick
     *              - thin (default value)
     *          > border-color - given border color; format: RRGGBB (without # in front!)
     *                           default value: 000000 (black)
     *   > comment (array):
     *      > lines (array) - lines of comment text and its parameters:
     *          > text (string) - line of text
     *          > options (array) - given text line parameters (optional):
     *              > bold (boolean)
     *      > options (array) - entire comment parameters (optional):
     *          > height (float) - comment field height (default: 55.5 points)
     *          > width (float) - comment field width (default: 96 punktów)
     *   > hyperlink (boolean) - convert cell value into a hyperlink (causion!)
     *
     * @param string $cell_coords Cell coordinates
     * @param array $options Parameters
     */
    private function applyCellOptions($cell_coords, $options)
    {
        $style = $this->sheet->getStyle($cell_coords);

        $options = array_merge(array_merge($options, $this->globalCellOptions), $this->currentRowOptions);

        if ( isset($options['bgcolor']) && !is_null($options['bgcolor']) ) {
            $style->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
            $style->getFill()->getStartColor()->setARGB('FF'. $options['bgcolor']);
        }

        if ( isset($options['bold']) ) {
            $style->getFont()->setBold((boolean)$options['bold']);
        }

        if ( isset($options['italic']) ) {
            $style->getFont()->setItalic((boolean)$options['italic']);
        }

        if ( isset($options['underline']) ) {
            $style->getFont()->setUnderline((boolean)$options['underline']);
        }

        if ( isset($options['strikethrough']) ) {
            $style->getFont()->setStrikethrough((boolean)$options['strikethrough']);
        }

        if ( isset($options['subscript']) ) {
            $style->getFont()->setSubScript((boolean)$options['subscript']);
        }

        if ( isset($options['superscript']) ) {
            $style->getFont()->setSuperScript((boolean)$options['superscript']);
        }

        if ( isset($options['wrap']) ) {
            $style->getAlignment()->setWrapText((boolean)$options['wrap']);
        }

        if ( isset($options['font-size']) ) {
            $style->getFont()->setSize((integer)$options['font-size']);
        } elseif ( !empty($this->defaultFontSize) ) {
            $style->getFont()->setSize((integer)$this->defaultFontSize);
        }

        // center, left, right, justify, general, centerContinous
        if ( isset($options['text-align']) && !is_null($options['text-align']) ) {
            $style->getAlignment()->setHorizontal($options['text-align']);
        }

        // bottom, center, justify, top
        if ( isset($options['vertical-align']) && !is_null($options['vertical-align']) ) {
            $style->getAlignment()->setVertical($options['vertical-align']);
        }

        if ( array_key_exists('borders', $options) && !empty($options['borders']) ) {
            //echo 'dla komórki '. $cell_coords .' ustawiasz obramowanie<br>';

            $bordersOptions = array();

            if ( array_key_exists('top', $options['borders']) && !empty($options['borders']['top']) ) {
                $bordersOptions['top'] = $this->extractBordersOptions($options['borders']['top']);
            }

            if ( array_key_exists('bottom', $options['borders']) && !empty($options['borders']['bottom']) ) {
                $bordersOptions['bottom'] = $this->extractBordersOptions($options['borders']['bottom']);
            }

            if ( array_key_exists('left', $options['borders']) && !empty($options['borders']['left']) ) {
                $bordersOptions['left'] = $this->extractBordersOptions($options['borders']['left']);
            }

            if ( array_key_exists('right', $options['borders']) && !empty($options['borders']['right']) ) {
                $bordersOptions['right'] = $this->extractBordersOptions($options['borders']['right']);
            }

            if ( !empty($bordersOptions) ) {
                $style->getBorders()->applyFromArray($bordersOptions);
            }
        }

        if ( isset($options['comment']) && is_array($options['comment']) && array_key_exists('lines', $options['comment']) && count($options['comment']['lines']) > 0 ) {
            $comment = $this->sheet->getComment($cell_coords);

            foreach ( $options['comment']['lines'] as $commentTextAndOptions ) {
                if ( !array_key_exists('text', $commentTextAndOptions) || '' === trim($commentTextAndOptions['text']) || null === $commentTextAndOptions['text'] ) {
                    break;
                }

                $commentText = $comment->getText()->createTextRun($commentTextAndOptions['text']);

                if ( array_key_exists('options', $commentTextAndOptions) && is_array($commentTextAndOptions['options']) ) {
                    $this->applyCellCommentTextOptions($commentText, $commentTextAndOptions['options']);
                }
                $comment->getText()->createTextRun("\r\n");
            }

            if ( array_key_exists('options', $options['comment']) && is_array($options['comment']['options']) ) {
                $this->applyCellCommentOptions($comment, $options['comment']['options']);
            }
        }

        if ( isset($options['hyperlink']) && true === (boolean)$options['hyperlink'] ) {
            $value = $this->sheet->getCell($cell_coords)->getValue();

            $isEmail = filter_var($value, FILTER_VALIDATE_EMAIL);
            $isUrl   = filter_var($value, FILTER_VALIDATE_URL);

            if ( $isEmail ) {
                $this->sheet->getCell($cell_coords)->getHyperlink()->setUrl('mailto:'. $value);
            } elseif ( $isUrl ) {
                $this->sheet->getCell($cell_coords)->getHyperlink()->setUrl($value);
            }
        }
    }

    private function applyCellCommentOptions(\PHPExcel_Comment $comment, array $options)
    {
        $setHeight = (array_key_exists('height', $options) && is_numeric($options['height']) && $options['height'] > 0 );

        if ( $setHeight ) {
            $height = $options['height'] .'pt';

            $comment->setHeight($height);
        }

        $setWidth = (array_key_exists('width', $options) && is_numeric($options['width']) && $options['width'] > 0 );

        if ( $setWidth ) {
            $width = $options['width'] .'pt';

            $comment->setWidth($width);
        }
    }

    private function applyCellCommentTextOptions(\PHPExcel_RichText_Run $commentText, array $options)
    {
        $makeBold = (array_key_exists('bold', $options) && true === $options['bold']);

        if ( $makeBold ) {
            $commentText->getFont()->setBold($makeBold);
        }
    }

    /**
     * Resets (forgets) the current row formatting options
     *
     * @see applyRowOptions()
     */
    private function clearRowOptions()
    {
        $this->currentRowOptions = [];
    }

    /**
     * Applies options to current row
     *
     * @param array $options Parametry
     */
    private function applyRowOptions($options = array())
    {
        if ( isset($options['height']) && is_numeric($options['height']) ) {
            $this->sheet->getRowDimension($this->row)->setRowHeight($options['height']);
        }
    }

    /**
     * Automatically sets width on all cells in use in the current sheet based on their contents (it's not very precise)
     *
     * @return XTable
     */
    public function autoSizeAllColumns()
    {
        for ( $c = 0; $c < $this->maxColl; $c++ ) {
            $collDim = $this->sheet->getColumnDimension($this->columnNumberToColumnName($c));

            if ( $collDim->getWidth() == -1 ) {
                $collDim->setAutoSize(true);
            }
        }

        return $this;
    }

    /**
     * Automatically sets the height on the current row based on its contents
     *
     * Height is calculated based on the following formula:
     *   given number of pixels * MAX(numer of lines of text in all the row's cells)
     *
     * @param integer $row_height How much height (in pixels) does every line of text get?
     */
    public function autoSizeRow($row_height = 12)
    {
        $rowCellsLinesCount = array();
        for ( $c = 0; $c < $this->coll; $c++ ) {
            $cellValue = $this->sheet->getCellByColumnAndRow($c, $this->row)->getValue();
            $rowCellsLinesCount[] = $this->countNewlines($cellValue);
        }

        $maxLinesInRow = ( max($rowCellsLinesCount) + 1 );

        if ( $maxLinesInRow > 1 ) {
            $this->sheet->getRowDimension($this->row)->setRowHeight($row_height * $maxLinesInRow);
        }
    }

    /**
     * Returns the number of newline chars in the given string
     *
     * @param string $string
     *
     * @return integer
     */
    private function countNewlines($string)
    {
        return ( strpos($string, "\n") !== false ? count(explode("\n", $string)) : 0 );
    }

    /**
     * Resets (forgets) the remembered range start and ending coordinates
     */
    public function resetRange()
    {
        unset(
            $this->rangeStart,
            $this->rangeEnd
        );
    }

    /**
     * Begins (opens) a range
     *
     * If no coordinates for the range start point are given it will start at the current internal pointer location
     *
     * @param string $coords Cell coordinates (optional)
     */
    public function startRange($coords = null)
    {
        $this->rangeStart = ( is_null($coords) ? $this->toCoords() : $coords );
    }

    /**
     * Ends (finishes) a range
     *
     * If no coordinates for the range ending point are given it will end at the current internal pointer location
     *
     * If the range wasn't opened before, calling this method will be interpreted as call to open a range, not close it
     *
     * @param string $coords Cell coordinates (optional)
     */
    public function endRange($coords = null)
    {
        // If I know the range starting point...
        if ( $this->rangeOpened() ) {
            $this->rangeEnd = ( is_null($coords) ? $this->toCoords(null, $this->coll-1) : $coords );
        }
        // If the range wasn't opened, open it now
        else {
            $this->startRange($coords);
        }
    }

    /**
     * Applies options to the given cell
     *
     * Available parameters:
     *   > border-style (string):
     *     - dashDot
     *     - dashDotDot
     *     - dashed
     *     - dotted
     *     - double
     *     - hair
     *     - medium
     *     - mediumDashDot
     *     - mediumDashDotDot
     *     - mediumDashed
     *     - none
     *     - slantDashDot
     *     - thick
     *     - thin (default value)
     *
     *   > border-color - border color; format: RRGGBB (without # in front!)
     *     default value: 000000 (black)
     *
     *   > bordering-type (string) - bordering type (aka. which borders are we using :))
     *     - allborders (all borders)
     *     - outline (default) (only outer borders)
     *     - inside (only interior borders)
     *     - vertical (only vertical borders)
     *     - horizontal (only horizontal borders)
     *
     *   > font (array)
     *     - bold (boolean)
     *     - italic (boolean)
     *     - size (integer)
     *     - underline (boolean)
     *     - strikethrough (boolean)
     *     - subscript (boolean)
     *     - superscript (boolean)
     *
     * WARNING: calling this method clears set range starting and ending points (if there were any set)
     *
     * If we only know the starting point of range, the current cell will become its ending point
     *
     * If range is unknown, given options are not applied.
     *
     * @param array $options Parameters
     */
    public function setRangeOptions($options = array())
    {
        // Close the range on the current cell if it's still open
        if ( !$this->rangeClosed() ) {
            $this->endRange();
        }

        // Bail if there is no range
        if ( !$this->rangeSet() ) {
            return;
        }

        if ( empty($options) ) {
            return;
        }

        $rangeOptions = array();

        if (
            ( array_key_exists('border-style', $options) && !empty($options['border-style']) ) ||
            ( array_key_exists('border-color', $options) && !empty($options['border-color']) ) ||
            ( array_key_exists('bordering-type', $options) && !empty($options['bordering-type']) )
        ) {
            $bordersOptions = $this->extractBordersOptions($options);

            $rangeOptions['borders'] = array(
                $bordersOptions['bordering-type'] => array(
                    'style' => $bordersOptions['style'],
                    'color' => $bordersOptions['color'],
                )
            );
        }

        if ( array_key_exists('font', $options) && !empty($options['font']) ) {
            $rangeOptions['font'] = $options['font'];
        }

        $this->sheet->getStyle($this->rangeStart .':'. $this->rangeEnd)->applyFromArray($rangeOptions);

        $this->resetRange();
    }

    /**
     * Tells if we know the range starting point
     *
     * @return boolean
     */
    private function rangeOpened()
    {
        return isset($this->rangeStart);
    }

    /**
     * Tells if we know the range ending point
     *
     * @return boolean
     */
    private function rangeClosed()
    {
        return isset($this->rangeEnd);
    }

    /**
     * Tells if we do have a range set (we know its starting and ending points)
     *
     * @return boolean
     */
    private function rangeSet()
    {
        return ( $this->rangeOpened() && $this->rangeClosed() );
    }

    protected function bumpMaxColl()
    {
        if ( $this->coll > $this->maxColl ) {
            $this->maxColl = $this->coll;
        }
    }

    /**
     * Applies options to the given column
     *
     * WARNING: calling this method will reset the currently set formatting for the given column
     *
     * see applyCollumOptions()
     *
     * @param integer $column_no Column number (1..n)
     * @param array $options Parameters
     *
     * @return XTable
     */
    public function setColumnOptions($column_no, $options = array())
    {
        $this->clearColumnOptions($column_no);
        $this->applyColumnOptions($column_no, $options);

        return $this;
    }

    /**
     * Resets the currently set formatting for the given column
     *
     * @see applyColumnOptions()
     *
     * @param integer $column_no Column number (1..n)
     */
    private function clearColumnOptions($column_no)
    {
        if ( is_numeric($column_no) && $column_no > 0 ) {
            $this->sheet->getColumnDimension($this->columnNumberToColumnName($column_no-1))->setWidth(-1);
        }
    }

    /**
     * Set options to the given column
     *
     * Available parameters:
     *   > width - column width
     *
     * @param integer $column_no Column number (1..n)
     * @param array $options Parameters
     */
    private function applyColumnOptions($column_no, $options = array())
    {
        if ( is_numeric($column_no) && $column_no > 0 ) {
            if ( isset($options['width']) && is_numeric($options['width']) && $options['width'] > 0 ) {
                $this->sheet->getColumnDimension($this->columnNumberToColumnName($column_no-1))->setWidth($options['width']);
            }
        }
    }

    private function extractBordersOptions($options)
    {
        $out = array(
            'style'          => PHPExcel_Style_Border::BORDER_THIN,
            'color'          => array('argb' => 'FF000000'),
            'bordering-type' => 'outline',
        );

        if ( array_key_exists('border-style', $options) ) {
            switch ( $options['border-style'] ) {
                case 'dashDot':
                    $out['style'] = PHPExcel_Style_Border::BORDER_DASHDOT;
                break;

                case 'dashDotDot':
                    $out['style'] = PHPExcel_Style_Border::BORDER_DASHDOTDOT;
                break;

                case 'dashed':
                    $out['style'] = PHPExcel_Style_Border::BORDER_DASHED;
                break;

                case 'dotted':
                    $out['style'] = PHPExcel_Style_Border::BORDER_DOTTED;
                break;

                case 'double':
                    $out['style'] = PHPExcel_Style_Border::BORDER_DOUBLE;
                break;

                case 'hair':
                    $out['style'] = PHPExcel_Style_Border::BORDER_HAIR;
                break;

                case 'medium':
                    $out['style'] = PHPExcel_Style_Border::BORDER_MEDIUM;
                break;

                case 'mediumDashDot':
                    $out['style'] = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT;
                break;

                case 'mediumDashDotDot':
                    $out['style'] = PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT;
                break;

                case 'mediumDashed':
                    $out['style'] = PHPExcel_Style_Border::BORDER_MEDIUMDASHED;
                break;

                case 'none':
                    $out['style'] = PHPExcel_Style_Border::BORDER_NONE;
                break;

                case 'slantDashDot':
                    $out['style'] = PHPExcel_Style_Border::BORDER_SLANTDASHDOT;
                break;

                case 'thick':
                    $out['style'] = PHPExcel_Style_Border::BORDER_THICK;
                break;

                case 'thin':
                    $out['style'] = PHPExcel_Style_Border::BORDER_THIN;
                break;

                /*default:
                    $out['style'] = PHPExcel_Style_Border::BORDER_THIN;
                break;*/
            }
        }

        if ( array_key_exists('bordering-type', $options) ) {
            switch ( $options['bordering-type'] ) {
                case 'allborders':
                case 'outline':
                case 'inside':
                case 'vertical':
                case 'horizontal':
                case 'top':
                case 'bottom':
                case 'left':
                case 'right':
                    $out['bordering-type'] = $options['bordering-type'];
                break;

                /*default:
                    $out['bordering-type'] = 'outline';
                break;*/
            }
        }

        if ( array_key_exists('border-color', $options) ) {
            if ( isset($options['border-color']) ) {
                $out['color'] = array('argb' => 'FF'. $options['border-color']);
            }
        }

        return $out;
    }

    /**
     * Sets page header (applies to printing)
     *
     * @param string $content Header content
     * @param boolean $odd true (default) if this header should be used only on odd pages, false otherwise (warning: all pages will get the exact same header if enableOddEvenHeaderAndFooter(true) is not called!)
     *
     * @return boolean
     */
    public function setSheetHeader($content, $odd = true)
    {
        if ( empty($content) ) {
            return false;
        }

        if ( $odd ) {
            $this->sheet->getHeaderFooter()->setOddHeader($content);
        }
        else {
            $this->sheet->getHeaderFooter()->setEvenHeader($content);
        }

        return true;
    }

    /**
     * Sets page footer (applies to printing)
     *
     * @param string $content Footer content
     * @param boolean $odd true (default) if this header should be used only on odd pages, false otherwise (warning: all pages will get the exact same header if enableOddEvenHeaderAndFooter(true) is not called!)
     *
     * @return boolean
     */
    public function setSheetFooter($content, $odd = true)
    {
        if ( empty($content) ) {
            return false;
        }

        if ( $odd ) {
            $this->sheet->getHeaderFooter()->setOddFooter($content);
        }
        else {
            $this->sheet->getHeaderFooter()->setEvenFooter($content);
        }

        return true;
    }

    /**
     * Returns the page header
     *
     * @param boolean $odd true (default) if you want the odd pages header, false - otherwise
     *
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function getSheetHeader($odd = true)
    {
        if ( $odd ) {
            return $this->sheet->getHeaderFooter()->getOddHeader();
        } else {
            return $this->sheet->getHeaderFooter()->getEvenHeader();
        }
    }

    /**
     * Returns the page footer
     *
     * @param boolean $odd true (default) if you want the odd pages footer, false - otherwise
     *
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function getSheetFooter($odd = true)
    {
        if ( $odd ) {
            return $this->sheet->getHeaderFooter()->getOddFooter();
        } else {
            return $this->sheet->getHeaderFooter()->getEvenFooter();
        }
    }

    /**
     * Turns on separation of headers and footers on odd and even pages
     *
     * WARNING: by default this separation IS NOT ON!
     *
     * @param boolean $enable
     */
    public function enableOddEvenHeaderAndFooter($enable)
    {
        $this->sheet->getHeaderFooter()->setDifferentOddEven($enable);
    }

    /**
     * Tells if the separation of headers and footers on odd and even pages is turned on
     *
     * @return boolean
     */
    public function isOddEvenHeaderAndFooterEnabled()
    {
        return $this->sheet->getHeaderFooter()->getDifferentOddEven();
    }

    /**
     * Sets the default font size for all the worksheet cells
     *
     * WARNING: this applies only to content added AFTER this method is called!
     *
     * @param integer $size Font size
     */
    public function setDefaultFontSize($size)
    {
        if ( is_numeric($size) && $size > 1 ) {
            $this->defaultFontSize = $size;
        }
    }

    /**
     * Returns the currently set default font size (if any was set)
     *
     * @return integer|null
     */
    public function getDefaultFontSize()
    {
        return $this->defaultFontSize;
    }

    /**
     * Sets the worksheet this wrapper should work on instead of the one it created itself
     *
     * @param \PHPExcel $excel Worksheet
     *
     * @return XTable
     */
    public function loadExistingExcel(\PHPExcel $excel)
    {
        $this->excel = $excel;

        return $this;
    }

    /**
     * Adds a new sheet to the worksheet and switches to it
     *
     * Available options:
     *  > show_lines (boolean) - true if grid lines should be shown
     *
     * @param string $title Name/title of the new sheet
     * @param array $options Options (optional)
     *
     * @return XTable
     */
    public function addAndSwitchToSheet($title = null, array $options = [])
    {
        $this->excel->createSheet();

        $this->switchToSheet(++$this->sheetNo);

        if ( null !== $title ) {
            $this->excel->getActiveSheet()->setTitle($title);
        }

        if ( is_array($options) && count($options) ) {
            if ( array_key_exists('show_lines', $options) ) {
                $this->excel->getActiveSheet()->setShowGridLines((boolean)$options['show_lines']);
            }
        }

        return $this;
    }

    /**
     * Switches to the given sheet (0-...)
     *
     * @param integer $sheetNumber
     *
     * @return XTable
     */
    public function switchToSheet($sheetNumber)
    {
        $this->sheetNo = $sheetNumber;
        $this->excel->setActiveSheetIndex($this->sheetNo);
        $this->sheet = $this->excel->getActiveSheet();

        return $this;
    }

    /**
     * Sets options that will be applied to every cell of the worksheet from now on
     *
     * Each of the options applied here can be overridden by addValue() third argument
     *
     * @see AddValue()
     *
     * @param array $options Options
     *
     * @return XTable
     */
    public function setGlobalCellOptions(array $options = array())
    {
        $this->globalCellOptions = $options;

        return $this;
    }

    public function setCurrentRowOptions(array $options = array())
    {
        $this->currentRowOptions = $options;

        return $this;
    }

    private function displayDebugMessage($message)
    {
        if ( $this->debug ) {
            echo $message . PHP_EOL;
        }
    }
}