<?php declare(strict_types=1);
namespace Noctis;

use PhpOffice\PhpSpreadsheet\Comment;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Wrapper for the PHPSpreadsheet library, used to generate Excel files (BIFF/Open XML formats)
 *
 * For now only the most basic options are available.
 *
 * @author Łukasz Czejgis
 */
final class XTable
{
    private Worksheet $sheet;
    private Spreadsheet $excel;
    private int $row;       // Row number (1..)
    private int $coll;      // Column number (yes, number) (0..)
    private int $maxColl;
    private bool $debug;    // Display additional information?
    private ?string $rangeStart;
    private ?string $rangeEnd;
    private ?int $defaultFontSize;
    private int $sheetNo;   // Number of the current sheet in the worksheet
    private array $globalCellOptions = [];
    private array $currentRowOptions = [];

    /**
     * Available parameters (all are optional):
     *   > start_row - row number from which we start to fill cells (default value: 1)
     *   > start_coll - column number from which we start to fill cells (default value: 1; 1 = A, 2 = B, 3 = C, etc.)
     *   > creator - worksheet author
     *   > modified_by - name of person who last modified the worksheet
     *   > title - worksheet title
     *   > subject - worksheet subject
     *   > description - worksheet description (default value: date and time when the document was created)
     *   > keywords - worksheet keywords
     *   > category - worksheet category
     *   > sheet_title - sheet name (sheet, not worksheet!)
     */
    public function __construct(array $params = [])
    {
        $this->row  = $params['start_row'] ?? 1;
        $this->coll = $params['start_coll'] ?? 1;
        $this->maxColl = $this->coll;
        $this->debug = false;               // By default we want to be quiet be (otherwise the generated file can not be sent)
        $this->rangeStart = $this->rangeEnd = null;
        $this->defaultFontSize = null;
        $this->excel = new Spreadsheet();

        if (isset($params['creator'])) {
            $this->excel
                ->getProperties()
                ->setCreator(
                    trim($params['creator'])
                );
        }

        if (isset($params['modified_by'])) {
            $this->excel
                ->getProperties()
                ->setLastModifiedBy(
                    trim($params['modified_by'])
                );
        }

        if (isset($params['title'])) {
            $this->excel
                ->getProperties()
                ->setTitle(
                    trim($params['title'])
                );
        }

        if (isset($params['subject'])) {
            $this->excel
                ->getProperties()
                ->setSubject(
                    trim($params['subject'])
                );
        }

        if (isset($params['description'])) {
            $this->excel
                ->getProperties()
                ->setDescription(
                    trim($params['description'])
                );
        } else {
            $this->excel
                ->getProperties()
                ->setDescription('Data/godzina wygenerowania: %data_godzina_wygenerowania%');
        }

        if (isset($params['keywords']) && is_array($params['keywords']) && count($params['keywords']) > 0) {
            $this->excel
                ->getProperties()
                ->setKeywords(
                    implode(', ', $params['keywords'])
                );
        }

        if (isset($params['category'])) {
            $this->excel
                ->getProperties()
                ->setCategory(
                    trim($params['category'])
                );
        }

        $this->switchToSheet(0);

        if (isset($params['sheet_title'])) {
            $sheetTitle = trim($params['sheet_title']);

            if (mb_strlen($sheetTitle) > 31) {
                $sheetTitle = mb_substr($sheetTitle, 0, 31);
            }

            $this->sheet
                ->setTitle($sheetTitle, true);
        }
    }

    /**
     * Turns on displaying additional information
     */
    public function debugOn(): void
    {
        $this->debug = true;
    }

    /**
     * Turns off displaying additional information
     */
    public function debugOff(): void
    {
        $this->debug = false;
    }

    /**
     * Inserts given value into the current cell and moves the internal pointer into the next cell on the right
     *
     * @param int $colspan How many cells (including the current one) you wish to span (default value: 1 = no spanning)
     * @param array $options Optional parameters (@see applyCellOptions())
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addValue(string $value, int $colspan = 1, array $options = []): self
    {
        if (is_object($value)) {
            $value = $value->__toString();
        }

        // Most of PHPExcel's methods operate on Excel coordinates (for example: A3, B7, AC10, etc.)
        // Because I operate on cell and rows numbers I have to convert given numbers into coordinates
        $cellCoords = $this->toCoords();

        $this->displayDebugMessage(
            sprintf(
                'Inserting "%s" into cell with coordinates (r:%s, c:%s) aka. %s<br>',
                $value,
                (string)$this->row,
                (string)$this->coll,
                $cellCoords
            )
        );

        $this->sheet->getCellByColumnAndRow($this->coll, $this->row, true)
            ->setValue($value);

        // If there is a newline char in the cell value (ALT+Enter) Excel automatically enables word wrap
        // That's why I do the exact same thing :)
        if (strpos($value, "\n") !== false) {
            $this->sheet
                ->getStyle($cellCoords)
                ->getAlignment()
                ->setWrapText(true);
        }

        // Optional cell parametrization
        $this->applyCellOptions($cellCoords, $options);

        // If I should merge cells...
        if ($colspan > 1) {
            $this->sheet
                ->mergeCells(
                    $cellCoords .':'. $this->toCoords(null, $this->coll + $colspan - 1)
                );
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
     */
    public function skip(int $num = 1): self
    {
        if ($num < 0) {
            $num = 1;
        }

        $this->coll += $num;

        $this->bumpMaxColl();

        return $this;
    }

    /**
     * Moves the internal pointer into the first column of the next row
     * (equivalent of \n\r, ie. new line + carriage return) and
     * resets the row formatting options
     */
    public function nextRow(): self
    {
        $this->row++;
        $this->coll = 1;
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
     */
    public function setRowOptions(array $options = []): self
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
     * @param int|null $coll Column number (optional)
     *
     * @return string Column coordinate
     */
    public function columnNumberToColumnName(int $coll = null): string
    {
        // @TODO try using ExcelUtils::convertNumberToColumnName()?

        if (is_null($coll)) {
            $coll = $this->coll;
        }

        $firstDigit = floor($coll / 26);
        $secondDigit = (int)($coll - ($firstDigit * 26));

        $collName = '';
        if ($firstDigit > 0) {
            $collName .= chr(64 + $firstDigit);
        }

        $collName .= chr(65 + $secondDigit);

        return $collName;
    }

    /**
     * Converts given column and row numbers into its Excel coordinates (A1, B2, C5, etc.)
     *
     * If no column number is given, the internal pointer's column will be used
     *
     * If no row number is given, the internal pointer's row will be used
     *
     * @param int|null $row  Row number (optional)
     * @param int|null $coll Column number (optional)
     *
     * @return string Cell coordinates
     */
    public function toCoords(int $row = null, int $coll = null): string
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
    public function toColumnAndRow(string $value): array
    {
        return [
            'coll' => (ord($value[0]) - 65),
            'row'  => $value[1]
        ];
    }

    /**
     * Get the row number where the internal pointer is currently at
     */
    public function getRow(): int
    {
        return $this->row;
    }

    /**
     * Get the column number where the internal pointer is currently at
     */
    public function getColl(): int
    {
        return $this->coll;
    }

    /**
     * @return Worksheet The current worksheet
     */
    public function getSheet(): Worksheet
    {
        return $this->sheet;
    }

    public function getExcel(): Spreadsheet
    {
        return $this->excel;
    }

    /**
     * Automatically sets width on all cells in use in the current sheet
     * based on their contents (it's not very precise)
     */
    public function autoSizeAllColumns(): self
    {
        for ($c = 0; $c < $this->maxColl; $c++) {
            $collDim = $this->sheet
                ->getColumnDimension(
                    $this->columnNumberToColumnName($c), true
                );

            if ($collDim->getWidth() == -1) {
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
     * @param int $rowHeight How much height (in pixels) does every line of text get?
     */
    public function autoSizeRow(int $rowHeight = 12): void
    {
        $rowCellsLinesCount = [];
        for ($c = 1; $c < $this->coll; $c++) {
            $cellValue = $this->sheet
                ->getCellByColumnAndRow($c, $this->row, true)
                ->getValue();
            $rowCellsLinesCount[] = $this->countNewlines($cellValue);
        }

        $maxLinesInRow = max($rowCellsLinesCount) + 1;

        if ($maxLinesInRow > 1) {
            $this->sheet
                ->getRowDimension($this->row, true)
                ->setRowHeight($rowHeight * $maxLinesInRow);
        }
    }

    /**
     * Resets (forgets) the remembered range start and ending coordinates
     */
    public function resetRange(): void
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
     * @param string|null $coords Cell coordinates (optional)
     */
    public function startRange(string $coords = null): void
    {
        $this->rangeStart = is_null($coords)
            ? $this->toCoords()
            : $coords;
    }

    /**
     * Ends (finishes) a range
     *
     * If no coordinates for the range ending point are given it will end at the current internal pointer location
     *
     * If the range wasn't opened before, calling this method will be interpreted as call to open a range, not close it
     *
     * @param string|null $coords Cell coordinates (optional)
     */
    public function endRange(string $coords = null): void
    {
        // If I know the range starting point...
        if ($this->rangeOpened()) {
            $this->rangeEnd = is_null($coords)
                ? $this->toCoords(null, $this->coll - 1)
                : $coords;
        } else {    // If the range wasn't opened, open it now
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
     */
    public function setRangeOptions(array $options = []): void
    {
        // Close the range on the current cell if it's still open
        if (!$this->rangeClosed()) {
            $this->endRange();
        }

        // Bail if there is no range
        if (!$this->rangeSet()) {
            return;
        }

        if (empty($options)) {
            return;
        }

        $rangeOptions = [];
        if (
            (array_key_exists('border-style', $options) && !empty($options['border-style']))
            || (array_key_exists('border-color', $options) && !empty($options['border-color']))
            || (array_key_exists('bordering-type', $options) && !empty($options['bordering-type']))
        ) {
            $bordersOptions = $this->extractBordersOptions($options);

            $rangeOptions['borders'] = [
                $bordersOptions['bordering-type'] => [
                    'style' => $bordersOptions['style'],
                    'color' => $bordersOptions['color'],
                ]
            ];
        }

        if (array_key_exists('font', $options) && !empty($options['font'])) {
            $rangeOptions['font'] = $options['font'];
        }

        $this->sheet
            ->getStyle($this->rangeStart .':'. $this->rangeEnd)
            ->applyFromArray($rangeOptions);

        $this->resetRange();
    }

    /**
     * Applies options to the given column
     *
     * WARNING: calling this method will reset the currently set formatting for the given column
     *
     * see applyCollumOptions()
     */
    public function setColumnOptions(int $columnNo, array $options = []): self
    {
        $this->clearColumnOptions($columnNo);
        $this->applyColumnOptions($columnNo, $options);

        return $this;
    }

    /**
     * Sets page header (applies to printing)
     *
     * @param bool $odd true (default) if this header should be used only on odd pages, false otherwise
     * (warning: all pages will get the exact same header if enableOddEvenHeaderAndFooter(true) is not called!)
     */
    public function setSheetHeader(string $content, bool $odd = true): bool
    {
        if (empty($content)) {
            return false;
        }

        if ($odd) {
            $this->sheet
                ->getHeaderFooter()
                ->setOddHeader($content);
        }
        else {
            $this->sheet
                ->getHeaderFooter()
                ->setEvenHeader($content);
        }

        return true;
    }

    /**
     * Sets page footer (applies to printing)
     *
     * @param bool $odd true (default) if this header should be used only on odd pages, false otherwise
     * (warning: all pages will get the exact same header if enableOddEvenHeaderAndFooter(true) is not called!)
     */
    public function setSheetFooter(string $content, bool $odd = true): bool
    {
        if (empty($content)) {
            return false;
        }

        if ($odd) {
            $this->sheet
                ->getHeaderFooter()
                ->setOddFooter($content);
        }
        else {
            $this->sheet
                ->getHeaderFooter()
                ->setEvenFooter($content);
        }

        return true;
    }

    /**
     * Returns the page header
     *
     * @param bool $odd true (default) if you want the odd pages header, false - otherwise
     */
    public function getSheetHeader(bool $odd = true): string
    {
        if ($odd) {
            return $this->sheet
                ->getHeaderFooter()
                ->getOddHeader();
        }

        return $this->sheet
            ->getHeaderFooter()
            ->getEvenHeader();
    }

    /**
     * Returns the page footer
     *
     * @param bool $odd true (default) if you want the odd pages footer, false - otherwise
     */
    public function getSheetFooter(bool $odd = true): string
    {
        if ($odd) {
            return $this->sheet
                ->getHeaderFooter()
                ->getOddFooter();
        }

        return $this->sheet
            ->getHeaderFooter()
            ->getEvenFooter();
    }

    /**
     * Turns on separation of headers and footers on odd and even pages
     *
     * WARNING: by default this separation IS NOT ON!
     */
    public function enableOddEvenHeaderAndFooter(bool $enable): void
    {
        $this->sheet
            ->getHeaderFooter()
            ->setDifferentOddEven($enable);
    }

    /**
     * Tells if the separation of headers and footers on odd and even pages is turned on
     */
    public function isOddEvenHeaderAndFooterEnabled(): bool
    {
        return $this->sheet
            ->getHeaderFooter()
            ->getDifferentOddEven();
    }

    /**
     * Sets the default font size for all the worksheet cells
     *
     * WARNING: this applies only to content added AFTER this method is called!
     */
    public function setDefaultFontSize(int $size): void
    {
        if ($size > 1) {
            $this->defaultFontSize = $size;
        }
    }

    /**
     * Returns the currently set default font size (if any was set)
     */
    public function getDefaultFontSize(): ?int
    {
        return $this->defaultFontSize;
    }

    /**
     * Sets the worksheet this wrapper should work on instead of the one it created itself
     */
    public function loadExistingExcel(Spreadsheet $excel): self
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
     * @param string|null $title Name/title of the new sheet
     */
    public function addAndSwitchToSheet(string $title = null, array $options = []): self
    {
        $this->excel
            ->createSheet(null);

        $this->switchToSheet(++$this->sheetNo);

        if (null !== $title) {
            $this->excel
                ->getActiveSheet()
                ->setTitle($title);
        }

        if (is_array($options) && count($options)) {
            if (array_key_exists('show_lines', $options)) {
                $this->excel
                    ->getActiveSheet()
                    ->setShowGridLines((bool)$options['show_lines']);
            }
        }

        return $this;
    }

    /**
     * Switches to the given sheet (0-...)
     */
    public function switchToSheet(int $sheetNumber): self
    {
        $this->sheetNo = $sheetNumber;
        $this->excel
            ->setActiveSheetIndex($this->sheetNo);
        $this->sheet = $this->excel
            ->getActiveSheet();

        return $this;
    }

    /**
     * Sets options that will be applied to every cell of the worksheet from now on
     *
     * Each of the options applied here can be overridden by addValue() third argument
     *
     * @see AddValue()
     */
    public function setGlobalCellOptions(array $options = []): self
    {
        $this->globalCellOptions = $options;

        return $this;
    }

    public function setCurrentRowOptions(array $options = []): self
    {
        $this->currentRowOptions = $options;

        return $this;
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
     *          > width (float) - comment field width (default: 96 points)
     *   > hyperlink (boolean) - convert cell value into a hyperlink (caution!)
     */
    private function applyCellOptions(string $cellCoords, array $options): void
    {
        $style = $this->sheet
                ->getStyle($cellCoords);

        $allOptions = array_merge(
            array_merge($options, $this->globalCellOptions),
            $this->currentRowOptions
        );

        if (isset($allOptions['bgcolor']) && !is_null($allOptions['bgcolor'])) {
            $style->getFill()
                ->setFillType(Fill::FILL_SOLID);
            $style->getFill()
                ->getStartColor()
                ->setARGB('FF'. $allOptions['bgcolor']);
        }

        if (isset($allOptions['bold'])) {
            $style->getFont()
                ->setBold((bool)$allOptions['bold']);
        }

        if (isset($allOptions['italic'])) {
            $style->getFont()
                ->setItalic((bool)$allOptions['italic']);
        }

        if (isset($allOptions['underline'])) {
            $style->getFont()
                ->setUnderline((bool)$allOptions['underline']);
        }

        if (isset($allOptions['strikethrough'])) {
            $style->getFont()
                ->setStrikethrough((bool)$allOptions['strikethrough']);
        }

        if (isset($allOptions['subscript'])) {
            $style->getFont()
                ->setSubScript((bool)$allOptions['subscript']);
        }

        if (isset($allOptions['superscript'])) {
            $style->getFont()
                ->setSuperScript((bool)$allOptions['superscript']);
        }

        if (isset($allOptions['wrap'])) {
            $style->getAlignment()
                ->setWrapText((bool)$allOptions['wrap']);
        }

        if (isset($allOptions['font-size'])) {
            $style->getFont()
                ->setSize((int)$allOptions['font-size']);
        } elseif (!empty($this->defaultFontSize)) {
            $style->getFont()
                ->setSize((int)$this->defaultFontSize);
        }

        // center, left, right, justify, general, centerContinous
        if (isset($allOptions['text-align']) && !is_null($allOptions['text-align'])) {
            $style->getAlignment()
                ->setHorizontal($allOptions['text-align']);
        }

        // bottom, center, justify, top
        if (isset($allOptions['vertical-align']) && !is_null($allOptions['vertical-align'])) {
            $style->getAlignment()
                ->setVertical($allOptions['vertical-align']);
        }

        if (array_key_exists('borders', $allOptions) && !empty($allOptions['borders'])) {
            $bordersOptions = [];

            if (array_key_exists('top', $allOptions['borders']) && !empty($allOptions['borders']['top'])) {
                $bordersOptions['top'] = $this->extractBordersOptions($allOptions['borders']['top']);
            }

            if (array_key_exists('bottom', $allOptions['borders']) && !empty($allOptions['borders']['bottom'])) {
                $bordersOptions['bottom'] = $this->extractBordersOptions($allOptions['borders']['bottom']);
            }

            if (array_key_exists('left', $allOptions['borders']) && !empty($allOptions['borders']['left'])) {
                $bordersOptions['left'] = $this->extractBordersOptions($allOptions['borders']['left']);
            }

            if (array_key_exists('right', $allOptions['borders']) && !empty($allOptions['borders']['right'])) {
                $bordersOptions['right'] = $this->extractBordersOptions($allOptions['borders']['right']);
            }

            if (!empty($bordersOptions)) {
                $style->getBorders()
                    ->applyFromArray($bordersOptions);
            }
        }

        if (isset($allOptions['comment'])
            && is_array($allOptions['comment'])
            && array_key_exists('lines', $allOptions['comment'])
            && count($allOptions['comment']['lines']) > 0
        ) {
            $comment = $this->sheet
                ->getComment($cellCoords);

            foreach ($allOptions['comment']['lines'] as $commentTextAndOptions) {
                if (!array_key_exists('text', $commentTextAndOptions)
                    || '' === trim($commentTextAndOptions['text'])
                    || null === $commentTextAndOptions['text']
                ) {
                    break;
                }

                $commentText = $comment->getText()
                    ->createTextRun($commentTextAndOptions['text']);

                if (array_key_exists('options', $commentTextAndOptions)
                    && is_array($commentTextAndOptions['options'])
                ) {
                    $this->applyCellCommentTextOptions($commentText, $commentTextAndOptions['options']);
                }
                $comment->getText()->createTextRun("\r\n");
            }

            if (array_key_exists('options', $allOptions['comment'])
                && is_array($allOptions['comment']['options'])
            ) {
                $this->applyCellCommentOptions($comment, $allOptions['comment']['options']);
            }
        }

        if (isset($allOptions['hyperlink']) && true === (bool)$allOptions['hyperlink']) {
            $value = $this->sheet
                ->getCell($cellCoords, true)
                ->getValue();

            $isEmail = filter_var($value, FILTER_VALIDATE_EMAIL);
            $isUrl   = filter_var($value, FILTER_VALIDATE_URL);

            if ($isEmail) {
                $this->sheet
                    ->getCell($cellCoords, true)
                    ->getHyperlink('A1')
                    ->setUrl('mailto:'. $value);
            } elseif ($isUrl) {
                $this->sheet
                    ->getCell($cellCoords, true)
                    ->getHyperlink('A1')
                    ->setUrl($value);
            }
        }
    }

    private function applyCellCommentOptions(Comment $comment, array $options): void
    {
        $setHeight = (array_key_exists('height', $options)
            && is_numeric($options['height'])
            && $options['height'] > 0);

        if ($setHeight) {
            $comment->setHeight($options['height'] .'pt');
        }

        $setWidth = (array_key_exists('width', $options)
            && is_numeric($options['width'])
            && $options['width'] > 0);

        if ($setWidth) {
            $comment->setWidth($options['width'] .'pt');
        }
    }

    private function applyCellCommentTextOptions(Run $commentText, array $options): void
    {
        $makeBold = (array_key_exists('bold', $options) && true === $options['bold']);

        if ($makeBold) {
            $commentText->getFont()
                ->setBold($makeBold);
        }
    }

    /**
     * Resets (forgets) the current row formatting options
     *
     * @see applyRowOptions()
     */
    private function clearRowOptions(): void
    {
        $this->currentRowOptions = [];
    }

    /**
     * Applies options to current row
     *
     * @param array $options
     */
    private function applyRowOptions(array $options = []): void
    {
        if (isset($options['height']) && is_numeric($options['height'])) {
            $this->sheet
                ->getRowDimension($this->row, true)
                ->setRowHeight($options['height']);
        }
    }

    /**
     * Returns the number of newline chars in the given string
     */
    private function countNewlines(string $string): int
    {
        return strpos($string, "\n") !== false
            ? count(explode("\n", $string))
            : 0;
    }

    /**
     * Tells if we know the range starting point
     */
    private function rangeOpened(): bool
    {
        return isset($this->rangeStart);
    }

    /**
     * Tells if we know the range ending point
     */
    private function rangeClosed(): bool
    {
        return isset($this->rangeEnd);
    }

    /**
     * Tells if we do have a range set (we know its starting and ending points)
     */
    private function rangeSet(): bool
    {
        return $this->rangeOpened() && $this->rangeClosed();
    }

    private function bumpMaxColl(): void
    {
        if ($this->coll > $this->maxColl) {
            $this->maxColl = $this->coll;
        }
    }

    /**
     * Resets the currently set formatting for the given column
     *
     * @see applyColumnOptions()
     */
    private function clearColumnOptions(int $columnNo): void
    {
        if ($columnNo > 0) {
            $this->sheet
                ->getColumnDimension(
                    $this->columnNumberToColumnName($columnNo-1), true
                )
                ->setWidth(-1);
        }
    }

    /**
     * Set options to the given column
     *
     * Available parameters:
     *   > width - column width
     */
    private function applyColumnOptions(int $columnNo, array $options = []): void
    {
        if ($columnNo > 0) {
            if (isset($options['width'])
                && is_numeric($options['width'])
                && $options['width'] > 0
            ) {
                $this->sheet
                    ->getColumnDimension(
                        $this->columnNumberToColumnName($columnNo-1), true
                    )
                    ->setWidth($options['width']);
            }
        }
    }

    private function extractBordersOptions(array $options): array
    {
        $out = [
            'style'          => Border::BORDER_THIN,
            'color'          => ['argb' => 'FF000000'],
            'bordering-type' => 'outline',
        ];

        if (array_key_exists('border-style', $options)) {
            switch ($options['border-style']) {
                case 'dashDot':
                    $out['style'] = Border::BORDER_DASHDOT;
                    break;

                case 'dashDotDot':
                    $out['style'] = Border::BORDER_DASHDOTDOT;
                    break;

                case 'dashed':
                    $out['style'] = Border::BORDER_DASHED;
                    break;

                case 'dotted':
                    $out['style'] = Border::BORDER_DOTTED;
                    break;

                case 'double':
                    $out['style'] = Border::BORDER_DOUBLE;
                    break;

                case 'hair':
                    $out['style'] = Border::BORDER_HAIR;
                    break;

                case 'medium':
                    $out['style'] = Border::BORDER_MEDIUM;
                    break;

                case 'mediumDashDot':
                    $out['style'] = Border::BORDER_MEDIUMDASHDOT;
                    break;

                case 'mediumDashDotDot':
                    $out['style'] = Border::BORDER_MEDIUMDASHDOTDOT;
                    break;

                case 'mediumDashed':
                    $out['style'] = Border::BORDER_MEDIUMDASHED;
                    break;

                case 'none':
                    $out['style'] = Border::BORDER_NONE;
                    break;

                case 'slantDashDot':
                    $out['style'] = Border::BORDER_SLANTDASHDOT;
                    break;

                case 'thick':
                    $out['style'] = Border::BORDER_THICK;
                    break;

                case 'thin':
                    $out['style'] = Border::BORDER_THIN;
                    break;

                /*default:
                    $out['style'] = Border::BORDER_THIN;
                    break;*/
            }
        }

        if (array_key_exists('bordering-type', $options)) {
            switch ($options['bordering-type']) {
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

        if (array_key_exists('border-color', $options)) {
            if (isset($options['border-color'])) {
                $out['color'] = [
                    'argb' => 'FF'. $options['border-color']
                ];
            }
        }

        return $out;
    }

    private function displayDebugMessage(string $message): void
    {
        if ($this->debug) {
            echo $message . PHP_EOL;
        }
    }
}