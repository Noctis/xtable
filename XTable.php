<?php
namespace OpenNexus\Bundle\PlatformBundle\Util;

use \PHPExcel;
use \PHPExcel_Style_Border;
use \PHPExcel_Style_Fill;
use \PHPExcel_Writer_Excel2007;
use \PHPExcel_Writer_Excel5;


/**
 * Wrapper do pluginu sfPhpExcel, służącego do generowania plików Excel (BIFF/Open XML)
 *
 * Na chwilę obecną obsługiwane jest minimum opcji
 *
 * @author Łukasz Czejgis
 */
class XTable
{
    private $sheet;     // Arkusz
    private $excel;     // Skoroszyt
    //private $row;       // Numer wiersza (1..)
    //private $coll;      // Numer (tak, numer) kolumny (0..)
    //private $maxColl;
    //private $debug;     // Wyświetlać dodatkowe informacje?
    //private $bgcolor;   // Kolor tła wiersza
    private $rangeStart;
    private $rangeEnd;
    private $defaultFontSize;
    private $sheetNo;   // numer bieżącego arkusza w skoroszycie
    private $globalCellOptions;
    private $currentRowOptions;

    /**
     * Parametry:
     *   > start_row (opcjonalny) - numer wiersza, od którego zaczynamy wypełniać komórki (domyślnie: 1)
     *   > start_coll (opcjonalny) - numer kolumny, od której zaczynamy wypełniać komórki (domyślnie: 0; 0 = A, 1 = B, 2 = C, itd.)
     *   > creator (opcjonalny) - autor skoroszyt
     *   > modified_by (opcjonalny) - kto ostatnio modyfikował skoroszyt
     *   > title (opcjonalny) - tytuł skoroszytu
     *   > subject (opcjonalny) - temat skoroszytu
     *   > description (opcjonalny) - opis skoroszytu (domyślnie: data/godzina wygenerowania dokumentu)
     *   > keywords (opcjonalny) - słowa kluczowe skoroszytu
     *   > category (opcjonalny) - kategoria skoroszytu
     *   > sheet_title (opcjonalny) - nazwa arkusza (arkusza, nie skoroszytu!)
     *
     * @param array $params Parametry
     */
    public function __construct($params = array())
    {
        $this->row  = ( isset($params['start_row'])  ? $params['start_row']  : 1 );
        $this->coll = ( isset($params['start_coll']) ? $params['start_coll'] : 0 );
        $this->maxColl = $this->coll;
        $this->debug = false;               // Domyślnie chcemy być cicho (inaczej serwer nie wyśle wygenerowanego pliku)
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
        }
        else {
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
     * Włącza wyświetlanie dodatkowych informacji
     */
    public function debugOn()
    {
        $this->debug = true;
    }

    /**
     * Wyłącza wyświetlanie dodatkowych informacji
     */
    public function debugOff()
    {
        $this->debug = false;
    }

    /**
     * Wpisuje podaną wartość w bieżącą komórkę (i przesuwa wew. wskaźnik na komórkę po prawej stronie)
     *
     * @param string $value Wartość komórki
     * @param integer $colspan Ile komórek (razem z tą) idąc w prawo chcemy scalić (1 = brak scalenia)
     * @param array $options Opcjonalne parametry (@see applyCellOptions())
     */
    public function addValue($value, $colspan = 1, $options = array())
    {
        if ( is_object($value) ) {
            $value = $value->__toString();
        }

        // Większość funkcji phpExcel operuje na współrzędnych Excelowych (np. A3, B7, AC10, itp.)
        // Ponieważ ja, wewnętrznie, operuję numerami (kolumn i wierszy), muszę na potrzeby chwili dokonać konwersji
        $cellCoords = $this->toCoords();

        if ( $this->debug ) {
            echo 'Do komórki o wspolrzednych (r:'. $this->row .', c:'. $this->coll .'), aka. '. $cellCoords .' wpisuje wartosc "'. $value .'"<br>';
        }

        $result = $this->sheet->setCellValueByColumnAndRow($this->coll, $this->row, $value);

        // Jeżeli w komórce użyty zostanie znak nowego wierwsza (ALT+Enter) Excel automatycznie włącza zawijanie wierszy
        // Dlatego też ja robię dokładnie to samo
        if ( strpos($value, "\n") !== false ) {
            $this->sheet->getStyle($cellCoords)->getAlignment()->setWrapText(true);
        }

        // Ewentualna parametryzaja komórki
        $this->applyCellOptions($cellCoords, $options);

        // Jeżeli mam dokonać scalenia komórek...
        if ( $colspan > 1 ) {
            $this->sheet->mergeCells($cellCoords .':'. $this->toCoords(null, $this->coll+$colspan-1));
        }

        // Przesunięcie wew. wskaźnika
        $this->coll += $colspan;

        $this->bumpMaxColl();

        return $this;
    }

    /**
     * Przesuwa wewnętrzny wskaźnik o wskazaną ilość kolumn w prawo (domyślnie: 1)
     *
     * @param int $num O ile kolumn przesunąć wskaźnik
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
     * Przesuwa wew. wskaźnik do nowego wiersza, na pierwszą kolumnę (odpowiednik \n\r)
     * oraz czyści (resetuje) formatowanie wiersza
     */
    public function nextRow()
    {
        $this->row++;
        $this->coll = 0;
        $this->clearRowOptions();

        return $this;
    }

    /**
     * Ustawia opcje formatowania bieżącego wiersza
     *
     * Użycie tej opcji powoduje wyczyszczenie dotychczas ustawionego formatowania wierwsza (jeżeli było ustawione)
     *
     * Obsługiwane parametry:
     *   > bgcolor - kolor tła, w formacie: RRGGBB (bez # z przodu!)
     *   > height - wysokość wiersza
     *   > bold - pogrubić tekst (true) czy nie (false; domyślnie)?
     *
     * @param array $options Opcje
     */
    public function setRowOptions(array $options = [])
    {
        $this->clearRowOptions();

        // Opcje 'bgcolor' i 'bold' są stosowane do komórek wiersza
        $this->currentRowOptions = $options;

        // Opcja 'height' jest stosowana do wiersza
        $this->applyRowOptions($options);

        return $this;
    }

    /**
     * Tłumaczy podany numer kolumny na jej Excelową nazwę (współrzędną)
     *
     * Jeżeli nie zostanie podany numer kolumny, użyta zostanie wartość wew. wskaźnika
     *
     * Ograniczenie: obsługiwane są kolumny od A do ZZ (~700 kolumn)
     *
     * @param integer $coll Numer kolumny (opcjonalnie)
     *
     * @return string Nazwa (współrzędna) kolumny
     */
    public function columnNumberToColumnName($coll = null)
    {
        // @todo pomyśleć czy nie użyć tutaj convertNumberToColumnName() z ExcelUtils :)

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
     * Tłumaczy podany numer kolumny i wiersza na Excelową nazwę (współrzędne) komórki
     *
     * Jeżeli nie zostanie podany numer kolumny, użyta zostanie wartość wew. wskaźnika
     *
     * Jeżeli nie zostanie podany numer wiersza, użyta zostanie wartość wew. wskaźnika
     *
     * @param integer $row Numer wiersza (opcjonalnie)
     * @param integer $coll Numer wiersza (opcjonalnie)
     * @return string Nazwa (współrzędne) komórki
     */
    public function toCoords($row = null, $coll = null)
    {
        if ( is_null($row) ) {
            $row = $this->row;
        }

        return $this->columnNumberToColumnName($coll) . $row;
    }

    /**
     * Tłumaczy podaną Excelową nazwę (współrzędne) komórki na numer wiersza i kolumny
     *
     * @return array Numer wiersza (row) i kolumny (coll)
     */
    public function toColumnAndRow($value)
    {
        return array(
            'coll' => (ord($value{0})-65),
            'row'  => $value{1}
        );
    }

    /**
     * Zwraca bieżące wskazanie wierwsza wew. wskaźnika
     *
     * @return integer Numer bieżącego wiersza (1..)
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * Zwraca bieżące wskazanie kolumny wew. wskaźnika
     *
     * @return integer Numer bieżącej kolumny (0..)
     */
    public function getColl()
    {
        return $this->coll;
    }

    /**
     * Zwraca aktywny arkusz
     *
     * @return PHPExcel_Worksheet Arkusz
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * Zwraca skoroszyt
     *
     * @return sfPhpExcel Skoroszyt
     */
    public function getExcel()
    {
        return $this->excel;
    }

    /**
     * Parametryzuje daną komórkę
     *
     * Ustawione tutaj opcje nadpisują opcje ustawione dla wiersza danej komórki
     *
     * Obsługiwane parametry:
     *   > bgcolor - kolor tła, w formacie: RRGGBB (bez # z przodu!)
     *   > font-size - rozmiar tekstu
     *   > bold - pogrubić tekst (true) czy nie (false; domyślnie)?
     *   > italic - kursywa (true) czy nie (false; domyślnie)
     *   > underline - podkreślenie (true) czy nie (false; domyślnie)
     *   > strikethrough - przekreślenie (true) czy nie (false; domyślnie)
     *   > subscript - indeks dolny (true) czy nie (false; domyślnie)
     *   > superscript - indeks górny (true) czy nie (false; domyślnie)
     *   > wrap - czy włączyć zawijanie tekstu w tej komórce (true) czy nie (false; domyślnie)
     *   > text-align - poziome (horyzontalne) wyrównanie:
     *     - center (do środka)
     *     - left (do lewej)
     *     - right (do prawej)
     *     - justify (wyjustowanie)
     *     - general
     *     - centerContinous
     *   > vertical-align - pionowe (wertykalne) wyrównanie:
     *     - bottom (do dołu)
     *     - center (do środka)
     *     - justify (wyjustowanie)
     *     - top (do góry)
     *   > borders (array) - obramowanie:
     *      > top, bottom, left, right (array):
     *          > border-style - styl obramowania:
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
     *              - thin (domyślnie)
     *          > border-color - kolor obramowania, w formacie: RRGGBB (bez # z przodu!)
     *                           domyślnie: 000000 (czerń)\
     *   > comment (array) - komentarz:
     *      > lines (array) - linie tekstu będącego treścią komentarza:
     *          > text (string) - linijka tekstu
     *          > options (array) - opcje dot. tej linijki tekstu (opcjonalnie):
     *              > bold (boolean) - pogrubienie
     *      > options (array) - opcje dot. całego komentarza (opcjonalnie):
     *          > height (float) - wysokość pola komentarza (domyślnie 55.5 punktu)
     *          > width (float) - szerokość pola komentarza (domyślnie: 96 punktów)
     *   > hyperlink (boolean) - traktować zawartość komórki jako adres WWW lub e-mail (ostrożnie!)
     *
     * @param integer $cell_coords Nazwa (współrzędne) Excelowe komórki
     * @param array $options Parametry
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
     * Czyści ustawione formatowanie wiersza
     *
     * @see applyRowOptions()
     */
    private function clearRowOptions()
    {
        $this->currentRowOptions = [];
    }

    /**
     * Parametryzuje bieżący wiersz
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
     * Ustawia "na oko" szerokości wszystkich kolumn jakie obecnie istnieją w arkuszu
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
     * Ustawia "na oko" wysokość bieżącego wiersza
     *
     * Wysokość obliczana jest wg formuły:
     *   dana liczba pikseli * MAX(ilość wierszy tekstu we wszystkich komórkach wiersza)
     *
     * Działanie to wykonywane jest na pojedynczym wierszu, według wartości wew. wskaźnika (numer wiersza)
     *
     * @param integer $row_height Ile pikseli wysokości dostaje każdy wiersz tekstu (patrz: wyżej)?
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
     * Zwraca ilość znaków nowego wiersza w podanym ciągu znaków
     *
     * @param string $string Badany ciąg znaków
     * @return integer
     */
    private function countNewlines($string)
    {
        return ( strpos($string, "\n") !== false ? count(explode("\n", $string)) : 0 );
    }

    /**
     * Czyści zapamiętane dane zakresu (adres początku i adres końca)
     */
    public function resetRange()
    {
        unset(
            $this->rangeStart,
            $this->rangeEnd
        );
    }

    /**
     * Otwiera (rozpoczyna) zakres
     *
     * Jeżeli nie zostaną podane współrzędne komórki, od której ma się rozpoczynać zakres,
     * zostanie on otwarty na bieżącej komórce (wskazywanej przez wewnętrzny wskaźnik)
     *
     * @param string $coords Excelowa nazwa (współrzędne) komórki (opcjonalnie)
     */
    public function startRange($coords = null)
    {
        $this->rangeStart = ( is_null($coords) ? $this->toCoords() : $coords );
    }

    /**
     * Zamyka (kończy) zakres
     *
     * Jeżeli nie zostaną podane współrzędne komórki, na której ma się kończyć zakres,
     * zostanie on zamknięty na bieżącej komórce (wskazywanej przez wewnętrzny wskaźnik)
     *
     * Jeżeli zakres nie został otwarty, użycie tej metody zostanie zinterpretowane
     * jako otwarcie zakresu, a nie zamknięcie
     *
     *
     * @param string $coords Excelowa nazwa (współrzędne) komórki (opcjonalnie)
     */
    public function endRange($coords = null)
    {
        // Jeżeli znam początek zakresu...
        if ( $this->rangeOpened() ) {
            $this->rangeEnd = ( is_null($coords) ? $this->toCoords(null, $this->coll-1) : $coords );
        }
        // Jeżeli zakres nie został otwarty, otwórz go zamiast zamykać
        else {
            $this->startRange($coords);
        }
    }

    /**
     * Parametryzuje ustalony zakres komórek
     *
     * Obsługiwane parametry:
     *   > border-style - styl obramowania:
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
     *     - thin (domyślnie)
     *
     *   > border-color - kolor obramowania, w formacie: RRGGBB (bez # z przodu!)
     *     domyślnie: 000000 (czerń)
     *
     *   > bordering-type - rodzaj obramowania (lub, które krawędzi ramkujemy :))
     *     - allborders (wszystkie krawędzie)
     *     - outline (domyślnie) (tylko górna, dolna, lewa i prawa krawędź, bez wewnętrznych)
     *     - inside (tylko wewnętrzne krawędzie)
     *     - vertical (tylko pionowe krawędzie)
     *     - horizontal (tylko poziome krawędzie)
     *
     *   > font - czcionka
     *     - bold (pogrubienie); boolean
     *     - italic (kursywa); boolean
     *     - size (rozmiar); integer
     *     - underline (podkreślenie); boolean
     *     - strikethrough (przekreślenie); boolean
     *     - subscript (indeks dolny); boolean
     *     - superscript (indeks górny); boolean
     *
     * UWAGA: Użycie tej metody czyści zapamiętane wartości początku i końca zakresu (zapominamy o zakresie)
     *
     * Jeżeli znamy tylko początek zakresu, za jego koniec przyjęta zostaje
     * bieżąca, wskazywana przez wewnętrzny wskaźnik, komórka
     *
     * Jeżeli zakres nie jest znany, nic nie jest parametryzowane.
     *
     * @param array $options Parametry
     * @return void
     */
    public function setRangeOptions($options = array())
    {
        // Zamknij zakres na bieżącej komórce jeżeli nie został on dotychczas jawnie zamknięty
        if ( !$this->rangeClosed() ) {
            $this->endRange();
        }

        // Ewakuacja jeżeli zakres nie jest zadeklarowany
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
     * Mówi czy znamy współrzędne początku zakresu
     *
     * @return boolean true jeżeli znamy, false - w przeciwnym wypadku
     */
    private function rangeOpened()
    {
        return isset($this->rangeStart);
    }

    /**
     * Mówi czy znamy współrzędne końca zakresu
     *
     * @return true jeżeli znamy, false - w przeciwnym wypadku
     */
    private function rangeClosed()
    {
        return isset($this->rangeEnd);
    }

    /**
     * Mówi czy mamy zadeklarowany zakres (znamy współrzędne jego początku i końca)
     *
     * @return boolean true jeżeli mamy, false - w przeciwnym wypadku
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
     * Ustawia opcje formatowania bieżącej kolumny
     *
     * Użycie tej opcji powoduje wyczyszczenie dotychczas ustawionego formatowania kolumny (jeżeli było ustawione)
     *
     * @see applyCollumOptions()
     *
     * @param integer $column_no Numer kolumny (1..n)
     * @param array $options Opcje
     */
    public function setColumnOptions($column_no, $options = array())
    {
        $this->clearColumnOptions($column_no);
        $this->applyColumnOptions($column_no, $options);

        return $this;
    }

    /**
     * Czyści ustawione formatowanie kolumny
     *
     * @see applyColumnOptions()
     *
     * @param integer $column_no Numer kolumny (1..n)
     */
    private function clearColumnOptions($column_no)
    {
        if ( is_numeric($column_no) && $column_no > 0 )
        {
            $this->sheet->getColumnDimension($this->columnNumberToColumnName($column_no-1))->setWidth(-1);
        }
    }

    /**
     * Parametryzuje bieżącą kolumnę
     *
     * Obsługiwane parametry:
     *   > width - szerokość kolumny
     *
     * @param integer $column_no Numer kolumny (1..n)
     * @param array $options Parametry
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

        if ( array_key_exists('bordering-type', $options) )
        {
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
     * Ustawia nagłówek arkusza
     *
     * @param string $content Treść nagłówka
     * @param boolean $odd true (domyślnie) jeśli ma to być nagłówek nieparzystych stron, false - w przeciwnym wypadku (uwaga: wszystkie strony otrzymają ten sam nagłówek jeżeli nie zostanie użyta metoda enableOddEvenHeaderAndFooter(true)!)
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
     * Ustawia stopkę arkusza
     *
     * @param string $content Treść nagłówka
     * @param boolean $odd true (domyślnie) jeśli ma to być nagłówek nieparzystych stron, false - w przeciwnym wypadku (uwaga: wszystkie strony otrzymają ten sam nagłówek jeżeli nie zostanie użyta metoda enableOddEvenHeaderAndFooter(true)!)
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
     * Zwraca nagłówek arkusza
     *
     * @param boolean $odd true (domyślnie) jeśli ma to być nagłówek nieparzystych stron, false - w przeciwnym wypadku
     *
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function getSheetHeader($odd = true)
    {
        if ( $odd ) {
            return $this->sheet->getHeaderFooter()->getOddHeader();
        }
        else {
            return $this->sheet->getHeaderFooter()->getEvenHeader();
        }
    }

    /**
     * Zwraca stopkę arkusza
     *
     * @param boolean $odd true (domyślnie) jeśli ma to być stopka nieparzystych stron, false - w przeciwnym wypadku
     *
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function getSheetFooter($odd = true)
    {
        if ( $odd ) {
            return $this->sheet->getHeaderFooter()->getOddFooter();
        }
        else {
            return $this->sheet->getHeaderFooter()->getEvenFooter();
        }
    }

    /**
     * Włącza rozróżnianie nagłówków i stopek stron parzystych i nieparzystych
     *
     * (domyślnie to rozróżnienie NIE JEST WŁĄCZONE)
     *
     * @param boolean $enable
     */
    public function enableOddEvenHeaderAndFooter($enable)
    {
        $this->sheet->getHeaderFooter()->setDifferentOddEven($enable);
    }

    /**
     * Mówi czy jest obecnie włączone rozróżnianie nagłówków i stron patrzystych i nieparzystych
     *
     * @return boolean
     */
    public function isOddEvenHeaderAndFooterEnabled()
    {
        return $this->sheet->getHeaderFooter()->getDifferentOddEven();
    }

    /**
     * Ustawia domyślną wielkość czcionki dla wszystkich komórek arkusza
     *
     * (UWAGA: zmiana obowiązuje od momentu użycia tej opcji, tj. NIE działa retroaktywnie)
     *
     * @param integer $size Rozmiar czcionki
     */
    public function setDefaultFontSize($size)
    {
        if ( is_numeric($size) && $size > 1 ) {
            $this->defaultFontSize = $size;
        }
    }

    /**
     * Zwraca obowiązującą obecnie domyślną wielkość czcionki (o ile została ustawiona)
     *
     * @return integer|null
     */
    public function getDefaultFontSize()
    {
        return $this->defaultFontSize;
    }

    /**
     * Przełącza klasę na działanie na danym skoroszycie Excela
     *
     * @param \PHPExcel $excel Skoroszyt
     *
     * @return \OpenNexus\Bundle\PlatformBundle\Util\XTable
     */
    public function loadExistingExcel(\PHPExcel $excel)
    {
        $this->excel = $excel;

        return $this;
    }

    /**
     * Dodaje do skoroszytu nowy arkusz i przełącza na niego działanie klasy
     *
     * Obsługiwane opcje:
     *  > show_lines (boolean) - pokazuj linie siatki?
     *
     * @param string $title Tytuł/nazwa nowego arkusza
     * @param array $options Opcje
     *
     * @return \OpenNexus\Bundle\PlatformBundle\Util\XTable
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
     * Przełącza działanie klasy na wskazany arkusz skoroszytu (0-...)
     *
     * @param integer $sheetNumber
     *
     * @return \OpenNexus\Bundle\PlatformBundle\Util\XTable
     */
    public function switchToSheet($sheetNumber)
    {
        $this->sheetNo = $sheetNumber;
        $this->excel->setActiveSheetIndex($this->sheetNo);
        $this->sheet = $this->excel->getActiveSheet();

        return $this;
    }

    /**
     * Ustanawia opcje, które będą odtąd stosowane do każdej komórki arkusza.
     * Każda z ustawionych tutaj opcji może zostać indywidualnie nadpisana przez 3-ci argument metody addValue()
     *
     * @see AddValue()
     *
     * @param array $options
     *
     * @return \OpenNexus\Bundle\PlatformBundle\Util\XTable
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
}