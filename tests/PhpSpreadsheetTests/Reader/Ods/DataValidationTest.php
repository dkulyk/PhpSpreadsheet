<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Reader\Ods as OdsReader;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods as OdsWriter;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use ZipArchive;

class DataValidationTest extends TestCase
{
    /**
     * Saved by LibreOffice 26.8 from an Xlsx file.
     */
    public function testLibreOfficeFile(): void
    {
        $spreadsheet = (new OdsReader())->load('tests/data/Reader/Ods/DataValidation.ods');
        $validations = $spreadsheet->getSheetByNameOrThrow('Data')->getDataValidationCollection();
        $actual = [];
        foreach ($validations as $range => $validation) {
            $actual[$range] = [
                $validation->getType(),
                $validation->getOperator(),
                $validation->getFormula1(),
                $validation->getFormula2(),
                $validation->getAllowBlank(),
            ];
        }
        self::assertSame([
            'C2' => [DataValidation::TYPE_DECIMAL, DataValidation::OPERATOR_LESSTHANOREQUAL, 'B2*2', '', false],
            'D2' => [DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_LESSTHANOREQUAL, '5', '', false],
            'E2' => [DataValidation::TYPE_CUSTOM, DataValidation::OPERATOR_BETWEEN, 'ISODD(E2)', '', false],
            'F2' => [DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, 'Lists!$A$1:$A$3', '', false],
            'G2' => [DataValidation::TYPE_DATE, DataValidation::OPERATOR_GREATERTHAN, '45658', '', false],
            'H2' => [DataValidation::TYPE_TIME, DataValidation::OPERATOR_NOTBETWEEN, '0.25', '0.75', false],
            'A2:A10' => [DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"red,green,blue"', '', true],
            'B2:B10' => [DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_BETWEEN, '1', '10', false],
            'I2:I1048576' => [DataValidation::TYPE_DECIMAL, DataValidation::OPERATOR_GREATERTHAN, '0', '', false],
        ], $actual);

        $list = $validations['A2:A10'];
        self::assertTrue($list->getShowDropDown());
        self::assertTrue($list->getShowInputMessage());
        self::assertSame('Colour', $list->getPromptTitle());
        self::assertSame('Pick a colour', $list->getPrompt());
        self::assertTrue($list->getShowErrorMessage());
        self::assertSame(DataValidation::STYLE_STOP, $list->getErrorStyle());
        self::assertSame('Not a colour', $list->getErrorTitle());
        self::assertSame('Pick one from the list', $list->getError());
        // two paragraphs, and two spaces written as text:s
        self::assertSame("Whole number,  1 to 10\nSecond line", $validations['B2:B10']->getPrompt());
        self::assertSame(DataValidation::STYLE_WARNING, $validations['B2:B10']->getErrorStyle());
        self::assertSame(DataValidation::STYLE_INFORMATION, $validations['C2']->getErrorStyle());
        self::assertSame('Too big', $validations['C2']->getError());
        self::assertFalse($validations['D2']->getShowInputMessage());
        self::assertFalse($validations['D2']->getShowErrorMessage());
        self::assertSame([], $spreadsheet->getSheetByNameOrThrow('Lists')->getDataValidationCollection());
        $spreadsheet->disconnectWorksheets();
    }

    public function testReadDataOnly(): void
    {
        $reader = new OdsReader();
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load('tests/data/Reader/Ods/DataValidation.ods');
        self::assertSame([], $spreadsheet->getSheetByNameOrThrow('Data')->getDataValidationCollection());
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * Forms ODF 1.2 to 1.4 allow that LibreOffice does not write: other namespaces, spaces, attributes left out.
     */
    public static function conditionProvider(): array
    {
        return [
            'ooow' => ['table:condition="ooow:cell-content-is-whole-number() and cell-content()&gt;=3"', DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_GREATERTHANOREQUAL, '3', ''],
            'oooc' => ['table:condition="oooc:cell-content-text-length-is-not-between(1,[.B1])"', DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_NOTBETWEEN, '1', 'B1'],
            'prefix with dot and digit' => ['table:condition="my.ns-2:cell-content-text-length()&lt;3"', DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_LESSTHAN, '3', ''],
            'spaces' => ['table:condition="of:cell-content-is-decimal-number()  and  cell-content-is-between( MAX(1;2) , 5 )"', DataValidation::TYPE_DECIMAL, DataValidation::OPERATOR_BETWEEN, 'MAX(1,2)', '5'],
            'no namespace' => ['table:condition="cell-content-is-in-list(&quot;a b&quot; ; &quot;say &quot;&quot;c&quot;&quot;&quot;)"', DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"a b,say ""c"""', ''],
            'list of numbers' => ['table:condition="of:cell-content-is-in-list(1;-2.5;3E2)"', DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"1,-2.5,3E2"', ''],
            'list of strings and numbers' => ['table:condition="of:cell-content-is-in-list(&quot;a&quot;; 1 ;&quot;2&quot;)"', DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"a,1,2"', ''],
            'string with comma' => ['table:condition="of:cell-content-is-time() and cell-content-is-between(TIMEVALUE(&quot;1,5&quot;),[.B1])"', DataValidation::TYPE_TIME, DataValidation::OPERATOR_BETWEEN, 'TIMEVALUE("1,5")', 'B1'],
            'sheet name with bracket' => ["table:condition=\"of:cell-content-is-whole-number() and cell-content-is-between(['Old (1'.B1],5)\"", DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_BETWEEN, "'Old (1'!B1", '5'],
            'type without comparison' => ['table:condition="of:cell-content-is-date()"', DataValidation::TYPE_NONE, DataValidation::OPERATOR_BETWEEN, '', ''],
            'unknown' => ['table:condition="of:cell-content-is-blue()"', DataValidation::TYPE_NONE, DataValidation::OPERATOR_BETWEEN, '', ''],
            // none is the schema value, LibreOffice writes and reads no
            'display-list none' => ['table:condition="of:cell-content-is-in-list(&quot;a&quot;)" table:display-list="none"', DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"a"', '', '', [true, false, false, false]],
            'messages not displayed' => [
                'table:condition="of:cell-content-text-length()=1" table:allow-empty-cell="false"',
                DataValidation::TYPE_TEXTLENGTH,
                DataValidation::OPERATOR_EQUAL,
                '1',
                '',
                '<table:help-message table:display="false"><text:p>Help</text:p></table:help-message>'
                . '<table:error-message table:display="false"><text:p>Error</text:p></table:error-message>',
                [false, true, false, false],
            ],
        ];
    }

    /**
     * @param array{bool, bool, bool, bool} $flags allow blank, show drop-down, show input message, show error message
     */
    #[DataProvider('conditionProvider')]
    public function testConditionForms(string $attributes, string $type, string $operator, string $formula1, string $formula2, string $messages = '', array $flags = [true, true, false, false]): void
    {
        $spreadsheet = $this->loadPatched($attributes, "'Worksheet'.C3", $messages);
        $validation = $spreadsheet->getActiveSheet()->getDataValidationCollection()['C3'];
        self::assertSame($type, $validation->getType());
        self::assertSame($operator, $validation->getOperator());
        self::assertSame($formula1, $validation->getFormula1());
        self::assertSame($formula2, $validation->getFormula2());
        self::assertSame($flags, [$validation->getAllowBlank(), $validation->getShowDropDown(), $validation->getShowInputMessage(), $validation->getShowErrorMessage()]);
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * ODF references are relative to the base cell, Xlsx ones to the top left cell (C3 here).
     * LibreOffice puts the base cell at the cursor, so it may be below or right of the top left cell;
     * relative references then wrap around the sheet edge, as in Excel.
     */
    public static function baseCellProvider(): array
    {
        return [
            'above left' => ['COUNTIF([.A1:.A5];&quot;A1&quot;)&gt;[.$B$1]', '$Worksheet.$B$2', 'COUNTIF(B2:B6,"A1")>$B$1'],
            'whole column' => ['COUNTIF([.A:.A];[.D5])=1', 'Worksheet.D5', 'COUNTIF(XFD:XFD,C3)=1'],
            'whole row' => ['SUM([.1:.1])&gt;[.D5]', 'Worksheet.D5', 'SUM(1048575:1048575)>C3'],
            'cell beyond the edge' => ['[.A1]&gt;[.D5]', 'Worksheet.D5', 'XFD1048575>C3'],
            'function name like a cell' => ['LOG10([.D5])&gt;1', 'Worksheet.D5', 'LOG10(C3)>1'],
        ];
    }

    #[DataProvider('baseCellProvider')]
    public function testBaseCellAddress(string $formula, string $baseCell, string $expected): void
    {
        $spreadsheet = $this->loadPatched('table:condition="of:is-true-formula(' . $formula . ')"', $baseCell);
        $validation = $spreadsheet->getActiveSheet()->getDataValidationCollection()['C3'];
        self::assertSame($expected, $validation->getFormula1());
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * Save a validation on C3 and replace it in content.xml.
     */
    private function loadPatched(string $attributes, string $baseCell, string $messages = ''): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->getDataValidation('C3')
            ->setType(DataValidation::TYPE_WHOLE)
            ->setOperator(DataValidation::OPERATOR_EQUAL)
            ->setFormula1('1');
        $file = File::temporaryFilename();
        (new OdsWriter($spreadsheet))->save($file);
        $spreadsheet->disconnectWorksheets();

        $zip = new ZipArchive();
        self::assertTrue($zip->open($file));
        $content = (string) $zip->getFromName('content.xml');
        $validation = '<table:content-validation table:name="val1" ' . $attributes . ' table:base-cell-address="' . $baseCell . '">' . $messages . '</table:content-validation>';
        $content = (string) preg_replace_callback('~<table:content-validation .*?(/>|</table:content-validation>)~', fn (): string => $validation, $content, 1, $count);
        self::assertSame(1, $count);
        $zip->addFromString('content.xml', $content);
        $zip->close();

        $spreadsheet = (new OdsReader())->load($file);
        unlink($file);

        return $spreadsheet;
    }
}
