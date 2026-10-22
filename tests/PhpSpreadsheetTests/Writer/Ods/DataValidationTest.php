<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Writer\Ods;

use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheetTests\Functional\AbstractFunctional;
use PHPUnit\Framework\Attributes\DataProvider;

class DataValidationTest extends AbstractFunctional
{
    /**
     * The conditions as LibreOffice 26.8 writes them for the same validations, converted from Xlsx.
     */
    public static function conditionProvider(): array
    {
        return [
            'list of values' => [DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '"a,b,c"', '', 'of:cell-content-is-in-list(&quot;a&quot;;&quot;b&quot;;&quot;c&quot;)'],
            'list from cells' => [DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, '$H$1:$J$1', '', 'of:cell-content-is-in-list([.$H$1:.$J$1])'],
            'whole between' => [DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_BETWEEN, '1', '10', 'of:cell-content-is-whole-number() and cell-content-is-between(1,10)'],
            'whole without comparison' => [DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_BETWEEN, '', '', 'of:cell-content-is-whole-number()'],
            'whole not between' => [DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_NOTBETWEEN, '1', '10', 'of:cell-content-is-whole-number() and cell-content-is-not-between(1,10)'],
            'whole not equal' => [DataValidation::TYPE_WHOLE, DataValidation::OPERATOR_NOTEQUAL, '3', '', 'of:cell-content-is-whole-number() and cell-content()!=3'],
            'decimal greater' => [DataValidation::TYPE_DECIMAL, DataValidation::OPERATOR_GREATERTHAN, '1.5', '', 'of:cell-content-is-decimal-number() and cell-content()&gt;1.5'],
            'decimal formula' => [DataValidation::TYPE_DECIMAL, DataValidation::OPERATOR_LESSTHANOREQUAL, 'H5*2', '', 'of:cell-content-is-decimal-number() and cell-content()&lt;=[.H5]*2'],
            'text length less' => [DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_LESSTHAN, '5', '', 'of:cell-content-text-length()&lt;5'],
            'text length without comparison' => [DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_BETWEEN, '', '', ''],
            'text length between' => [DataValidation::TYPE_TEXTLENGTH, DataValidation::OPERATOR_BETWEEN, '2', '5', 'of:cell-content-text-length-is-between(2,5)'],
            'date' => [DataValidation::TYPE_DATE, DataValidation::OPERATOR_GREATERTHANOREQUAL, '45000', '', 'of:cell-content-is-date() and cell-content()&gt;=45000'],
            'time' => [DataValidation::TYPE_TIME, DataValidation::OPERATOR_EQUAL, '0.5', '', 'of:cell-content-is-time() and cell-content()=0.5'],
            'custom' => [DataValidation::TYPE_CUSTOM, DataValidation::OPERATOR_BETWEEN, 'AND(C3>1,C3<5)', '', 'of:is-true-formula(AND([.C3]&gt;1;[.C3]&lt;5))'],
            'other sheet' => [DataValidation::TYPE_LIST, DataValidation::OPERATOR_BETWEEN, 'Sheet2!$A$1:$A$3', '', "of:cell-content-is-in-list(['Sheet2'.\$A\$1:.\$A\$3])"],
        ];
    }

    #[DataProvider('conditionProvider')]
    public function testCondition(string $type, string $operator, string $formula1, string $formula2, string $expected): void
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->getDataValidation('C3')
            ->setType($type)
            ->setOperator($operator)
            ->setFormula1($formula1)
            ->setFormula2($formula2);

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        if ($expected === '') {
            self::assertStringContainsString('<table:content-validation table:name="val1" table:base-cell-address', $content);
        } else {
            self::assertStringContainsString('table:condition="' . $expected . '"', $content);
        }
        $spreadsheet->disconnectWorksheets();
    }

    public function testMessagesAndCells(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle("Bob's data");
        $sheet->setCellValue('A1', 'Colour');
        $sheet->setCellValue('B2', 'red');
        $sheet->setDataValidation('B2:B4', (new DataValidation())
            ->setType(DataValidation::TYPE_LIST)
            ->setFormula1('"red,green"')
            ->setShowDropDown(true)
            ->setAllowBlank(true)
            ->setShowInputMessage(true)
            ->setPromptTitle('Colour')
            ->setPrompt("Pick a colour\r\nfrom the list")
            ->setShowErrorMessage(true)
            ->setErrorStyle(DataValidation::STYLE_WARNING)
            ->setErrorTitle('Unknown')
            ->setError("Not a\rcolour"));
        // a prompt alone, and a validation without anything to write
        $sheet->getDataValidation('D1')->setPrompt('Notes');
        $sheet->getDataValidation('E1');

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        self::assertStringContainsString(
            '<table:content-validations>'
            . '<table:content-validation table:name="val1" table:base-cell-address="\'Bob\'\'s data\'.D1">'
            . '<table:help-message table:display="false"><text:p>Notes</text:p></table:help-message>'
            . '</table:content-validation>'
            . '<table:content-validation table:name="val2" table:condition="of:cell-content-is-in-list(&quot;red&quot;;&quot;green&quot;)"'
            . ' table:allow-empty-cell="true" table:display-list="unsorted" table:base-cell-address="\'Bob\'\'s data\'.B2">'
            . '<table:help-message table:title="Colour" table:display="true"><text:p>Pick a colour</text:p><text:p>from the list</text:p></table:help-message>'
            . '<table:error-message table:message-type="warning" table:title="Unknown" table:display="true"><text:p>Not a</text:p><text:p>colour</text:p></table:error-message>'
            . '</table:content-validation>'
            . '</table:content-validations>',
            $content
        );
        // getDataValidationCollection() lists single cells first; the cells of D1 and B2:B4, with a value or empty, below the data too
        self::assertSame(1, preg_match('~<table:table table:name="Bob\'s data".*?</table:table>~', $content, $matches));
        $rows = preg_replace(['~ table:style-name="[^"]*"~', '~<text:p>.*?</text:p>~'], '', $matches[0]);
        self::assertStringContainsString(
            '<table:table-row><table:table-cell office:value-type="string"></table:table-cell>'
            . '<table:table-cell table:number-columns-repeated="2"/><table:table-cell table:content-validation-name="val1"/></table:table-row>'
            . '<table:table-row><table:table-cell/><table:table-cell table:content-validation-name="val2" office:value-type="string"></table:table-cell></table:table-row>'
            . '<table:table-row table:number-rows-repeated="2"><table:table-cell/><table:table-cell table:content-validation-name="val2"/></table:table-row>'
            . '</table:table>',
            $rows
        );
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * An error alert alone is written; a validation without a type ('' when Xlsx has no type attribute) and nothing else is not.
     */
    public function testErrorAlertAlone(): void
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->getDataValidation('A1')->setShowErrorMessage(true);
        $spreadsheet->getActiveSheet()->getDataValidation('B1')->setType('');

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        self::assertStringContainsString(
            '<table:content-validations>'
            . '<table:content-validation table:name="val1" table:base-cell-address="\'Worksheet\'.A1">'
            . '<table:error-message table:message-type="stop" table:display="true"/>'
            . '</table:content-validation>'
            . '</table:content-validations>',
            $content
        );
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * Excel leaves out errorStyle when it is stop, and the Xlsx Reader keeps it as ''.
     */
    public function testErrorStyleMissingInXlsx(): void
    {
        $spreadsheet = (new Xlsx())->load('tests/data/Reader/XLSX/issue.3863.xlsx');
        self::assertSame('', $spreadsheet->getActiveSheet()->getDataValidation('A1')->getErrorStyle());

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        self::assertStringContainsString('<table:error-message table:message-type="stop"', $content);
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * Overlapping validations (Excel does not allow them) are clipped: the one starting further left wins.
     */
    public function testOverlap(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('B3', 7);
        $sheet->setDataValidation('A1:B5', (new DataValidation())->setType(DataValidation::TYPE_WHOLE)->setFormula1('1'));
        $sheet->setDataValidation('B3:C3', (new DataValidation())->setType(DataValidation::TYPE_DECIMAL)->setFormula1('1'));

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        $rows = (string) preg_replace(['~ table:style-name="[^"]*"~', '~<text:p>.*?</text:p>~'], '', $content);
        self::assertStringContainsString(
            '<table:table-row><table:table-cell table:content-validation-name="val1"/>'
            . '<table:table-cell table:content-validation-name="val1" office:value-type="float" office:value="7"></table:table-cell>'
            . '<table:table-cell table:content-validation-name="val2"/></table:table-row>',
            $rows
        );
        $spreadsheet->disconnectWorksheets();
    }

    /**
     * Cells with and without a value, joined back into one range.
     */
    public function testRoundTripWithoutDropDown(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('C5', 'b');
        $sheet->setDataValidation('B4:D8', (new DataValidation())
            ->setType(DataValidation::TYPE_LIST)
            ->setFormula1('"a,b"')
            ->setShowDropDown(false));

        $content = (new Ods($spreadsheet))->getWriterPartContent()->write();
        self::assertStringContainsString('table:display-list="no"', $content);
        $reloaded = $this->writeAndReload($spreadsheet, 'Ods');
        $spreadsheet->disconnectWorksheets();

        $validations = $reloaded->getActiveSheet()->getDataValidationCollection();
        self::assertSame(['B4:D8'], array_keys($validations));
        self::assertFalse($validations['B4:D8']->getShowDropDown());
        self::assertSame('"a,b"', $validations['B4:D8']->getFormula1());
        self::assertSame('b', $reloaded->getActiveSheet()->getCell('C5')->getValue());
        $reloaded->disconnectWorksheets();
    }

    public function testRoundTrip(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 5);
        $sheet->setDataValidation('A1:A3 C5', (new DataValidation())
            ->setType(DataValidation::TYPE_WHOLE)
            ->setOperator(DataValidation::OPERATOR_NOTBETWEEN)
            ->setFormula1('1')
            ->setFormula2('B1*2')
            ->setShowInputMessage(true)
            ->setPromptTitle('Number')
            ->setPrompt('Not 1 to twice B')
            ->setShowErrorMessage(true)
            ->setErrorStyle(DataValidation::STYLE_INFORMATION)
            ->setError('Are you sure?'));
        $sheet->setDataValidation('E2:F1048576', (new DataValidation())
            ->setType(DataValidation::TYPE_TEXTLENGTH)
            ->setOperator(DataValidation::OPERATOR_LESSTHANOREQUAL)
            ->setFormula1('10'));

        $reloaded = $this->writeAndReload($spreadsheet, 'Ods');
        $spreadsheet->disconnectWorksheets();

        $validations = $reloaded->getActiveSheet()->getDataValidationCollection();
        self::assertSame(['A1:A3 C5', 'E2:F1048576'], array_keys($validations));
        $validation = $validations['A1:A3 C5'];
        self::assertSame(DataValidation::TYPE_WHOLE, $validation->getType());
        self::assertSame(DataValidation::OPERATOR_NOTBETWEEN, $validation->getOperator());
        self::assertSame('1', $validation->getFormula1());
        self::assertSame('B1*2', $validation->getFormula2());
        self::assertFalse($validation->getAllowBlank());
        self::assertTrue($validation->getShowInputMessage());
        self::assertSame('Number', $validation->getPromptTitle());
        self::assertSame('Not 1 to twice B', $validation->getPrompt());
        self::assertTrue($validation->getShowErrorMessage());
        self::assertSame(DataValidation::STYLE_INFORMATION, $validation->getErrorStyle());
        self::assertSame('', $validation->getErrorTitle());
        self::assertSame('Are you sure?', $validation->getError());
        $validation = $validations['E2:F1048576'];
        self::assertSame(DataValidation::TYPE_TEXTLENGTH, $validation->getType());
        self::assertSame(DataValidation::OPERATOR_LESSTHANOREQUAL, $validation->getOperator());
        self::assertSame('10', $validation->getFormula1());
        self::assertFalse($validation->getShowInputMessage());
        self::assertSame(5, $reloaded->getActiveSheet()->getCell('A1')->getValue());
        $reloaded->disconnectWorksheets();
    }
}
