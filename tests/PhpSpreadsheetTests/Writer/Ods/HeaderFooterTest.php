<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Writer\Ods;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheetTests\Functional\AbstractFunctional;
use PHPUnit\Framework\Attributes\DataProvider;

class HeaderFooterTest extends AbstractFunctional
{
    #[DataProvider('providerHeaderText')]
    public function testHeaderText(string $header, string $expected): void
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->getHeaderFooter()->setOddHeader($header);
        $data = (new Ods\Styles(new Ods($spreadsheet)))->write();
        self::assertStringContainsString(
            '<style:master-page style:name="Mp1" style:page-layout-name="Mpm1_1"><style:header>' . $expected . '</style:header></style:master-page>',
            $data
        );
        $spreadsheet->disconnectWorksheets();
    }

    public static function providerHeaderText(): array
    {
        return [
            'centre only, no region' => ['Page &P of &N', '<text:p>Page <text:page-number/> of <text:page-count/></text:p>'],
            'regions' => [
                '&LLeft&CCentre&R&A',
                '<style:region-left><text:p>Left</text:p></style:region-left>'
                . '<style:region-center><text:p>Centre</text:p></style:region-center>'
                . '<style:region-right><text:p><text:sheet-name/></text:p></style:region-right>',
            ],
            'formatting dropped' => ['&L&B&"Arial,Bold Italic"&12&KFF0000&U&IText&B', '<style:region-left><text:p>Text</text:p></style:region-left>'],
            'fields' => [
                '&C&D &T &F &Z &Z&F &&',
                '<text:p><text:date/> <text:time/> <text:file-name text:display="name-and-extension"/>'
                . ' <text:file-name text:display="full"/> <text:file-name text:display="full"/> &amp;</text:p>',
            ],
            'line break' => ["Two\nlines", '<text:p>Two</text:p><text:p>lines</text:p>'],
            'lower case codes' => [
                '&lLeft &p&r&z&f',
                '<style:region-left><text:p>Left <text:page-number/></text:p></style:region-left>'
                . '<style:region-right><text:p><text:file-name text:display="full"/></text:p></style:region-right>',
            ],
        ];
    }

    public function testHeaderFooterPages(): void
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->getHeaderFooter()
            ->setOddHeader('&COdd')
            ->setEvenHeader('&CEven')
            ->setFirstFooter('&RFirst')
            ->setDifferentOddEven(true)
            ->setDifferentFirst(true);
        $spreadsheet->createSheet()->setTitle('Sheet2');
        $data = (new Ods\Styles(new Ods($spreadsheet)))->write();
        // the height of the header and the footer, in a page layout of the sheet that has them
        $height = '<style:header-footer-properties fo:min-height="0.2953in" fo:margin-left="0in" fo:margin-right="0in" ';
        self::assertStringContainsString(
            '<style:page-layout style:name="Mpm1"/><style:page-layout style:name="Mpm1_1">'
            . '<style:header-style>' . $height . 'fo:margin-bottom="0.0984in"/></style:header-style>'
            . '<style:footer-style>' . $height . 'fo:margin-top="0.0984in"/></style:footer-style>'
            . '</style:page-layout></office:automatic-styles>',
            $data
        );
        self::assertStringContainsString(
            '<style:master-page style:name="Mp1" style:page-layout-name="Mpm1_1">'
            . '<style:header><text:p>Odd</text:p></style:header>'
            . '<style:header-left><text:p>Even</text:p></style:header-left>'
            . '<loext:header-first/>'
            . '<style:footer/><style:footer-left/>'
            . '<loext:footer-first><style:region-right><text:p>First</text:p></style:region-right></loext:footer-first>'
            . '</style:master-page>'
            . '<style:master-page style:name="Mp2" style:page-layout-name="Mpm1"/>',
            $data
        );
        $spreadsheet->disconnectWorksheets();
    }

    public function testPrintTitles(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'before');
        $sheet->setCellValue('A3', 'title');
        $sheet->getColumnDimension('C')->setWidth(20);
        $sheet->getPageSetup()->setRowsToRepeatAtTop([3, 5]);
        $sheet->getPageSetup()->setColumnsToRepeatAtLeft(['B', 'C']);
        $data = (new Ods\Content(new Ods($spreadsheet)))->write();
        self::assertStringContainsString(
            '<office:forms/><table:table-column table:number-columns-repeated="1"/>'
            . '<table:table-header-columns><table:table-column table:number-columns-repeated="1"/>'
            . '<table:table-column table:style-name="co_0_3"/></table:table-header-columns>',
            $data
        );
        // the empty row 2 stays outside, the empty rows 4 and 5 past the data are written inside
        self::assertMatchesRegularExpression(
            '~</table:table-row><table:table-row table:number-rows-repeated="1"/>'
            . '<table:table-header-rows><table:table-row>.*title.*</table:table-row>'
            . '<table:table-row table:number-rows-repeated="2"/></table:table-header-rows></table:table>~',
            $data
        );
        $spreadsheet->disconnectWorksheets();
    }

    public function testMalformedPrintTitles(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'one');
        // a start that is no row is ignored, as is a range that ends before it starts
        $sheet->getPageSetup()->setRowsToRepeatAtTop([-1, 1]);
        $sheet->getPageSetup()->setColumnsToRepeatAtLeft(['$B', '$C']);
        $sheet = $spreadsheet->createSheet();
        $sheet->setCellValue('A1', 'two');
        $sheet->getPageSetup()->setRowsToRepeatAtTop([2, 1]);
        $sheet->getPageSetup()->setColumnsToRepeatAtLeft(['C', 'B']);

        $reloaded = $this->writeAndReload($spreadsheet, 'Ods');
        $spreadsheet->disconnectWorksheets();
        self::assertSame(2, $reloaded->getSheetCount());
        self::assertFalse($reloaded->getSheet(0)->getPageSetup()->isRowsToRepeatAtTopSet());
        self::assertSame(['B', 'C'], $reloaded->getSheet(0)->getPageSetup()->getColumnsToRepeatAtLeft());
        self::assertSame('two', $reloaded->getSheet(1)->getCell('A1')->getValue());
        self::assertFalse($reloaded->getSheet(1)->getPageSetup()->isRowsToRepeatAtTopSet());
        self::assertFalse($reloaded->getSheet(1)->getPageSetup()->isColumnsToRepeatAtLeftSet());
        $reloaded->disconnectWorksheets();
    }

    public function testRoundTrip(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray([['Name', 'Value'], ['a', 1]]);
        $sheet->getPageSetup()->setRowsToRepeatAtTop([1, 2]);
        $sheet->getPageSetup()->setColumnsToRepeatAtLeft(['A', 'A']);
        $sheet->getHeaderFooter()
            ->setOddHeader('&LLeft &P&CCentre&R&A')
            ->setOddFooter("&CPage &P of &N\nTom && Jerry")
            ->setEvenHeader('&CEven')
            ->setFirstFooter('&RFirst &F')
            ->setDifferentOddEven(true)
            ->setDifferentFirst(true);
        $spreadsheet->createSheet()->setCellValue('A1', 'plain');

        $reloaded = $this->writeAndReload($spreadsheet, 'Ods');
        $spreadsheet->disconnectWorksheets();
        $sheet = $reloaded->getSheet(0);
        self::assertSame([1, 2], $sheet->getPageSetup()->getRowsToRepeatAtTop());
        self::assertSame(['A', 'A'], $sheet->getPageSetup()->getColumnsToRepeatAtLeft());
        $headerFooter = $sheet->getHeaderFooter();
        self::assertSame('&LLeft &P&CCentre&R&A', $headerFooter->getOddHeader());
        self::assertSame("&CPage &P of &N\nTom && Jerry", $headerFooter->getOddFooter());
        self::assertTrue($headerFooter->getDifferentOddEven());
        self::assertSame('&CEven', $headerFooter->getEvenHeader());
        self::assertSame('', $headerFooter->getEvenFooter());
        self::assertTrue($headerFooter->getDifferentFirst());
        self::assertSame('', $headerFooter->getFirstHeader());
        self::assertSame('&RFirst &F', $headerFooter->getFirstFooter());

        $sheet = $reloaded->getSheet(1);
        self::assertFalse($sheet->getPageSetup()->isRowsToRepeatAtTopSet());
        self::assertFalse($sheet->getPageSetup()->isColumnsToRepeatAtLeftSet());
        self::assertSame('', $sheet->getHeaderFooter()->getOddHeader());
        self::assertFalse($sheet->getHeaderFooter()->getDifferentOddEven());
        $reloaded->disconnectWorksheets();
    }
}
