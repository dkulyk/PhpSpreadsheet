<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Reader\Ods;

use Composer\Pcre\Preg;
use PhpOffice\PhpSpreadsheet\Reader\Ods;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PHPUnit\Framework\TestCase;
use ZipArchive;

class HeaderFooterTest extends TestCase
{
    // Made in PhpSpreadsheet as Xlsx and saved as Ods (ODF 1.4) by LibreOffice 26.8
    private const FILE = 'tests/data/Reader/Ods/HeaderFooterPrintTitles.ods';

    public function testHeaderFooter(): void
    {
        $spreadsheet = (new Ods())->load(self::FILE);

        $headerFooter = $spreadsheet->getSheetByNameOrThrow('Data')->getHeaderFooter();
        // bold and font codes are dropped, text:span and text:s are read
        self::assertSame('&LBold &P&CCentre &N&RSheet &A  of &F &D &T &&', $headerFooter->getOddHeader());
        self::assertSame('&CPage &P of &N', $headerFooter->getOddFooter());
        self::assertTrue($headerFooter->getDifferentOddEven());
        self::assertSame('&LEven &P', $headerFooter->getEvenHeader());
        self::assertSame('', $headerFooter->getEvenFooter());
        // style:header-first, ODF 1.3
        self::assertTrue($headerFooter->getDifferentFirst());
        self::assertSame('', $headerFooter->getFirstHeader());
        self::assertSame('&RFirst &Z&F', $headerFooter->getFirstFooter());

        // paragraphs without regions
        $headerFooter = $spreadsheet->getSheetByNameOrThrow('Middle')->getHeaderFooter();
        self::assertSame("&CTwo\nlines", $headerFooter->getOddHeader());
        self::assertSame('', $headerFooter->getOddFooter());
        self::assertFalse($headerFooter->getDifferentOddEven());
        self::assertFalse($headerFooter->getDifferentFirst());

        // style:display="false"
        $headerFooter = $spreadsheet->getSheetByNameOrThrow('Plain')->getHeaderFooter();
        self::assertSame('', $headerFooter->getOddHeader());
        self::assertSame('', $headerFooter->getOddFooter());
        self::assertFalse($headerFooter->getDifferentOddEven());
        self::assertFalse($headerFooter->getDifferentFirst());
        $spreadsheet->disconnectWorksheets();
    }

    public function testPrintTitles(): void
    {
        $spreadsheet = (new Ods())->load(self::FILE);

        $pageSetup = $spreadsheet->getSheetByNameOrThrow('Data')->getPageSetup();
        self::assertSame([1, 1], $pageSetup->getRowsToRepeatAtTop());
        self::assertSame(['A', 'A'], $pageSetup->getColumnsToRepeatAtLeft());
        self::assertSame('Value', $spreadsheet->getSheetByNameOrThrow('Data')->getCell('B1')->getValue());

        // header rows and columns that do not start the sheet
        $sheet = $spreadsheet->getSheetByNameOrThrow('Middle');
        self::assertSame([2, 3], $sheet->getPageSetup()->getRowsToRepeatAtTop());
        self::assertSame(['B', 'C'], $sheet->getPageSetup()->getColumnsToRepeatAtLeft());
        self::assertSame('b', $sheet->getCell('A5')->getValue());

        $pageSetup = $spreadsheet->getSheetByNameOrThrow('Plain')->getPageSetup();
        self::assertFalse($pageSetup->isRowsToRepeatAtTopSet());
        self::assertFalse($pageSetup->isColumnsToRepeatAtLeftSet());
        $spreadsheet->disconnectWorksheets();
    }

    public function testPrintTitlesSplitByGroups(): void
    {
        // Print titles 1:5 and A:C with grouped rows 3:7 and columns B:D, made as Xlsx and saved as Ods by
        // LibreOffice 26.8, which splits table:table-header-rows and -columns at the group boundaries
        $spreadsheet = (new Ods())->load('tests/data/Reader/Ods/PrintTitlesInGroups.ods');

        $pageSetup = $spreadsheet->getSheetByNameOrThrow('Worksheet')->getPageSetup();
        self::assertSame([1, 5], $pageSetup->getRowsToRepeatAtTop());
        self::assertSame(['A', 'C'], $pageSetup->getColumnsToRepeatAtLeft());
        $spreadsheet->disconnectWorksheets();
    }

    public function testHeaderFooterForms(): void
    {
        $masterPage = '<style:master-page style:name="PageStyle_5f_Plain" style:page-layout-name="Mpm4">'
            // a first header and an even footer only: the other pages share the odd ones
            . '<style:header><text:h>Title <text:title/></text:h></style:header>'
            . '<style:header-first><text:p>First<text:line-break/>page</text:p></style:header-first>'
            . '<style:footer><text:p><text:file-name text:display="path"/></text:p></style:footer>'
            . '<style:footer-left><text:p>Even</text:p></style:footer-left>'
            . '</style:master-page>';
        $file = (string) tempnam(File::sysGetTempDir(), 'ods');
        copy(self::FILE, $file);
        $zip = new ZipArchive();
        $zip->open($file);
        $styles = (string) $zip->getFromName('styles.xml');
        $zip->addFromString('styles.xml', Preg::replace('~<style:master-page style:name="PageStyle_5f_Plain".*?</style:master-page>~s', $masterPage, $styles));
        $zip->close();

        $spreadsheet = (new Ods())->load($file);
        unlink($file);

        $headerFooter = $spreadsheet->getSheetByNameOrThrow('Plain')->getHeaderFooter();
        self::assertSame('&CTitle &F', $headerFooter->getOddHeader());
        self::assertSame('&C&Z', $headerFooter->getOddFooter());
        self::assertTrue($headerFooter->getDifferentOddEven());
        self::assertSame('&CTitle &F', $headerFooter->getEvenHeader());
        self::assertSame('&CEven', $headerFooter->getEvenFooter());
        self::assertTrue($headerFooter->getDifferentFirst());
        self::assertSame("&CFirst\npage", $headerFooter->getFirstHeader());
        self::assertSame('&C&Z', $headerFooter->getFirstFooter());
        $spreadsheet->disconnectWorksheets();
    }
}
