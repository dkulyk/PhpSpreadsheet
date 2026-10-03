<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Functional;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Writer\Ods as OdsWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PHPUnit\Framework\Attributes\DataProvider;

class DrawingTitleTest extends AbstractFunctional
{
    private static function spreadsheet(): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        // One cell anchors the stamp, two cells the seal
        foreach ([['B2', '', 'Stamp', 'Paid stamp'], ['B10', 'D14', 'Seal', 'Company seal'], ['B20', '', 'Logo', '']] as [$coordinates, $coordinates2, $name, $title]) {
            $drawing = new Drawing();
            $drawing->setPath('tests/data/Writer/XLSX/blue_square.png')
                ->setCoordinates($coordinates)
                ->setCoordinates2($coordinates2)
                ->setName($name)
                ->setDescription("$name description")
                ->setTitle($title)
                ->setWorksheet($spreadsheet->getActiveSheet());
        }

        return $spreadsheet;
    }

    #[DataProvider('providerFormats')]
    public function testTitleSurvivesTheRoundTrip(string $format): void
    {
        $spreadsheetOld = self::spreadsheet();
        $spreadsheet = $this->writeAndReload($spreadsheetOld, $format);
        $spreadsheetOld->disconnectWorksheets();
        $titles = [];
        foreach ($spreadsheet->getActiveSheet()->getDrawingCollection() as $drawing) {
            $titles[$drawing->getName()] = [$drawing->getTitle(), $drawing->getDescription()];
        }
        ksort($titles);
        self::assertSame([
            'Logo' => ['', 'Logo description'],
            'Seal' => ['Company seal', 'Seal description'],
            'Stamp' => ['Paid stamp', 'Stamp description'],
        ], $titles);
        $spreadsheet->disconnectWorksheets();
    }

    public static function providerFormats(): array
    {
        return [['Xlsx'], ['Ods']];
    }

    public function testWrittenAsLibreOfficeWritesIt(): void
    {
        $spreadsheet = self::spreadsheet();
        $writer = new XlsxWriter($spreadsheet);
        $data = $writer->getWriterPartDrawing()->writeDrawings($spreadsheet->getActiveSheet());
        self::assertStringContainsString('name="Stamp" descr="Stamp description" title="Paid stamp"/>', $data);
        self::assertStringContainsString('name="Logo" descr="Logo description"/>', $data);

        $data = (new OdsWriter\Content(new OdsWriter($spreadsheet)))->write();
        // svg:title comes before svg:desc in the schema
        self::assertStringContainsString('<svg:title>Paid stamp</svg:title><svg:desc>Stamp description</svg:desc>', $data);
        self::assertSame(2, substr_count($data, '<svg:title>'));
        $spreadsheet->disconnectWorksheets();
    }
}
