<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Ods;

use Composer\Pcre\Preg;
use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter;

class Styles extends WriterPart
{
    // Excel header and footer codes of fields, and the ODF fields LibreOffice imports them as
    private const HEADER_FOOTER_FIELDS = [
        'P' => ['text:page-number', ''],
        'N' => ['text:page-count', ''],
        'A' => ['text:sheet-name', ''],
        'D' => ['text:date', ''],
        'T' => ['text:time', ''],
        'F' => ['text:file-name', 'name-and-extension'],
        // LibreOffice Calc has no path-only file name, it imports &Z, and &Z&F, as the full path
        'Z' => ['text:file-name', 'full'],
        'Z&F' => ['text:file-name', 'full'],
    ];

    /**
     * Write styles.xml to XML format.
     *
     * @return string XML Output
     */
    public function write(): string
    {
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // Content
        $objWriter->startElement('office:document-styles');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
        $objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
        $objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
        $objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
        $objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
        $objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
        $objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
        $objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
        $objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
        $objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
        $objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
        $objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
        $objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
        $objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
        $objWriter->writeAttribute('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
        $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $objWriter->writeAttribute('xmlns:tableooo', 'http://openoffice.org/2009/table');
        $objWriter->writeAttribute('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
        $objWriter->writeAttribute('xmlns:loext', 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0');
        $objWriter->writeAttribute('office:version', '1.2');

        $objWriter->writeElement('office:font-face-decls');
        $objWriter->startElement('office:styles');
        $defaultStyle = $this->getParentWriter()
            ->getSpreadsheet()
            ->getDefaultStyle();
        $objWriter->startElement('style:default-style');
        $objWriter->writeAttribute('style:family', 'table-cell');
        $writer2 = new Cell\Style($objWriter);
        $writer2->writeTextProperties($defaultStyle);
        $objWriter->endElement(); // style:default-style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:name', 'Default');
        $objWriter->writeAttribute('style:family', 'table-cell');
        $writer2->writeCellProperties($defaultStyle);
        $objWriter->endElement(); // style:style 'Default' table-cell
        $objWriter->endElement(); // office:styles
        $worksheets = iterator_to_array($this->getParentWriter()->getSpreadsheet()->getWorksheetIterator());
        $objWriter->startElement('office:automatic-styles');
        $objWriter->startElement('style:page-layout');
        $objWriter->writeAttribute('style:name', 'Mpm1');
        $objWriter->endElement(); // style:page-layout
        // A sheet with a header or a footer gets its own page layout with their height:
        // without it LibreOffice gives a header the whole page; these are its defaults
        foreach ($worksheets as $sheetIndex => $worksheet) {
            $shown = array_filter([
                'header-style' => $this->isShown($worksheet->getHeaderFooter(), 'header'),
                'footer-style' => $this->isShown($worksheet->getHeaderFooter(), 'footer'),
            ]);
            if ($shown === []) {
                continue;
            }
            $objWriter->startElement('style:page-layout');
            $objWriter->writeAttribute('style:name', 'Mpm1_' . ($sheetIndex + 1));
            foreach (array_keys($shown) as $name) {
                $objWriter->startElement("style:$name");
                $objWriter->startElement('style:header-footer-properties');
                $objWriter->writeAttribute('fo:min-height', '0.2953in');
                $objWriter->writeAttribute('fo:margin-left', '0in');
                $objWriter->writeAttribute('fo:margin-right', '0in');
                $objWriter->writeAttribute($name === 'header-style' ? 'fo:margin-bottom' : 'fo:margin-top', '0.0984in');
                $objWriter->endElement();
                $objWriter->endElement();
            }
            $objWriter->endElement(); // style:page-layout
        }
        $objWriter->endElement(); // office:automatic-styles
        $objWriter->startElement('office:master-styles');
        // One master page per sheet, which carries its headers and footers
        foreach ($worksheets as $sheetIndex => $worksheet) {
            $headerFooter = $worksheet->getHeaderFooter();
            $objWriter->startElement('style:master-page');
            $objWriter->writeAttribute('style:name', Cell\Style::MASTER_PAGE_PREFIX . ($sheetIndex + 1));
            $hasHeaderFooter = $this->isShown($headerFooter, 'header') || $this->isShown($headerFooter, 'footer');
            $objWriter->writeAttribute('style:page-layout-name', $hasHeaderFooter ? 'Mpm1_' . ($sheetIndex + 1) : 'Mpm1');
            $this->writeHeaderFooter($objWriter, 'header', $headerFooter);
            $this->writeHeaderFooter($objWriter, 'footer', $headerFooter);
            $objWriter->endElement(); //style:master-page
        }
        $objWriter->endElement(); //office:master-styles
        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * The texts of the odd, even and first pages.
     *
     * @return array{string, string, string}
     */
    private function texts(HeaderFooter $headerFooter, string $name): array
    {
        return $name === 'header'
            ? [$headerFooter->getOddHeader(), $headerFooter->getEvenHeader(), $headerFooter->getFirstHeader()]
            : [$headerFooter->getOddFooter(), $headerFooter->getEvenFooter(), $headerFooter->getFirstFooter()];
    }

    /**
     * Whether a sheet has a header (or a footer) on any page.
     */
    private function isShown(HeaderFooter $headerFooter, string $name): bool
    {
        [$odd, $even, $first] = $this->texts($headerFooter, $name);

        return $odd !== ''
            || ($headerFooter->getDifferentOddEven() && $even !== '')
            || ($headerFooter->getDifferentFirst() && $first !== '');
    }

    /**
     * Write a header or footer: odd pages, then even (left) and first pages when they differ.
     */
    private function writeHeaderFooter(XMLWriter $objWriter, string $name, HeaderFooter $headerFooter): void
    {
        if (!$this->isShown($headerFooter, $name)) {
            return;
        }
        [$odd, $even, $first] = $this->texts($headerFooter, $name);
        $objWriter->startElement("style:$name");
        $this->writeHeaderFooterText($objWriter, $odd);
        $objWriter->endElement();
        if ($headerFooter->getDifferentOddEven()) {
            $objWriter->startElement("style:$name-left");
            $this->writeHeaderFooterText($objWriter, $even);
            $objWriter->endElement();
        }
        if ($headerFooter->getDifferentFirst()) {
            // style:header-first is ODF 1.3, LibreOffice writes loext:header-first in 1.2 extended
            $objWriter->startElement("loext:$name-first");
            $this->writeHeaderFooterText($objWriter, $first);
            $objWriter->endElement();
        }
    }

    /**
     * Write the text of a header or footer, converted from Excel codes as LibreOffice imports them.
     * Formatting codes (font, size, colour, bold and so on) and pictures are dropped.
     */
    private function writeHeaderFooterText(XMLWriter $objWriter, string $text): void
    {
        $regions = ['L' => [], 'C' => [], 'R' => []];
        $region = 'C';
        $tokens = Preg::split('/(&(?:Z&F|K.{6}|"[^"]*"?|\d+|.)?)/sui', $text, -1, PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_NO_EMPTY);
        foreach ($tokens as $token) {
            if ($token[0] !== '&') {
                $regions[$region][] = $token;

                continue;
            }
            $code = strtoupper(substr($token, 1));
            if ($code === '&') {
                $regions[$region][] = '&';
            } elseif (isset($regions[$code])) {
                $region = $code;
            } elseif (isset(self::HEADER_FOOTER_FIELDS[$code])) {
                $regions[$region][] = ['field' => $code];
            }
        }
        if ($regions['L'] === [] && $regions['R'] === []) {
            // A centre-only text needs no region, as LibreOffice writes it
            if ($regions['C'] !== []) {
                $this->writeHeaderFooterParagraphs($objWriter, $regions['C']);
            }

            return;
        }
        $names = ['L' => 'style:region-left', 'C' => 'style:region-center', 'R' => 'style:region-right'];
        foreach (array_filter($regions) as $region => $parts) {
            $objWriter->startElement($names[$region]);
            $this->writeHeaderFooterParagraphs($objWriter, $parts);
            $objWriter->endElement();
        }
    }

    /**
     * @param array<array{field: string}|string> $parts
     */
    private function writeHeaderFooterParagraphs(XMLWriter $objWriter, array $parts): void
    {
        $objWriter->startElement('text:p');
        foreach ($parts as $part) {
            if (is_array($part)) {
                [$element, $display] = self::HEADER_FOOTER_FIELDS[$part['field']];
                $objWriter->startElement($element);
                if ($display !== '') {
                    $objWriter->writeAttribute('text:display', $display);
                }
                $objWriter->endElement();

                continue;
            }
            foreach (Preg::split('/\r?\n/', $part) as $index => $line) {
                if ($index > 0) {
                    $objWriter->endElement();
                    $objWriter->startElement('text:p');
                }
                $objWriter->text($line);
            }
        }
        $objWriter->endElement();
    }
}
