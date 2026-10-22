<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use DOMDocument;
use DOMElement;
use DOMNode;
use DOMText;
use PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooter;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use stdClass;

class PageSettings
{
    private const LOEXT_NS = 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0';

    private string $officeNs = '';

    private string $stylesNs = '';

    private string $stylesFo = '';

    private string $tableNs = '';

    private string $textNs = '';

    /**
     * @var string[]
     */
    private array $tableStylesCrossReference = [];

    /**
     * @var bool[] whether a table style lays its sheet out right to left
     */
    private array $tableStylesRightToLeft = [];

    /** @var mixed[] */
    private array $pageLayoutStyles = [];

    /**
     * @var string[]
     */
    private array $masterStylesCrossReference = [];

    /**
     * @var string[]
     */
    private array $masterPrintStylesCrossReference = [];

    /**
     * @var DOMElement[]
     */
    private array $masterPages = [];

    public function __construct(DOMDocument $styleDom)
    {
        $this->setDomNameSpaces($styleDom);
        $this->readPageSettingStyles($styleDom);
        $this->readStyleMasterLookup($styleDom);
    }

    private function setDomNameSpaces(DOMDocument $styleDom): void
    {
        $this->officeNs = (string) $styleDom->lookupNamespaceUri('office');
        $this->stylesNs = (string) $styleDom->lookupNamespaceUri('style');
        $this->stylesFo = (string) $styleDom->lookupNamespaceUri('fo');
        $this->tableNs = (string) $styleDom->lookupNamespaceUri('table');
        $this->textNs = (string) $styleDom->lookupNamespaceUri('text');
    }

    private function readPageSettingStyles(DOMDocument $styleDom): void
    {
        $item0 = $styleDom->getElementsByTagNameNS($this->officeNs, 'automatic-styles')->item(0);
        $styles = ($item0 === null) ? [] : $item0->getElementsByTagNameNS($this->stylesNs, 'page-layout');

        foreach ($styles as $styleSet) {
            $styleName = $styleSet->getAttributeNS($this->stylesNs, 'name');
            $pageLayoutProperties = $styleSet->getElementsByTagNameNS($this->stylesNs, 'page-layout-properties')->item(0);
            $styleOrientation = $pageLayoutProperties?->getAttributeNS($this->stylesNs, 'print-orientation');
            $styleScale = $pageLayoutProperties?->getAttributeNS($this->stylesNs, 'scale-to');
            $stylePrintOrder = $pageLayoutProperties?->getAttributeNS($this->stylesNs, 'print-page-order');
            $centered = $pageLayoutProperties?->getAttributeNS($this->stylesNs, 'table-centering');

            $marginLeft = $pageLayoutProperties?->getAttributeNS($this->stylesFo, 'margin-left');
            $marginRight = $pageLayoutProperties?->getAttributeNS($this->stylesFo, 'margin-right');
            $marginTop = $pageLayoutProperties?->getAttributeNS($this->stylesFo, 'margin-top');
            $marginBottom = $pageLayoutProperties?->getAttributeNS($this->stylesFo, 'margin-bottom');
            $header = $styleSet->getElementsByTagNameNS($this->stylesNs, 'header-style')->item(0);
            $headerProperties = $header?->getElementsByTagNameNS($this->stylesNs, 'header-footer-properties')?->item(0);
            $marginHeader = $headerProperties?->getAttributeNS($this->stylesFo, 'min-height');
            $footer = $styleSet->getElementsByTagNameNS($this->stylesNs, 'footer-style')->item(0);
            $footerProperties = $footer?->getElementsByTagNameNS($this->stylesNs, 'header-footer-properties')?->item(0);
            $marginFooter = $footerProperties?->getAttributeNS($this->stylesFo, 'min-height');

            $this->pageLayoutStyles[$styleName] = (object) [
                'orientation' => $styleOrientation ?: PageSetup::ORIENTATION_DEFAULT,
                'scale' => $styleScale ?: 100,
                'printOrder' => $stylePrintOrder,
                'horizontalCentered' => $centered === 'horizontal' || $centered === 'both',
                'verticalCentered' => $centered === 'vertical' || $centered === 'both',
                // margin size is already stored in inches, so no UOM conversion is required
                'marginLeft' => (float) ($marginLeft ?? 0.7),
                'marginRight' => (float) ($marginRight ?? 0.7),
                'marginTop' => (float) ($marginTop ?? 0.3),
                'marginBottom' => (float) ($marginBottom ?? 0.3),
                'marginHeader' => (float) ($marginHeader ?? 0.45),
                'marginFooter' => (float) ($marginFooter ?? 0.45),
            ];
        }
    }

    private function readStyleMasterLookup(DOMDocument $styleDom): void
    {
        $item0 = $styleDom->getElementsByTagNameNS($this->officeNs, 'master-styles')->item(0);
        $styleMasterLookup = ($item0 === null) ? [] : $item0->getElementsByTagNameNS($this->stylesNs, 'master-page');

        foreach ($styleMasterLookup as $styleMasterSet) {
            $styleMasterName = $styleMasterSet->getAttributeNS($this->stylesNs, 'name');
            $pageLayoutName = $styleMasterSet->getAttributeNS($this->stylesNs, 'page-layout-name');
            $this->masterPrintStylesCrossReference[$styleMasterName] = $pageLayoutName;
            $this->masterPages[$styleMasterName] = $styleMasterSet;
        }
    }

    public function readStyleCrossReferences(DOMDocument $contentDom): void
    {
        $item0 = $contentDom->getElementsByTagNameNS($this->officeNs, 'automatic-styles')->item(0);
        $styleXReferences = ($item0 === null) ? [] : $item0->getElementsByTagNameNS($this->stylesNs, 'style');

        foreach ($styleXReferences as $styleXreferenceSet) {
            $styleXRefName = $styleXreferenceSet->getAttributeNS($this->stylesNs, 'name');
            $stylePageLayoutName = $styleXreferenceSet->getAttributeNS($this->stylesNs, 'master-page-name');
            $styleFamilyName = $styleXreferenceSet->getAttributeNS($this->stylesNs, 'family');
            if (!empty($styleFamilyName) && $styleFamilyName === 'table') {
                $styleVisibility = 'true';
                foreach ($styleXreferenceSet->getElementsByTagNameNS($this->stylesNs, 'table-properties') as $tableProperties) {
                    $styleVisibility = $tableProperties->getAttributeNS($this->tableNs, 'display');
                    // rl-tb, or rl, which ODF also allows
                    $this->tableStylesRightToLeft[$styleXRefName] = str_starts_with($tableProperties->getAttributeNS($this->stylesNs, 'writing-mode'), 'rl');
                }
                $this->tableStylesCrossReference[$styleXRefName] = $styleVisibility;
            }
            if (!empty($stylePageLayoutName)) {
                $this->masterStylesCrossReference[$styleXRefName] = $stylePageLayoutName;
            }
        }
    }

    public function setVisibilityForWorksheet(Worksheet $worksheet, string $styleName): void
    {
        if (!array_key_exists($styleName, $this->tableStylesCrossReference)) {
            return;
        }

        $worksheet->setSheetState(
            $this->tableStylesCrossReference[$styleName] === 'false'
                ? Worksheet::SHEETSTATE_HIDDEN
                : Worksheet::SHEETSTATE_VISIBLE
        );
    }

    public function setRightToLeftForWorksheet(Worksheet $worksheet, string $styleName): void
    {
        if (array_key_exists($styleName, $this->tableStylesRightToLeft)) {
            $worksheet->setRightToLeft($this->tableStylesRightToLeft[$styleName]);
        }
    }

    public function setPrintSettingsForWorksheet(Worksheet $worksheet, string $styleName): void
    {
        if (!array_key_exists($styleName, $this->masterStylesCrossReference)) {
            return;
        }
        $masterStyleName = $this->masterStylesCrossReference[$styleName];

        if (!array_key_exists($masterStyleName, $this->masterPrintStylesCrossReference)) {
            return;
        }
        $this->setHeaderFooter($worksheet->getHeaderFooter(), $this->masterPages[$masterStyleName]);
        $printSettingsIndex = $this->masterPrintStylesCrossReference[$masterStyleName];

        if (!array_key_exists($printSettingsIndex, $this->pageLayoutStyles)) {
            return;
        }
        /** @var (object{orientation: string, scale: int|string, printOrder: ?string,
         * horizontalCentered: bool, verticalCentered: bool, marginLeft: float, marginRight: float, marginTop: float,
         * marginBottom: float, marginHeader: float, marginFooter: float}&stdClass) */
        $printSettings = $this->pageLayoutStyles[$printSettingsIndex];

        $worksheet->getPageSetup()
            ->setOrientation($printSettings->orientation ?? PageSetup::ORIENTATION_DEFAULT)
            ->setPageOrder($printSettings->printOrder === 'ltr' ? PageSetup::PAGEORDER_OVER_THEN_DOWN : PageSetup::PAGEORDER_DOWN_THEN_OVER)
            ->setScale((int) trim((string) $printSettings->scale, '%'))
            ->setHorizontalCentered($printSettings->horizontalCentered)
            ->setVerticalCentered($printSettings->verticalCentered);

        $worksheet->getPageMargins()
            ->setLeft($printSettings->marginLeft)
            ->setRight($printSettings->marginRight)
            ->setTop($printSettings->marginTop)
            ->setBottom($printSettings->marginBottom)
            ->setHeader($printSettings->marginHeader)
            ->setFooter($printSettings->marginFooter);
    }

    /**
     * Read the headers and footers of a master page into Excel codes, as LibreOffice exports them to Xlsx.
     */
    private function setHeaderFooter(HeaderFooter $headerFooter, DOMElement $masterPage): void
    {
        $texts = [];
        foreach ($masterPage->childNodes as $child) {
            // style:header-first is ODF 1.3, loext:header-first is what LibreOffice writes in 1.2 extended
            if (
                $child instanceof DOMElement
                && ($child->namespaceURI === $this->stylesNs || $child->namespaceURI === self::LOEXT_NS)
                && $child->getAttributeNS($this->stylesNs, 'display') !== 'false'
            ) {
                $texts[$child->localName] = $this->headerFooterText($child);
            }
        }
        $differentOddEven = isset($texts['header-left']) || isset($texts['footer-left']);
        $differentFirst = isset($texts['header-first']) || isset($texts['footer-first']);
        $headerFooter->setOddHeader($texts['header'] ?? '')
            ->setOddFooter($texts['footer'] ?? '')
            ->setDifferentOddEven($differentOddEven)
            ->setDifferentFirst($differentFirst);
        // A missing left or first header shares the header of the other pages
        if ($differentOddEven) {
            $headerFooter->setEvenHeader($texts['header-left'] ?? $texts['header'] ?? '')
                ->setEvenFooter($texts['footer-left'] ?? $texts['footer'] ?? '');
        }
        if ($differentFirst) {
            $headerFooter->setFirstHeader($texts['header-first'] ?? $texts['header'] ?? '')
                ->setFirstFooter($texts['footer-first'] ?? $texts['footer'] ?? '');
        }
    }

    private function headerFooterText(DOMElement $headerFooter): string
    {
        $text = '';
        foreach (['region-left' => '&L', 'region-center' => '&C', 'region-right' => '&R'] as $region => $code) {
            $regionElement = $headerFooter->getElementsByTagNameNS($this->stylesNs, $region)->item(0);
            $regionText = $regionElement === null ? '' : $this->paragraphsText($regionElement);
            if ($regionText !== '') {
                $text .= $code . $regionText;
            }
        }
        if ($text === '') {
            // Paragraphs without regions are centred
            $regionText = $this->paragraphsText($headerFooter);
            $text = $regionText === '' ? '' : '&C' . $regionText;
        }

        return $text;
    }

    private function paragraphsText(DOMElement $element): string
    {
        $paragraphs = [];
        foreach ($element->childNodes as $child) {
            if ($child instanceof DOMElement && $child->namespaceURI === $this->textNs && ($child->localName === 'p' || $child->localName === 'h')) {
                $paragraphs[] = $this->inlineText($child);
            }
        }

        return implode("\n", $paragraphs);
    }

    private function inlineText(DOMNode $node): string
    {
        $text = '';
        foreach ($node->childNodes as $child) {
            if ($child instanceof DOMText) {
                $text .= str_replace('&', '&&', $child->data);
            } elseif ($child instanceof DOMElement) {
                $text .= match ($child->namespaceURI === $this->textNs ? $child->localName : '') {
                    's' => str_repeat(' ', max(1, (int) $child->getAttributeNS($this->textNs, 'c'))),
                    'line-break' => "\n",
                    'page-number' => '&P',
                    'page-count' => '&N',
                    'sheet-name' => '&A',
                    'date' => '&D',
                    'time' => '&T',
                    // LibreOffice exports the document title as the file name
                    'title' => '&F',
                    'file-name' => match ($child->getAttributeNS($this->textNs, 'display')) {
                        'name', 'name-and-extension' => '&F',
                        'path' => '&Z',
                        default => '&Z&F',
                    },
                    default => $this->inlineText($child),
                };
            }
        }

        return $text;
    }
}
