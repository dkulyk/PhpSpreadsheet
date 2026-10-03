<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Ods;

use PhpOffice\PhpSpreadsheet\Cell\AddressRange;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Writes table:content-validations, and tells the cell writer which validation covers which cell.
 */
class DataValidations
{
    private const OPERATORS = [
        DataValidation::OPERATOR_EQUAL => '=',
        DataValidation::OPERATOR_GREATERTHAN => '>',
        DataValidation::OPERATOR_GREATERTHANOREQUAL => '>=',
        DataValidation::OPERATOR_LESSTHAN => '<',
        DataValidation::OPERATOR_LESSTHANOREQUAL => '<=',
        DataValidation::OPERATOR_NOTEQUAL => '!=',
    ];

    private const VALUE_TYPES = [
        DataValidation::TYPE_WHOLE => 'cell-content-is-whole-number()',
        DataValidation::TYPE_DECIMAL => 'cell-content-is-decimal-number()',
        DataValidation::TYPE_DATE => 'cell-content-is-date()',
        DataValidation::TYPE_TIME => 'cell-content-is-time()',
    ];

    /**
     * Cell areas per sheet index, ordered by first row: first column (0-based), first row, last column, last row, validation name.
     *
     * @var array<int, list<array{int, int, int, int, string}>>
     */
    private array $areas = [];

    /**
     * The writer asks for the rows of a sheet top down, so the areas are swept: the sheet,
     * the index of the next area to start, and the areas covering the row last asked.
     *
     * @var array{int, int, list<array{int, int, int, int, string}>}
     */
    private array $sweep = [-1, 0, []];

    public function __construct(
        private XMLWriter $objWriter,
        private Spreadsheet $spreadsheet,
        private Formula $formulaConvertor
    ) {
    }

    public function write(): void
    {
        $count = 0;
        foreach ($this->spreadsheet->getWorksheetIterator() as $sheetIndex => $worksheet) {
            foreach ($worksheet->getDataValidationCollection() as $sqref => $dataValidation) {
                if (!$this->isWorthWriting($dataValidation)) {
                    continue;
                }
                if ($count === 0) {
                    $this->objWriter->startElement('table:content-validations');
                }
                $name = 'val' . ++$count;
                $baseCell = '';
                foreach (explode(' ', str_replace('$', '', (string) $sqref)) as $range) {
                    [[$colStart, $rowStart], [$colEnd, $rowEnd]] = Coordinate::rangeBoundaries($range);
                    $this->areas[$sheetIndex][] = [$colStart - 1, $rowStart, $colEnd - 1, min($rowEnd, AddressRange::MAX_ROW), $name];
                    $baseCell = $baseCell ?: Coordinate::stringFromColumnIndex($colStart) . $rowStart;
                }
                $this->writeValidation($dataValidation, $name, $worksheet, $baseCell);
            }
        }
        if ($count > 0) {
            $this->objWriter->endElement();
        }
        foreach ($this->areas as $sheetIndex => $areas) {
            usort($areas, fn (array $a, array $b): int => $a[1] <=> $b[1]);
            $this->areas[$sheetIndex] = $areas;
        }
    }

    /**
     * The validations covering a row, as column runs ordered from the left: first column (0-based), last column, name.
     *
     * @return list<array{int, int, string}>
     */
    public function rowSegments(int $sheetIndex, int $row): array
    {
        $segments = [];
        foreach ($this->activeAreas($sheetIndex, $row) as [$colStart, , $colEnd, , $name]) {
            $segments[] = [$colStart, $colEnd, $name];
        }
        sort($segments);
        // ponytail: overlapping validations (Excel does not allow them) are clipped, the one starting further left wins
        $clipped = [];
        $nextColumn = 0;
        foreach ($segments as [$colStart, $colEnd, $name]) {
            $colStart = max($colStart, $nextColumn);
            if ($colStart <= $colEnd) {
                $clipped[] = [$colStart, $colEnd, $name];
                $nextColumn = $colEnd + 1;
            }
        }

        return $clipped;
    }

    /**
     * The first row after $row whose segments may differ from those of $row.
     */
    public function nextChange(int $sheetIndex, int $row): int
    {
        $active = $this->activeAreas($sheetIndex, $row);
        $next = $this->areas[$sheetIndex][$this->sweep[1]][1] ?? PHP_INT_MAX;
        foreach ($active as [, , , $rowEnd]) {
            $next = min($next, $rowEnd + 1);
        }

        return $next;
    }

    /**
     * The areas covering a row.
     *
     * @return list<array{int, int, int, int, string}>
     */
    private function activeAreas(int $sheetIndex, int $row): array
    {
        [$sweepSheet, $nextArea, $active] = $this->sweep;
        if ($sweepSheet !== $sheetIndex) {
            [$nextArea, $active] = [0, []];
        }
        $areas = $this->areas[$sheetIndex] ?? [];
        while (isset($areas[$nextArea]) && $areas[$nextArea][1] <= $row) {
            $active[] = $areas[$nextArea++];
        }
        $active = array_values(array_filter($active, fn (array $area): bool => $area[3] >= $row));
        $this->sweep = [$sheetIndex, $nextArea, $active];

        return $active;
    }

    /**
     * The last row any validation of the sheet covers, 0 when there is none.
     */
    public function lastRow(int $sheetIndex): int
    {
        return max([0, ...array_column($this->areas[$sheetIndex] ?? [], 3)]);
    }

    /**
     * LibreOffice skips a validation with neither a rule nor a message (sc/source/filter/xml/XMLStylesExportHelper.cxx).
     */
    private function isWorthWriting(DataValidation $dataValidation): bool
    {
        // an Xlsx validation without a type attribute has the type ''
        return !in_array($dataValidation->getType(), [DataValidation::TYPE_NONE, ''], true)
            || $dataValidation->getShowInputMessage()
            || $dataValidation->getShowErrorMessage()
            || $dataValidation->getPromptTitle() !== ''
            || $dataValidation->getPrompt() !== ''
            || $dataValidation->getErrorTitle() !== ''
            || $dataValidation->getError() !== '';
    }

    private function writeValidation(DataValidation $dataValidation, string $name, Worksheet $worksheet, string $baseCell): void
    {
        $this->objWriter->startElement('table:content-validation');
        $this->objWriter->writeAttribute('table:name', $name);
        $condition = $this->condition($dataValidation);
        if ($condition !== '') {
            $this->objWriter->writeAttribute('table:condition', $condition);
            $this->objWriter->writeAttribute('table:allow-empty-cell', $dataValidation->getAllowBlank() ? 'true' : 'false');
            if ($dataValidation->getType() === DataValidation::TYPE_LIST) {
                $this->objWriter->writeAttribute('table:display-list', $dataValidation->getShowDropDown() ? 'unsorted' : 'no');
            }
        }
        // relative references in the formulas are relative to this cell, as in Xlsx to the top left cell of the first range
        $this->objWriter->writeAttribute(
            'table:base-cell-address',
            "'" . str_replace("'", "''", $worksheet->getTitle()) . "'." . $baseCell
        );

        if ($dataValidation->getShowInputMessage() || $dataValidation->getPromptTitle() !== '' || $dataValidation->getPrompt() !== '') {
            $this->objWriter->startElement('table:help-message');
            $this->writeMessage($dataValidation->getPromptTitle(), $dataValidation->getPrompt(), $dataValidation->getShowInputMessage());
        }
        if ($dataValidation->getShowErrorMessage() || $dataValidation->getErrorTitle() !== '' || $dataValidation->getError() !== '') {
            $this->objWriter->startElement('table:error-message');
            // the error styles are named as the ODF message types: stop, warning, information; Excel leaves out stop
            $errorStyle = $dataValidation->getErrorStyle();
            $this->objWriter->writeAttribute(
                'table:message-type',
                $errorStyle === DataValidation::STYLE_WARNING || $errorStyle === DataValidation::STYLE_INFORMATION ? $errorStyle : DataValidation::STYLE_STOP
            );
            $this->writeMessage($dataValidation->getErrorTitle(), $dataValidation->getError(), $dataValidation->getShowErrorMessage());
        }

        $this->objWriter->endElement();
    }

    /**
     * Ends the message element started by the caller.
     */
    private function writeMessage(string $title, string $text, bool $display): void
    {
        if ($title !== '') {
            $this->objWriter->writeAttribute('table:title', $title);
        }
        $this->objWriter->writeAttribute('table:display', $display ? 'true' : 'false');
        if ($text !== '') {
            foreach (explode("\n", str_replace(["\r\n", "\r"], "\n", $text)) as $line) {
                $this->objWriter->writeElement('text:p', $line);
            }
        }
        $this->objWriter->endElement();
    }

    /**
     * The rule as LibreOffice writes it (ScMyValidationsContainer::GetCondition), e.g.
     * of:cell-content-is-whole-number() and cell-content-is-between(1,10).
     */
    private function condition(DataValidation $dataValidation): string
    {
        $type = $dataValidation->getType();
        $formula1 = $dataValidation->getFormula1();
        if ($type === DataValidation::TYPE_LIST) {
            return 'of:cell-content-is-in-list(' . $this->listFormula($formula1) . ')';
        }
        if ($type === DataValidation::TYPE_CUSTOM) {
            return 'of:is-true-formula(' . $this->formula($formula1) . ')';
        }
        $operator = $dataValidation->getOperator();
        $between = $operator === DataValidation::OPERATOR_BETWEEN || $operator === DataValidation::OPERATOR_NOTBETWEEN;
        $hasComparison = $formula1 !== '' || ($between && $dataValidation->getFormula2() !== '');
        if ($type === DataValidation::TYPE_TEXTLENGTH) {
            return $hasComparison ? 'of:' . $this->comparison('cell-content-text-length', $dataValidation) : '';
        }
        if (!isset(self::VALUE_TYPES[$type])) {
            return '';
        }
        // without a comparison, the type alone is written, as LibreOffice does
        $condition = 'of:' . self::VALUE_TYPES[$type];

        return $hasComparison ? $condition . ' and ' . $this->comparison('cell-content', $dataValidation) : $condition;
    }

    private function comparison(string $function, DataValidation $dataValidation): string
    {
        $formula1 = $this->formula($dataValidation->getFormula1());
        $operator = $dataValidation->getOperator();
        if ($operator === DataValidation::OPERATOR_BETWEEN || $operator === DataValidation::OPERATOR_NOTBETWEEN) {
            $function .= $operator === DataValidation::OPERATOR_BETWEEN ? '-is-between(' : '-is-not-between(';

            return $function . $formula1 . ',' . $this->formula($dataValidation->getFormula2()) . ')';
        }

        return $function . '()' . (self::OPERATORS[$operator] ?? '=') . $formula1;
    }

    /**
     * Xlsx keeps a list of values as one string, "a,b,c"; ODF as strings, "a";"b";"c".
     */
    private function listFormula(string $formula): string
    {
        if (strlen($formula) >= 2 && $formula[0] === '"' && str_ends_with($formula, '"')) {
            return '"' . str_replace(',', '";"', substr($formula, 1, -1)) . '"';
        }

        return $this->formula($formula);
    }

    private function formula(string $formula): string
    {
        // convertFormula() returns of:=...; a condition takes the bare expression
        return substr($this->formulaConvertor->convertFormula($formula), 4);
    }
}
