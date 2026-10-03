<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use Closure;
use Composer\Pcre\Preg;
use DOMElement;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\ReferenceHelper;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Reads table:content-validations, and applies them to the cells that name them in table:content-validation-name.
 */
class DataValidations
{
    private const OPERATORS = [
        '=' => DataValidation::OPERATOR_EQUAL,
        '>' => DataValidation::OPERATOR_GREATERTHAN,
        '>=' => DataValidation::OPERATOR_GREATERTHANOREQUAL,
        '<' => DataValidation::OPERATOR_LESSTHAN,
        '<=' => DataValidation::OPERATOR_LESSTHANOREQUAL,
        '!=' => DataValidation::OPERATOR_NOTEQUAL,
    ];

    private const VALUE_TYPES = [
        'whole-number' => DataValidation::TYPE_WHOLE,
        'decimal-number' => DataValidation::TYPE_DECIMAL,
        'date' => DataValidation::TYPE_DATE,
        'time' => DataValidation::TYPE_TIME,
    ];

    /**
     * Validations by name, with the base cell of their relative references.
     *
     * @var array<string, array{DataValidation, ?array{int, int}}>
     */
    private array $validations = [];

    /**
     * Cell areas of the current sheet by validation name: first column, first row, last column, last row.
     *
     * @var array<string, list<array{int, int, int, int}>>
     */
    private array $areas = [];

    /**
     * @param Closure(DOMElement): string $textOf the text of a text:p
     */
    public function __construct(
        private string $tableNs,
        private string $textNs,
        private Closure $textOf
    ) {
    }

    public function read(DOMElement $workbookData): void
    {
        foreach ($workbookData->getElementsByTagNameNS($this->tableNs, 'content-validation') as $element) {
            $dataValidation = new DataValidation();
            $dataValidation->setAllowBlank($element->getAttributeNS($this->tableNs, 'allow-empty-cell') !== 'false');
            // none in the ODF schema, no as LibreOffice writes it
            $dataValidation->setShowDropDown(!in_array($element->getAttributeNS($this->tableNs, 'display-list'), ['no', 'none'], true));
            $this->readCondition($dataValidation, $element->getAttributeNS($this->tableNs, 'condition'));

            foreach ($element->getElementsByTagNameNS($this->tableNs, 'help-message') as $message) {
                $dataValidation->setShowInputMessage($message->getAttributeNS($this->tableNs, 'display') === 'true');
                $dataValidation->setPromptTitle($message->getAttributeNS($this->tableNs, 'title'));
                $dataValidation->setPrompt($this->messageText($message));
            }
            foreach ($element->getElementsByTagNameNS($this->tableNs, 'error-message') as $message) {
                $dataValidation->setShowErrorMessage($message->getAttributeNS($this->tableNs, 'display') === 'true');
                $dataValidation->setErrorTitle($message->getAttributeNS($this->tableNs, 'title'));
                $dataValidation->setError($this->messageText($message));
                // the ODF message types are named as the error styles; macro and reject have none
                $messageType = $message->getAttributeNS($this->tableNs, 'message-type');
                if ($messageType === DataValidation::STYLE_WARNING || $messageType === DataValidation::STYLE_INFORMATION) {
                    $dataValidation->setErrorStyle($messageType);
                }
            }

            $baseCell = null;
            if (Preg::isMatch('/\$?([A-Z]{1,3})\$?(\d+)$/i', $element->getAttributeNS($this->tableNs, 'base-cell-address'), $matches)) {
                $baseCell = [Coordinate::columnIndexFromString($matches[1]), (int) $matches[2]];
            }
            $this->validations[$element->getAttributeNS($this->tableNs, 'name')] = [$dataValidation, $baseCell];
        }
    }

    /**
     * Note cells, from a table:table-cell with its repeats, that name a validation.
     */
    public function addCells(string $name, int $column, int $row, int $columns, int $rows): void
    {
        if (isset($this->validations[$name])) {
            $this->areas[$name][] = [$column, $row, $column + $columns - 1, $row + $rows - 1];
        }
    }

    /**
     * Set the validations noted for a sheet on that sheet.
     */
    public function applyTo(Worksheet $worksheet): void
    {
        foreach ($this->areas as $name => $areas) {
            $areas = $this->joinAreas($areas);
            [$prototype, $baseCell] = $this->validations[$name];
            $dataValidation = clone $prototype;
            // ODF formulas are relative to the base cell, Xlsx ones to the top left cell
            if ($baseCell !== null && [$areas[0][0], $areas[0][1]] !== $baseCell) {
                $dataValidation->setFormula1($this->moveFormula($dataValidation->getFormula1(), $areas[0][0] - $baseCell[0], $areas[0][1] - $baseCell[1]));
                $dataValidation->setFormula2($this->moveFormula($dataValidation->getFormula2(), $areas[0][0] - $baseCell[0], $areas[0][1] - $baseCell[1]));
            }
            $ranges = [];
            foreach ($areas as [$colStart, $rowStart, $colEnd, $rowEnd]) {
                $range = Coordinate::stringFromColumnIndex($colStart) . $rowStart;
                if ($colEnd !== $colStart || $rowEnd !== $rowStart) {
                    $range .= ':' . Coordinate::stringFromColumnIndex($colEnd) . $rowEnd;
                }
                $ranges[] = $range;
            }
            $worksheet->setDataValidation(implode(' ', $ranges), $dataValidation);
        }
        $this->areas = [];
    }

    /**
     * The cells come a row at a time, left to right: join cells next to each other in a row,
     * then runs of the same columns in rows below each other.
     *
     * @param list<array{int, int, int, int}> $areas
     *
     * @return non-empty-list<array{int, int, int, int}>
     */
    private function joinAreas(array $areas): array
    {
        $rowRuns = [];
        foreach ($areas as $area) {
            $last = count($rowRuns) - 1;
            if ($last >= 0 && $rowRuns[$last][1] === $area[1] && $rowRuns[$last][3] === $area[3] && $rowRuns[$last][2] + 1 === $area[0]) {
                $rowRuns[$last][2] = $area[2];
            } else {
                $rowRuns[] = $area;
            }
        }
        $joined = [];
        $lastByColumns = [];
        foreach ($rowRuns as $area) {
            $key = $area[0] . ':' . $area[2];
            $index = $lastByColumns[$key] ?? null;
            if ($index !== null && $joined[$index][3] + 1 === $area[1]) {
                $joined[$index][3] = $area[3];
            } else {
                $lastByColumns[$key] = count($joined);
                $joined[] = $area;
            }
        }

        /** @var non-empty-list<array{int, int, int, int}> $joined addCells() is called before applyTo() for each name */
        return $joined;
    }

    /**
     * Move the relative references outside string literals.
     */
    private function moveFormula(string $formula, int $columns, int $rows): string
    {
        $parts = Preg::split('/(' . Calculation::CALCULATION_REGEXP_STRING . ')/', $formula, -1, PREG_SPLIT_DELIM_CAPTURE);
        foreach ($parts as $index => $part) {
            if ($index % 2 === 0) {
                $parts[$index] = ReferenceHelper::getInstance()->updateFormulaReferencesAnyWorksheet($part, $columns, $rows);
            }
        }

        return implode('', $parts);
    }

    private function messageText(DOMElement $message): string
    {
        $lines = [];
        foreach ($message->getElementsByTagNameNS($this->textNs, 'p') as $paragraph) {
            $lines[] = ($this->textOf)($paragraph);
        }

        return implode("\n", $lines);
    }

    /**
     * Parse the rule, as ODF 1.2 to 1.4 write it: a namespace prefix (of:, oooc:, ooow:) followed by e.g.
     * cell-content-is-whole-number() and cell-content-is-between(1,10).
     */
    private function readCondition(DataValidation $dataValidation, string $condition): void
    {
        $condition = trim(Preg::replace('/^\s*[\w.-]+:/', '', $condition));
        if (Preg::isMatch('/^cell-content-is-in-list\s*\((.*)\)$/s', $condition, $matches)) {
            $dataValidation->setType(DataValidation::TYPE_LIST);
            $list = trim($matches[1]);
            // "a";"b";"c" (or numbers) is a list of values, kept in Xlsx as one string "a,b,c"
            $item = '"(?:[^"]|"")*"|[-+]?(?:\d+\.?\d*|\.\d+)(?:E[-+]?\d+)?';
            if (Preg::isMatch('/^(?:' . $item . ')(?:\s*;\s*(?:' . $item . '))*$/i', $list)) {
                Preg::matchAll('/' . $item . '/i', $list, $items);
                $values = array_map(fn (string $value): string => $value[0] === '"' ? substr($value, 1, -1) : $value, $items[0]);
                $dataValidation->setFormula1('"' . implode(',', $values) . '"');
            } else {
                $dataValidation->setFormula1($this->formula($list));
            }

            return;
        }
        if (Preg::isMatch('/^is-true-formula\s*\((.*)\)$/s', $condition, $matches)) {
            $dataValidation->setType(DataValidation::TYPE_CUSTOM);
            $dataValidation->setFormula1($this->formula($matches[1]));

            return;
        }

        if (Preg::isMatch('/^cell-content-is-(whole-number|decimal-number|date|time)\s*\(\s*\)\s*and\s+(.*)$/s', $condition, $matches)) {
            $type = self::VALUE_TYPES[$matches[1]];
            $function = 'cell-content';
            $comparison = $matches[2];
        } elseif (str_starts_with($condition, 'cell-content-text-length')) {
            $type = DataValidation::TYPE_TEXTLENGTH;
            $function = 'cell-content-text-length';
            $comparison = $condition;
        } else {
            // no rule, or one this reader does not know
            return;
        }
        if (Preg::isMatch('/^' . $function . '\s*\(\s*\)\s*(<=|>=|!=|<|>|=)(.*)$/s', $comparison, $matches)) {
            $dataValidation->setOperator(self::OPERATORS[$matches[1]]);
            $dataValidation->setFormula1($this->formula($matches[2]));
        } elseif (Preg::isMatch('/^' . $function . '-is-(between|not-between)\s*\((.*)\)$/s', $comparison, $matches)) {
            $dataValidation->setOperator($matches[1] === 'between' ? DataValidation::OPERATOR_BETWEEN : DataValidation::OPERATOR_NOTBETWEEN);
            [$formula1, $formula2] = $this->splitArguments($matches[2]);
            $dataValidation->setFormula1($this->formula($formula1));
            $dataValidation->setFormula2($this->formula($formula2));
        } else {
            return;
        }
        $dataValidation->setType($type);
    }

    /**
     * Split the arguments of between() at the comma outside brackets and string literals.
     *
     * @return array{string, string}
     */
    private function splitArguments(string $arguments): array
    {
        $depth = 0;
        $quote = '';
        $length = strlen($arguments);
        for ($i = 0; $i < $length; ++$i) {
            $char = $arguments[$i];
            if ($quote !== '') {
                $quote = $char === $quote ? '' : $quote;
            } elseif ($char === '"' || $char === "'") {
                $quote = $char;
            } elseif ($char === '(' || $char === '{' || $char === '[') {
                ++$depth;
            } elseif ($char === ')' || $char === '}' || $char === ']') {
                --$depth;
            } elseif ($char === ',' && $depth === 0) {
                return [substr($arguments, 0, $i), substr($arguments, $i + 1)];
            }
        }

        return [$arguments, ''];
    }

    private function formula(string $formula): string
    {
        return FormulaTranslator::convertToExcelFormulaValue(trim($formula));
    }
}
