<?php

namespace AymaneAsserrar\ExcelValidator;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelValidator
{
    private const ERROR_BG     = 'FFCCCC';
    private const ERROR_FONT   = 'CC0000';
    private const HEADER_ERROR = 'FFE0B2';

    private array $errors        = [];
    private array $uniqueTracker = [];
    private array $rows          = [];

    /**
     * Validate an Excel file against a set of column rules.
     *
     * @param  string|\Illuminate\Http\UploadedFile  $file   Path to the file, or a Laravel UploadedFile instance.
     * @param  array<string, list<string>>            $rules  Column rules keyed by header name (case-insensitive).
     * @param  int                                    $headerRow  1-based row index of the header row.
     */
    public static function validate(
        mixed $file,
        array $rules,
        int $headerRow = 1,
    ): ValidationResult {
        return (new self())->run($file, $rules, $headerRow);
    }

    private function coord(int $col, int $row): string
    {
        return Coordinate::stringFromColumnIndex($col) . $row;
    }

    private function run(mixed $file, array $rules, int $headerRow): ValidationResult
    {
        // Accept both a plain path string and a Laravel/Symfony UploadedFile
        if (is_object($file) && method_exists($file, 'getRealPath')) {
            $path     = $file->getRealPath();
            $filename = method_exists($file, 'getClientOriginalName')
                ? $file->getClientOriginalName()
                : basename($path);
        } else {
            $path     = (string) $file;
            $filename = basename($path);
        }

        $spreadsheet = IOFactory::load($path);
        $sheet       = $spreadsheet->getActiveSheet();

        $parsedRules = [];
        foreach ($rules as $colKey => $colRules) {
            $label      = ucfirst(str_replace(['_', '-'], ' ', $colKey));
            $cleanRules = [];
            foreach ($colRules as $rule) {
                if (str_starts_with($rule, '_label:')) {
                    $label = trim(substr($rule, 7));
                } else {
                    $cleanRules[] = $rule;
                }
            }
            $parsedRules[strtolower(trim($colKey))] = ['rules' => $cleanRules, 'label' => $label];
        }

        $headers = $this->readHeaders($sheet, $headerRow);
        $lastRow = $sheet->getHighestDataRow();

        for ($row = $headerRow + 1; $row <= $lastRow; $row++) {
            $rowData = [];
            foreach ($headers as $normKey => $colIdx) {
                $rowData[$normKey] = $sheet->getCell($this->coord($colIdx, $row))->getFormattedValue();
            }

            if (empty(array_filter($rowData, fn($v) => trim((string) $v) !== ''))) {
                continue; // skip blank rows
            }

            $this->rows[$row] = $rowData;

            foreach ($parsedRules as $normKey => $def) {
                $colIdx = $headers[$normKey] ?? null;
                $value  = trim((string) ($rowData[$normKey] ?? ''));
                $raw    = $colIdx ? $sheet->getCell($this->coord($colIdx, $row))->getValue() : '';

                foreach ($def['rules'] as $rule) {
                    $error = $this->applyRule($rule, $value, $raw, $def['label'], $normKey);
                    if ($error !== null) {
                        $this->errors[$row][$normKey][] = $error;
                        break;
                    }
                }
            }
        }

        $annotatedPath = null;
        if (! empty($this->errors)) {
            $annotatedPath = $this->annotate($sheet, $spreadsheet, $headers, $parsedRules, $headerRow, $filename);
        }

        return new ValidationResult(
            rows:          array_values($this->rows),
            errors:        $this->errors,
            annotatedPath: $annotatedPath,
            filename:      $filename,
        );
    }

    private function annotate(mixed $sheet, mixed $spreadsheet, array $headers, array $parsedRules, int $headerRow, string $filename): string
    {
        $columnsWithErrors = [];
        foreach ($this->errors as $colErrors) {
            foreach (array_keys($colErrors) as $normKey) {
                $columnsWithErrors[$normKey] = true;
            }
        }

        // Highlight invalid cells
        foreach ($this->errors as $rowIdx => $colErrors) {
            foreach ($colErrors as $normKey => $messages) {
                $colIdx = $headers[$normKey] ?? null;
                if ($colIdx === null) continue;

                $coord = $this->coord($colIdx, $rowIdx);

                $sheet->getStyle($coord)->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setRGB(self::ERROR_BG);
                $sheet->getStyle($coord)->getFont()
                    ->getColor()->setRGB(self::ERROR_FONT);

                $rt    = new RichText();
                $title = $rt->createTextRun("⚠ Validation error:\n");
                $title->getFont()->setBold(true)->setSize(9);
                foreach ($messages as $msg) {
                    $rt->createTextRun("• {$msg}\n")->getFont()->setSize(9);
                }
                $comment = $sheet->getComment($coord);
                $comment->setText($rt);
                $comment->setVisible(false);
                $comment->setWidth('200pt');
                $comment->setHeight((count($messages) * 18 + 28) . 'pt');
            }
        }

        // Amber header on columns that have errors
        foreach ($headers as $normKey => $colIdx) {
            if (isset($columnsWithErrors[$normKey])) {
                $coord = $this->coord($colIdx, $headerRow);
                $sheet->getStyle($coord)->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setRGB(self::HEADER_ERROR);
                $sheet->getStyle($coord)->getFont()->setBold(true);
            }
        }

        $this->addSummarySheet($spreadsheet, $headers, $parsedRules);

        $stem = pathinfo($filename, PATHINFO_FILENAME);
        $out  = sys_get_temp_dir() . DIRECTORY_SEPARATOR . $stem . '_errors_' . time() . '.xlsx';
        IOFactory::createWriter($spreadsheet, 'Xlsx')->save($out);

        return $out;
    }

    private function addSummarySheet(mixed $spreadsheet, array $headers, array $parsedRules): void
    {
        if ($spreadsheet->sheetNameExists('Validation Errors')) {
            $spreadsheet->removeSheetByIndex(
                $spreadsheet->getIndex($spreadsheet->getSheetByName('Validation Errors'))
            );
        }

        $s = $spreadsheet->createSheet();
        $s->setTitle('Validation Errors');

        $cols = ['Row', 'Column', 'Value', 'Error'];
        foreach ($cols as $i => $h) {
            $c = Coordinate::stringFromColumnIndex($i + 1) . '1';
            $s->getCell($c)->setValue($h);
            $s->getStyle($c)->getFont()->setBold(true);
            $s->getStyle($c)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('2D3748');
            $s->getStyle($c)->getFont()->getColor()->setRGB('FFFFFF');
        }

        $r         = 2;
        $mainSheet = $spreadsheet->getActiveSheet();

        foreach ($this->errors as $rowIdx => $colErrors) {
            foreach ($colErrors as $normKey => $messages) {
                $colIdx = $headers[$normKey] ?? null;
                $value  = $colIdx ? $mainSheet->getCell($this->coord($colIdx, $rowIdx))->getFormattedValue() : '—';
                $label  = $parsedRules[$normKey]['label'] ?? $normKey;

                $s->getCell($this->coord(1, $r))->setValue($rowIdx);
                $s->getCell($this->coord(2, $r))->setValue($label);
                $s->getCell($this->coord(3, $r))->setValue($value);
                $s->getCell($this->coord(4, $r))->setValue(implode(' | ', $messages));

                if ($r % 2 === 0) {
                    $s->getStyle("A{$r}:D{$r}")->getFill()
                        ->setFillType(Fill::FILL_SOLID)
                        ->getStartColor()->setRGB('F7FAFC');
                }
                $r++;
            }
        }

        foreach (['A', 'B', 'C', 'D'] as $col) {
            $s->getColumnDimension($col)->setAutoSize(true);
        }

        $s->getCell($this->coord(1, $r))->setValue('Total:');
        $s->getCell($this->coord(2, $r))->setValue(($r - 2) . ' error(s)');
        $s->getStyle("A{$r}:B{$r}")->getFont()->setBold(true);
    }

    private function applyRule(string $rule, string $value, mixed $raw, string $label, string $normKey): ?string
    {
        $empty = ($value === '');

        if ($rule === 'required') {
            return $empty ? "{$label} is required." : null;
        }

        if ($empty) return null;

        return match (true) {
            $rule === 'numeric'   => is_numeric($value) ? null : "{$label} must be a numeric value.",
            $rule === 'integer'   => (is_numeric($value) && (int) $value == $value) ? null : "{$label} must be an integer.",
            $rule === 'string'    => (! is_numeric($value)) ? null : "{$label} must be text, not a number.",
            $rule === 'email'     => filter_var($value, FILTER_VALIDATE_EMAIL) ? null : "{$label} must be a valid email address.",
            $rule === 'url'       => filter_var($value, FILTER_VALIDATE_URL) ? null : "{$label} must be a valid URL.",
            $rule === 'boolean'   => in_array(strtolower($value), ['0', '1', 'true', 'false', 'yes', 'no'], true) ? null : "{$label} must be true/false, yes/no, or 1/0.",
            $rule === 'date'      => $this->validateDate($value, $raw, $label),
            $rule === 'unique'    => $this->validateUnique($value, $normKey, $label),

            str_starts_with($rule, 'min:')        => (is_numeric($value) && (float) $value >= (float) substr($rule, 4)) ? null : "{$label} must be at least " . substr($rule, 4) . ".",
            str_starts_with($rule, 'max:')        => (is_numeric($value) && (float) $value <= (float) substr($rule, 4)) ? null : "{$label} must be at most " . substr($rule, 4) . ".",
            str_starts_with($rule, 'min_length:') => mb_strlen($value) >= (int) substr($rule, 11) ? null : "{$label} must be at least " . substr($rule, 11) . " characters.",
            str_starts_with($rule, 'max_length:') => mb_strlen($value) <= (int) substr($rule, 11) ? null : "{$label} must not exceed " . substr($rule, 11) . " characters.",
            str_starts_with($rule, 'in:')         => in_array($value, array_map('trim', explode(',', substr($rule, 3))), true) ? null : "{$label} must be one of: " . substr($rule, 3) . ".",
            str_starts_with($rule, 'not_in:')     => ! in_array($value, array_map('trim', explode(',', substr($rule, 7))), true) ? null : "{$label} must not be one of: " . substr($rule, 7) . ".",
            str_starts_with($rule, 'regex:')      => (@preg_match('/' . substr($rule, 6) . '/', $value) ? null : "{$label} format is invalid."),

            default => null,
        };
    }

    private function validateDate(string $value, mixed $raw, string $label): ?string
    {
        $ts = is_numeric($raw)
            ? \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp((float) $raw)
            : strtotime($value);

        return ($ts !== false) ? null : "{$label} must be a valid date.";
    }

    private function validateUnique(string $value, string $normKey, string $label): ?string
    {
        if (isset($this->uniqueTracker[$normKey][$value])) {
            return "{$label} must be unique — \"{$value}\" is duplicated.";
        }
        $this->uniqueTracker[$normKey][$value] = true;

        return null;
    }

    private function readHeaders(Worksheet $sheet, int $headerRow): array
    {
        $headers = [];
        $lastCol = Coordinate::columnIndexFromString($sheet->getHighestDataColumn());
        for ($c = 1; $c <= $lastCol; $c++) {
            $val = trim((string) $sheet->getCell($this->coord($c, $headerRow))->getValue());
            if ($val !== '') {
                $headers[strtolower($val)] = $c;
            }
        }

        return $headers;
    }
}
