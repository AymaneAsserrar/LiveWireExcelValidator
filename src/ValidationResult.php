<?php

namespace AymaneAsserrar\ExcelValidator;

class ValidationResult
{
    public function __construct(
        public readonly array   $rows,
        public readonly array   $errors,
        public readonly ?string $annotatedPath,
        public readonly string  $filename,
    ) {}

    public function hasErrors(): bool
    {
        return ! empty($this->errors);
    }

    public function errorCount(): int
    {
        return array_sum(array_map('count', $this->errors));
    }

    /**
     * Returns a flat list of errors: [['row' => n, 'column' => 'key', 'messages' => [...]], ...]
     */
    public function flatErrors(): array
    {
        $flat = [];
        foreach ($this->errors as $row => $cols) {
            foreach ($cols as $col => $messages) {
                $flat[] = ['row' => $row, 'column' => $col, 'messages' => $messages];
            }
        }

        return $flat;
    }

    /**
     * Stream the annotated file as a download response (requires Laravel).
     *
     * @throws \LogicException  if validation passed (no annotated file exists).
     * @throws \RuntimeException if called outside a Laravel/Symfony HTTP context.
     */
    public function download(?string $filename = null): mixed
    {
        if ($this->annotatedPath === null) {
            throw new \LogicException('No annotated file — validation passed without errors.');
        }

        if (! function_exists('response')) {
            throw new \RuntimeException(
                'ValidationResult::download() requires a Laravel HTTP context. ' .
                'Use $result->annotatedPath directly and serve the file yourself.'
            );
        }

        $stem         = pathinfo($this->filename, PATHINFO_FILENAME);
        $downloadName = $filename ?? ($stem . '_validation_errors.xlsx');

        return response()->download(
            $this->annotatedPath,
            $downloadName,
            ['Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
        )->deleteFileAfterSend(true);
    }
}
