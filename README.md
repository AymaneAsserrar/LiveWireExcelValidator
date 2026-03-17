# Laravel Excel Validator

Validate Excel file uploads row-by-row against configurable column rules.
If errors are found, returns an annotated `.xlsx` file with **red cell highlights**, **hover comments**, and a **"Validation Errors" summary sheet**.
If clean, returns the parsed rows ready to insert into your database.

---

## Installation

```bash
composer require aymane-asserrar/livewire-excel-validator
```

Requires PHP 8.1+ and Laravel 10/11/12 (auto-discovered via service provider).

---

## Basic Usage

```php
use AymaneAsserrar\ExcelValidator\ExcelValidator;

$result = ExcelValidator::validate(
    file:      $request->file('import'),  // UploadedFile or a file path string
    rules:     [
        'name'   => ['required', 'string', 'max_length:100'],
        'email'  => ['required', 'email', 'unique'],
        'age'    => ['numeric', 'min:18', 'max:120'],
        'role'   => ['required', 'in:admin,editor,viewer'],
        'status' => ['required', 'in:active,inactive'],
    ],
    headerRow: 1,  // optional, defaults to 1
);

if ($result->hasErrors()) {
    // Save the annotated file path in the session and redirect
    session(['annotated_path' => $result->annotatedPath]);
    return back()->with('error', $result->errorCount() . ' error(s) found.');
}

// No errors — use $result->rows to insert into the database
foreach ($result->rows as $row) {
    User::create([
        'name'  => $row['name'],
        'email' => $row['email'],
    ]);
}
```

---

## Custom Column Labels

Add `_label:My Label` to any rule list to override how the column name appears in error messages:

```php
'email' => ['required', 'email', 'unique', '_label:Email Address'],
```

Without `_label`, the column key is title-cased automatically (`email_address` → `Email address`).

---

## Available Rules

| Rule | Description |
|------|-------------|
| `required` | Cell must not be empty |
| `string` | Must be non-numeric text |
| `numeric` | Must be a numeric value |
| `integer` | Must be a whole number |
| `email` | Must be a valid email address |
| `url` | Must be a valid URL |
| `date` | Must be a parseable date (or an Excel serial date) |
| `boolean` | Must be `true`/`false`, `yes`/`no`, or `1`/`0` |
| `min:{n}` | Numeric value must be ≥ n |
| `max:{n}` | Numeric value must be ≤ n |
| `min_length:{n}` | String length must be ≥ n characters |
| `max_length:{n}` | String length must be ≤ n characters |
| `in:{a,b,c}` | Value must be one of the listed options |
| `not_in:{a,b,c}` | Value must not be any of the listed options |
| `regex:{pattern}` | Value must match the regex (no delimiters — e.g. `regex:^\d{4}$`) |
| `unique` | Value must be unique within that column across all rows |

Rules are evaluated in order and stop at the **first failure** per cell.

---

## The `ValidationResult` Object

| Property / Method | Description |
|-------------------|-------------|
| `$result->hasErrors()` | `true` if any row failed validation |
| `$result->errorCount()` | Total number of individual cell errors |
| `$result->errors` | Raw errors array: `[rowIndex => [colKey => [messages]]]` |
| `$result->flatErrors()` | Flat array: `[['row' => n, 'column' => 'email', 'messages' => [...]]]` |
| `$result->rows` | Parsed rows as `array<int, array<string, string>>`, keyed by lowercased header name |
| `$result->annotatedPath` | Absolute path to the annotated `.xlsx` file (null if no errors) |
| `$result->filename` | Original filename |
| `$result->download()` | Stream the annotated file as a download response **(Laravel only)** |

---

## Serving the Annotated Error File

The annotated file is written to `sys_get_temp_dir()`. The recommended pattern is to store the path in the session and serve it via a dedicated route:

```php
// In your controller / Livewire component
session(['excel_errors_path' => $result->annotatedPath]);

// In routes/web.php
Route::get('/excel/download-errors', function () {
    $path = session('excel_errors_path');
    abort_unless($path && file_exists($path), 404);
    session()->forget('excel_errors_path');
    return response()->download($path, 'validation_errors.xlsx')->deleteFileAfterSend(true);
})->name('excel.download-errors');
```

Or use the built-in helper directly in a controller:

```php
return $result->download('my_errors.xlsx');
```

---

## What the Annotated File Looks Like

- **Red background + red text** on every invalid cell
- **Hover comment** on each red cell listing the exact rule(s) that failed
- **Amber header** on every column that contains at least one error
- **"Validation Errors" sheet** appended at the end with columns: Row / Column / Value / Error

---

## Non-Laravel Usage

The validator has no hard Laravel dependency. Pass a file path string instead of an `UploadedFile`:

```php
$result = ExcelValidator::validate('/tmp/upload.xlsx', $rules);
```

`ValidationResult::download()` will throw a `RuntimeException` if called outside a Laravel context — use `$result->annotatedPath` directly and serve the file yourself.

---

## License

MIT — [Aymane Asserrar](https://github.com/AymaneAsserrar)
