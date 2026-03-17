<?php

namespace AymaneAsserrar\ExcelValidator;

use Illuminate\Support\ServiceProvider;

class ExcelValidatorServiceProvider extends ServiceProvider
{
    public function register(): void
    {
        // No binding needed — ExcelValidator::validate() is a static entry point.
        // This provider exists for auto-discovery and future extensibility.
    }

    public function boot(): void
    {
        //
    }
}
