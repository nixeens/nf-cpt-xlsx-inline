<?php

namespace PhpOffice\PhpSpreadsheet;

use Throwable;

/**
 * Base exception class for PhpSpreadsheet.
 */
class Exception extends \Exception
{
    public function __construct(string $message = '', int $code = 0, ?Throwable $previous = null)
    {
        parent::__construct($message, $code, $previous);
    }
}
