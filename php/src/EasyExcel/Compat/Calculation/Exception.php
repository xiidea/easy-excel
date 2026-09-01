<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Calculation;

use EasyExcel\Compat\Exception as CompatException;

/**
 * Thrown by the calculation engine, PhpSpreadsheet style. Extends the flat
 * Compat exception so broad catches keep working while
 * `catch (Calculation\Exception $e)` narrows correctly (wave 5.2).
 */
class Exception extends CompatException
{
}
