<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader;

use EasyExcel\Compat\Exception as CompatException;

/**
 * Thrown by readers, PhpSpreadsheet style. Extends the flat Compat exception
 * so broad catches keep working while `catch (Reader\Exception $e)` narrows
 * correctly (wave 5.2).
 */
class Exception extends CompatException
{
}
