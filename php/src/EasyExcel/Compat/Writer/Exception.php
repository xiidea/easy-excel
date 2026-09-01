<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Exception as CompatException;

/**
 * Thrown by writers, PhpSpreadsheet style. Extends the flat Compat exception
 * so existing `catch (PhpOffice\PhpSpreadsheet\Exception $e)` blocks keep
 * working while `catch (Writer\Exception $e)` narrows correctly (wave 5.2).
 */
class Exception extends CompatException
{
}
