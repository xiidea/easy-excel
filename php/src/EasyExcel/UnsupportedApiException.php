<?php

declare(strict_types=1);

namespace EasyExcel;

/**
 * Thrown in all-or-nothing ("strict") aliasing mode when application code
 * references a PhpOffice\PhpSpreadsheet\* class the Compat layer does not
 * implement.
 *
 * Failing loudly here is deliberate: it prevents a handle-based Compat
 * workbook from being silently mixed with a real PhpSpreadsheet object graph
 * in the same request (which would corrupt output). See COMPAT.md "aliasing
 * modes".
 */
class UnsupportedApiException extends \LogicException
{
}
