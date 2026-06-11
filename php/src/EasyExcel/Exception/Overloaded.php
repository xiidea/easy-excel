<?php

declare(strict_types=1);

namespace EasyExcel\Exception;

/**
 * Raised when the extension's admission control or memory budget rejects a
 * heavy operation (PLAN.md §7). Convert to HTTP 429 or dispatch the export
 * to your queue.
 */
class Overloaded extends EasyExcelException
{
}
