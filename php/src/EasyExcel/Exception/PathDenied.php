<?php

declare(strict_types=1);

namespace EasyExcel\Exception;

/** Raised when a path falls outside EASY_EXCEL_ALLOWED_PATHS (PLAN.md §8). */
class PathDenied extends EasyExcelException
{
}
