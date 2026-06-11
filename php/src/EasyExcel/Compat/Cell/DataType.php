<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

/** PhpSpreadsheet's cell data-type constants, unchanged. */
class DataType
{
    public const TYPE_STRING2 = 'str';
    public const TYPE_STRING = 's';
    public const TYPE_FORMULA = 'f';
    public const TYPE_NUMERIC = 'n';
    public const TYPE_BOOL = 'b';
    public const TYPE_NULL = 'null';
    public const TYPE_INLINE = 'inlineStr';
    public const TYPE_ERROR = 'e';
    public const TYPE_ISO_DATE = 'd';
}
