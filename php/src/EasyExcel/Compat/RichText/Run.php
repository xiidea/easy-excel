<?php

declare(strict_types=1);

namespace EasyExcel\Compat\RichText;

use EasyExcel\Compat\Exception;

/** One rich text run; formatting is not supported in Phase 2 (COMPAT.md). */
class Run
{
    public function __construct(private string $text)
    {
    }

    public function getText(): string
    {
        return $this->text;
    }

    public function getFont(): never
    {
        throw new Exception(
            'easy-excel: rich text run formatting is not supported yet (see COMPAT.md)'
        );
    }
}
