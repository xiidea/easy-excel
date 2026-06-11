<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Phase-1 style facade: number formats only (PLAN.md §5).
 * Font/Fill/Borders/Alignment land in Phase 2; calling anything beyond
 * getNumberFormat() throws a clear "not yet supported" error instead of
 * silently producing wrong files.
 */
class Style
{
    public function __construct(private Worksheet $worksheet, private string $range)
    {
    }

    public function getNumberFormat(): NumberFormat
    {
        return new NumberFormat($this->worksheet, $this->range);
    }

    /** @param array<string, mixed> $styleArray */
    public function applyFromArray(array $styleArray): static
    {
        foreach ($styleArray as $key => $value) {
            if ($key === 'numberFormat' && isset($value['formatCode'])) {
                $this->getNumberFormat()->setFormatCode((string) $value['formatCode']);
                continue;
            }

            throw new \EasyExcel\Compat\Exception(
                "easy-excel: style component \"$key\" is not supported yet (Phase 2, see COMPAT.md)"
            );
        }

        return $this;
    }

    public function __call(string $name, array $arguments): mixed
    {
        throw new \EasyExcel\Compat\Exception(
            "easy-excel: Style::$name() is not supported yet (Phase 2, see COMPAT.md)"
        );
    }
}
