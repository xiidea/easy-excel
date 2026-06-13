<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/** Chart legend (PhpSpreadsheet position constants map to excelize). */
class Legend
{
    public const POSITION_RIGHT = 'r';
    public const POSITION_LEFT = 'l';
    public const POSITION_TOP = 't';
    public const POSITION_BOTTOM = 'b';
    public const POSITION_TOPRIGHT = 'tr';

    private const NATIVE = [
        self::POSITION_RIGHT => 'right',
        self::POSITION_LEFT => 'left',
        self::POSITION_TOP => 'top',
        self::POSITION_BOTTOM => 'bottom',
        self::POSITION_TOPRIGHT => 'right',
    ];

    public function __construct(
        private string $position = self::POSITION_RIGHT,
        private mixed $layout = null,
        private bool $overlay = false,
    ) {
    }

    public function getPosition(): string
    {
        return $this->position;
    }

    /** @internal excelize legend position */
    public function nativePosition(): string
    {
        return self::NATIVE[$this->position] ?? 'right';
    }
}
