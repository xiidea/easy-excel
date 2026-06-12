<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** Cell protection component (PhpSpreadsheet uses string flags). */
class Protection
{
    public const PROTECTION_INHERIT = 'inherit';
    public const PROTECTION_PROTECTED = 'protected';
    public const PROTECTION_UNPROTECTED = 'unprotected';

    private string $locked = self::PROTECTION_INHERIT;
    private string $hidden = self::PROTECTION_INHERIT;

    public function __construct(private Style $style)
    {
    }

    public function setLocked(string $lockType): static
    {
        $this->locked = $lockType;
        $this->style->mergeComponent('protection', ['locked' => $lockType]);

        return $this;
    }

    public function setHidden(string $hiddenType): static
    {
        $this->hidden = $hiddenType;
        $this->style->mergeComponent('protection', ['hidden' => $hiddenType]);

        return $this;
    }

    public function getLocked(): string
    {
        return $this->locked;
    }

    public function getHidden(): string
    {
        return $this->hidden;
    }
}
