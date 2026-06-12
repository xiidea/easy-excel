<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/**
 * ARGB color holder. When attached to a style component the setters notify
 * the owner, which forwards the partial spec to the extension.
 */
class Color
{
    public const COLOR_BLACK = 'FF000000';
    public const COLOR_WHITE = 'FFFFFFFF';
    public const COLOR_RED = 'FFFF0000';
    public const COLOR_DARKRED = 'FF800000';
    public const COLOR_BLUE = 'FF0000FF';
    public const COLOR_DARKBLUE = 'FF000080';
    public const COLOR_GREEN = 'FF00FF00';
    public const COLOR_DARKGREEN = 'FF008000';
    public const COLOR_YELLOW = 'FFFFFF00';
    public const COLOR_DARKYELLOW = 'FF808000';

    private string $argb;

    private ?\Closure $onChange = null;

    public function __construct(string $colorValue = self::COLOR_BLACK)
    {
        $this->argb = self::normalize($colorValue);
    }

    /** @internal */
    public function bind(\Closure $onChange): static
    {
        $this->onChange = $onChange;

        return $this;
    }

    public function setARGB(?string $colorValue = self::COLOR_BLACK): static
    {
        $this->argb = self::normalize($colorValue ?? self::COLOR_BLACK);
        ($this->onChange)?->__invoke($this->argb);

        return $this;
    }

    public function setRGB(?string $colorValue = '000000'): static
    {
        return $this->setARGB('FF' . ($colorValue ?? '000000'));
    }

    public function getARGB(): string
    {
        return $this->argb;
    }

    public function getRGB(): string
    {
        return \substr($this->argb, 2);
    }

    public function applyFromArray(array $colorArray): static
    {
        if (isset($colorArray['argb'])) {
            return $this->setARGB((string) $colorArray['argb']);
        }
        if (isset($colorArray['rgb'])) {
            return $this->setRGB((string) $colorArray['rgb']);
        }

        return $this;
    }

    private static function normalize(string $color): string
    {
        $color = \ltrim($color, '#');

        return \strlen($color) === 6 ? 'FF' . $color : $color;
    }
}
