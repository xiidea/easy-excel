<?php

declare(strict_types=1);

namespace EasyExcel\Compat\RichText;

use EasyExcel\Compat\Style\Font;
use EasyExcel\Compat\Style\Style;

/**
 * One rich text run: text plus an optional font (wave 4.4). The font is
 * collected on a detached Style so its spec matches cell styling.
 */
class Run
{
    private ?Style $fontStyle = null;

    public function __construct(private string $text)
    {
    }

    public function getText(): string
    {
        return $this->text;
    }

    public function setText(string $text): static
    {
        $this->text = $text;

        return $this;
    }

    public function getFont(): Font
    {
        $this->fontStyle ??= Style::detached();

        return $this->fontStyle->getFont();
    }

    /** @internal @return array{text: string, font?: array<string, mixed>} */
    public function toRunSpec(): array
    {
        $spec = ['text' => $this->text];
        $font = $this->fontStyle?->getCollectedSpec()['font'] ?? null;
        if (\is_array($font) && $font !== []) {
            $spec['font'] = $font;
        }

        return $spec;
    }
}
