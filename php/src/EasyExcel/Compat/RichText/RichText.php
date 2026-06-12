<?php

declare(strict_types=1);

namespace EasyExcel\Compat\RichText;

/**
 * Minimal rich text container (Phase-2 subset: plain text only, used for
 * comments). Per-run formatting raises a clear unsupported error instead of
 * silently dropping it.
 */
class RichText
{
    /** @var list<Run> */
    private array $runs = [];

    private ?\Closure $onChange = null;

    /** @internal */
    public function bind(\Closure $onChange): static
    {
        $this->onChange = $onChange;

        return $this;
    }

    public function createText(string $text): Run
    {
        return $this->append($text);
    }

    public function createTextRun(string $text): Run
    {
        return $this->append($text);
    }

    public function getPlainText(): string
    {
        return \implode('', \array_map(static fn (Run $r): string => $r->getText(), $this->runs));
    }

    public function __toString(): string
    {
        return $this->getPlainText();
    }

    /** @return list<Run> */
    public function getRichTextElements(): array
    {
        return $this->runs;
    }

    private function append(string $text): Run
    {
        $run = new Run($text);
        $this->runs[] = $run;
        ($this->onChange)?->__invoke($this->getPlainText());

        return $run;
    }
}
