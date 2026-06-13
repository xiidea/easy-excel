<?php

declare(strict_types=1);

namespace EasyExcel\Compat\RichText;

/**
 * Rich text container. Used both as a cell value (multi-format runs, wave 4.4)
 * and for plain-text comments (the onChange hook). Run formatting is sent to
 * the extension when the rich text is assigned to a cell.
 */
class RichText
{
    /** @var list<Run> */
    private array $runs = [];

    private ?\Closure $onChange = null;

    public function __construct(?string $text = null)
    {
        if ($text !== null && $text !== '') {
            $this->createText($text);
        }
    }

    /** @internal comment hook: notified with the plain text on each append */
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

    public function addText(Run $run): static
    {
        $this->runs[] = $run;
        ($this->onChange)?->__invoke($this->getPlainText());

        return $this;
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

    /** @internal @return list<array{text: string, font?: array<string, mixed>}> */
    public function toRunSpecs(): array
    {
        return \array_map(static fn (Run $r): array => $r->toRunSpec(), $this->runs);
    }

    private function append(string $text): Run
    {
        $run = new Run($text);
        $this->runs[] = $run;
        ($this->onChange)?->__invoke($this->getPlainText());

        return $run;
    }
}
