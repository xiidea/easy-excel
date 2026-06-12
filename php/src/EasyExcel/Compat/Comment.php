<?php

declare(strict_types=1);

namespace EasyExcel\Compat;

use EasyExcel\Compat\RichText\RichText;
use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/** Plain-text cell comment; mutations re-push the full comment (replace). */
class Comment
{
    private string $author = 'Author';

    private RichText $text;

    public function __construct(private Worksheet $worksheet, private string $cell)
    {
        $this->text = (new RichText())->bind(fn () => $this->push());
    }

    public function getText(): RichText
    {
        return $this->text;
    }

    public function setAuthor(string $author): static
    {
        $this->author = $author;
        if ($this->text->getPlainText() !== '') {
            $this->push();
        }

        return $this;
    }

    public function getAuthor(): string
    {
        return $this->author;
    }

    private function push(): void
    {
        Native::setComment(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->cell,
            $this->author,
            $this->text->getPlainText(),
        );
    }
}
