<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

use EasyExcel\Compat\RichText\RichText;

/** Chart or axis title (caption may be a string or RichText). */
class Title
{
    public function __construct(private string|RichText|null $caption = null)
    {
    }

    public function getCaption(): string|RichText|null
    {
        return $this->caption;
    }

    public function getCaptionText(): string
    {
        if ($this->caption instanceof RichText) {
            return $this->caption->getPlainText();
        }

        return (string) $this->caption;
    }

    public function setCaption(string|RichText|null $caption): static
    {
        $this->caption = $caption;

        return $this;
    }
}
