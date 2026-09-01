<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

/**
 * Shared state for Drawing and MemoryDrawing, PhpSpreadsheet style (wave 5.2).
 *
 * Extracted from the two existing drawing classes rather than bolted on: this
 * holds only what both genuinely share (name, placement, size, the owning
 * sheet). Attachment stays in the subclasses because each sends a different
 * payload to the extension — a file path vs. rendered image bytes.
 */
abstract class BaseDrawing
{
    protected string $name = '';

    protected string $description = '';

    protected string $coordinates = 'A1';

    protected int $offsetX = 0;

    protected int $offsetY = 0;

    protected int $width = 0;

    protected int $height = 0;

    protected ?Worksheet $worksheet = null;

    public function setName(string $name): static
    {
        $this->name = $name;

        return $this;
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function setDescription(string $description): static
    {
        $this->description = $description;

        return $this;
    }

    public function getDescription(): string
    {
        return $this->description;
    }

    public function setCoordinates(string $coordinates): static
    {
        $this->coordinates = $coordinates;

        return $this;
    }

    public function getCoordinates(): string
    {
        return $this->coordinates;
    }

    public function setOffsetX(int $offsetX): static
    {
        $this->offsetX = $offsetX;

        return $this;
    }

    public function getOffsetX(): int
    {
        return $this->offsetX;
    }

    public function setOffsetY(int $offsetY): static
    {
        $this->offsetY = $offsetY;

        return $this;
    }

    public function getOffsetY(): int
    {
        return $this->offsetY;
    }

    public function setWidth(int $width): static
    {
        $this->width = $width;

        return $this;
    }

    public function getWidth(): int
    {
        return $this->width;
    }

    public function setHeight(int $height): static
    {
        $this->height = $height;

        return $this;
    }

    public function getHeight(): int
    {
        return $this->height;
    }

    public function getWorksheet(): ?Worksheet
    {
        return $this->worksheet;
    }

    /** Attach to a sheet, sending the image to the extension. */
    abstract public function setWorksheet(?Worksheet $worksheet, bool $overrideOld = false): static;
}
