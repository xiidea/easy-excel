<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Exception;
use EasyExcel\Native;

/**
 * Image drawing, PhpSpreadsheet style: configure, then attach with
 * setWorksheet() (which sends it to the extension).
 */
class Drawing
{
    private string $name = '';
    private string $description = '';
    private string $path = '';
    private string $coordinates = 'A1';
    private int $offsetX = 0;
    private int $offsetY = 0;
    private int $width = 0;
    private int $height = 0;
    private ?Worksheet $worksheet = null;

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

    public function setPath(string $path, bool $verifyFile = true): static
    {
        if ($verifyFile && !\is_file($path)) {
            throw new Exception("File $path not found!");
        }
        $this->path = $path;

        return $this;
    }

    public function getPath(): string
    {
        return $this->path;
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

    public function setOffsetY(int $offsetY): static
    {
        $this->offsetY = $offsetY;

        return $this;
    }

    public function setWidth(int $width): static
    {
        $this->width = $width;

        return $this;
    }

    public function setHeight(int $height): static
    {
        $this->height = $height;

        return $this;
    }

    public function setWorksheet(?Worksheet $worksheet, bool $overrideOld = false): static
    {
        if ($worksheet === null) {
            return $this;
        }
        if ($this->path === '') {
            throw new Exception('easy-excel: set the drawing path before attaching it to a worksheet');
        }
        $this->worksheet = $worksheet;
        Native::addImage(
            $worksheet->getParent()->getHandle(),
            $worksheet->getTitle(),
            $this->coordinates,
            [
                'path' => $this->path,
                'name' => $this->name,
                'offsetX' => $this->offsetX,
                'offsetY' => $this->offsetY,
                'width' => $this->width,
                'height' => $this->height,
            ],
        );

        return $this;
    }

    public function getWorksheet(): ?Worksheet
    {
        return $this->worksheet;
    }
}
