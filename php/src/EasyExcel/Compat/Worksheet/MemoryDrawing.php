<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Exception;
use EasyExcel\Native;

/**
 * In-memory (GD) drawing, PhpSpreadsheet style (wave 4.4): hold a GD image
 * resource, render it to PNG bytes on attach, and send them to the extension
 * as base64 (no temp file). Requires ext-gd.
 */
class MemoryDrawing
{
    public const RENDERING_DEFAULT = 'imagepng';
    public const RENDERING_PNG = 'imagepng';
    public const RENDERING_GIF = 'imagegif';
    public const RENDERING_JPEG = 'imagejpeg';

    public const MIMETYPE_DEFAULT = 'image/png';
    public const MIMETYPE_PNG = 'image/png';
    public const MIMETYPE_GIF = 'image/gif';
    public const MIMETYPE_JPEG = 'image/jpeg';

    private string $name = '';
    private string $coordinates = 'A1';
    private int $offsetX = 0;
    private int $offsetY = 0;
    private int $width = 0;
    private int $height = 0;
    private mixed $imageResource = null;
    private string $renderingFunction = self::RENDERING_DEFAULT;
    private string $mimeType = self::MIMETYPE_DEFAULT;
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

    public function setImageResource(mixed $value): static
    {
        $this->imageResource = $value;
        if ($value !== null) {
            $this->width = \imagesx($value);
            $this->height = \imagesy($value);
        }

        return $this;
    }

    public function getImageResource(): mixed
    {
        return $this->imageResource;
    }

    public function setRenderingFunction(string $value): static
    {
        $this->renderingFunction = $value;

        return $this;
    }

    public function setMimeType(string $value): static
    {
        $this->mimeType = $value;

        return $this;
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
        if ($this->imageResource === null) {
            throw new Exception('easy-excel: set the image resource before attaching the drawing');
        }
        if (!\function_exists('imagepng')) {
            throw new Exception('easy-excel: ext-gd is required for MemoryDrawing');
        }
        $this->worksheet = $worksheet;
        [$data, $extension] = $this->render();
        Native::addImageBytes(
            $worksheet->getParent()->getHandle(),
            $worksheet->getTitle(),
            $this->coordinates,
            [
                'data' => $data,
                'extension' => $extension,
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

    /** @return array{0: string, 1: string} [base64 data, extension] */
    private function render(): array
    {
        $extension = match ($this->renderingFunction) {
            self::RENDERING_JPEG => '.jpeg',
            self::RENDERING_GIF => '.gif',
            default => '.png',
        };
        \ob_start();
        ($this->renderingFunction)($this->imageResource);
        $bytes = (string) \ob_get_clean();

        return [\base64_encode($bytes), $extension];
    }
}
