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
class MemoryDrawing extends BaseDrawing
{
    public const RENDERING_DEFAULT = 'imagepng';
    public const RENDERING_PNG = 'imagepng';
    public const RENDERING_GIF = 'imagegif';
    public const RENDERING_JPEG = 'imagejpeg';

    public const MIMETYPE_DEFAULT = 'image/png';
    public const MIMETYPE_PNG = 'image/png';
    public const MIMETYPE_GIF = 'image/gif';
    public const MIMETYPE_JPEG = 'image/jpeg';

    private mixed $imageResource = null;
    private string $renderingFunction = self::RENDERING_DEFAULT;
    private string $mimeType = self::MIMETYPE_DEFAULT;



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
