<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Exception;
use EasyExcel\Native;

/**
 * Image drawing, PhpSpreadsheet style: configure, then attach with
 * setWorksheet() (which sends it to the extension).
 */
class Drawing extends BaseDrawing
{
    private string $path = '';





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







    public function setWorksheet(?Worksheet $worksheet, bool $overrideOld = false): static
    {
        if ($worksheet === null) {
            return $this;
        }
        if ($this->path === '') {
            throw new Exception('easy-excel: set the drawing path before attaching it to a worksheet');
        }
        $this->worksheet = $worksheet;
        
        // Track the drawing in PHP side so HTML writer can fetch it
        $worksheet->addDrawing($this);

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

}
