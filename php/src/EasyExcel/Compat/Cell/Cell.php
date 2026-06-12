<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/**
 * Lightweight cell facade: holds only coordinates, never cell data —
 * the data lives in Go (PLAN.md B2). Cheap to create and discard.
 */
class Cell
{
    public function __construct(private Worksheet $worksheet, private string $coordinate)
    {
    }

    public function getValue(): mixed
    {
        return $this->worksheet->readCell($this->coordinate, Native::GET_RAW);
    }

    public function getCalculatedValue(bool $resetLog = true): mixed
    {
        $v = $this->worksheet->readCell($this->coordinate, Native::GET_CALCULATED);
        if (\is_string($v) && \is_numeric($v)) {
            return $v + 0;
        }

        return $v;
    }

    public function getFormattedValue(): string
    {
        return (string) $this->worksheet->readCell($this->coordinate, Native::GET_FORMATTED);
    }

    public function setValue(mixed $value): static
    {
        $this->worksheet->setCellValue($this->coordinate, $value);

        return $this;
    }

    public function setValueExplicit(mixed $value, string $dataType = DataType::TYPE_STRING): static
    {
        $this->worksheet->setCellValueExplicit($this->coordinate, $value, $dataType);

        return $this;
    }

    public function getCoordinate(): string
    {
        return $this->coordinate;
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }

    public function getHyperlink(): Hyperlink
    {
        return (new Hyperlink())->bind($this->worksheet, $this->coordinate);
    }

    public function setHyperlink(?Hyperlink $hyperlink): static
    {
        $this->worksheet->setHyperlink($this->coordinate, $hyperlink);

        return $this;
    }

    public function getStyle(): \EasyExcel\Compat\Style\Style
    {
        return $this->worksheet->getStyle($this->coordinate);
    }

    public function getDataValidation(): DataValidation
    {
        return (new DataValidation())->bind($this->worksheet, $this->coordinate);
    }

    public function setDataValidation(?DataValidation $dataValidation): static
    {
        $this->worksheet->setDataValidation($this->coordinate, $dataValidation);

        return $this;
    }

    public function getDataType(): string
    {
        $v = $this->getValue();

        return match (true) {
            $v === null => DataType::TYPE_NULL,
            \is_bool($v) => DataType::TYPE_BOOL,
            \is_int($v), \is_float($v) => DataType::TYPE_NUMERIC,
            \is_string($v) && \strlen($v) > 1 && $v[0] === '=' => DataType::TYPE_FORMULA,
            default => DataType::TYPE_STRING,
        };
    }
}
