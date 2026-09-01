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
    private static ?IValueBinder $valueBinder = null;

    public function __construct(private Worksheet $worksheet, private string $coordinate)
    {
    }

    /** Custom binders run in PHP before values reach the write buffer. */
    public static function setValueBinder(IValueBinder $binder): void
    {
        self::$valueBinder = $binder;
    }

    public static function getValueBinder(): IValueBinder
    {
        return self::$valueBinder ??= new DefaultValueBinder();
    }

    /** @internal null when the Go-side default binding can be used as-is */
    public static function customValueBinder(): ?IValueBinder
    {
        return self::$valueBinder;
    }

    public function getValue(): mixed
    {
        return $this->worksheet->readCell($this->coordinate, Native::GET_RAW);
    }

    public function getCalculatedValue(bool $resetLog = true): mixed
    {
        $v = $this->worksheet->readCell($this->coordinate, Native::GET_CALCULATED);
        if (!\is_string($v) || !\is_numeric($v)) {
            return $v;
        }
        // Only cast when the string is the *canonical* rendering of the
        // number: TEXT(A1,"0000") returns "0042", which is_numeric() accepts
        // but which becomes 42 if cast — silently dropping the leading zeros
        // the formula exists to produce. Same for "1.50" and " 3".
        $number = $v + 0;
        if ((string) $number !== $v) {
            return $v;
        }

        return $number;
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
        $dv = new DataValidation();
        // hydrate from an existing rule covering this cell (loaded files and
        // rules set earlier this session)
        $handle = $this->worksheet->getParent()->getHandle();
        [$col, $row] = Coordinate::indexesFromString($this->coordinate);
        foreach (\EasyExcel\Native::getValidations($handle, $this->worksheet->getTitle()) as $entry) {
            foreach (\explode(' ', (string) $entry['sqref']) as $ref) {
                [[$c1, $r1], [$c2, $r2]] = Coordinate::rangeBoundaries($ref);
                if ($col >= $c1 && $col <= $c2 && $row >= $r1 && $row <= $r2) {
                    $dv->hydrate($entry['spec']);
                    break 2;
                }
            }
        }

        return $dv->bind($this->worksheet, $this->coordinate);
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
