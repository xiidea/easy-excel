<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Exception;
use EasyExcel\Compat\RichText\RichText;

/**
 * PhpSpreadsheet's StringValueBinder: binds every value as a string unless
 * conversion for its type is suppressed, in which case the value keeps its
 * native data type.
 *
 * Differences from the real binder, by design:
 *  - no UTF-8 sanitisation (the Go side stores strings verbatim);
 *  - setSetIgnoredErrors() is accepted for API parity but has no effect —
 *    the extension does not model per-cell ignored-error flags.
 */
class StringValueBinder extends DefaultValueBinder
{
    protected bool $convertNull = true;

    protected bool $convertBoolean = true;

    protected bool $convertNumeric = true;

    protected bool $convertFormula = true;

    protected bool $setIgnoredErrors = false;

    public function setSetIgnoredErrors(bool $setIgnoredErrors = false): self
    {
        $this->setIgnoredErrors = $setIgnoredErrors;

        return $this;
    }

    public function setNullConversion(bool $suppressConversion = false): self
    {
        $this->convertNull = $suppressConversion;

        return $this;
    }

    public function setBooleanConversion(bool $suppressConversion = false): self
    {
        $this->convertBoolean = $suppressConversion;

        return $this;
    }

    public function getBooleanConversion(): bool
    {
        return $this->convertBoolean;
    }

    public function setNumericConversion(bool $suppressConversion = false): self
    {
        $this->convertNumeric = $suppressConversion;

        return $this;
    }

    public function setFormulaConversion(bool $suppressConversion = false): self
    {
        $this->convertFormula = $suppressConversion;

        return $this;
    }

    public function setConversionForAllValueTypes(bool $suppressConversion = false): self
    {
        $this->convertNull = $suppressConversion;
        $this->convertBoolean = $suppressConversion;
        $this->convertNumeric = $suppressConversion;
        $this->convertFormula = $suppressConversion;

        return $this;
    }

    public function bindValue(Cell $cell, mixed $value): bool
    {
        if (\is_object($value)) {
            return $this->bindObjectValue($cell, $value);
        }
        if ($value !== null && !\is_scalar($value)) {
            throw new Exception('Unable to bind unstringable ' . \gettype($value));
        }

        if ($value === null && $this->convertNull === false) {
            $cell->setValueExplicit($value, DataType::TYPE_NULL);
        } elseif (\is_bool($value) && $this->convertBoolean === false) {
            $cell->setValueExplicit($value, DataType::TYPE_BOOL);
        } elseif ((\is_int($value) || \is_float($value)) && $this->convertNumeric === false) {
            $cell->setValueExplicit($value, DataType::TYPE_NUMERIC);
        } elseif (\is_string($value) && \strlen($value) > 1 && $value[0] === '=' && $this->convertFormula === false && parent::dataTypeForValue($value) === DataType::TYPE_FORMULA) {
            $cell->setValueExplicit($value, DataType::TYPE_FORMULA);
        } else {
            $cell->setValueExplicit((string) $value, DataType::TYPE_STRING);
        }

        return true;
    }

    protected function bindObjectValue(Cell $cell, object $value): bool
    {
        if ($value instanceof \DateTimeInterface) {
            $cell->setValueExplicit($value->format('Y-m-d H:i:s'), DataType::TYPE_STRING);
        } elseif ($value instanceof RichText) {
            $cell->setValueExplicit($value, DataType::TYPE_INLINE);
        } elseif ($value instanceof \Stringable) {
            $cell->setValueExplicit((string) $value, DataType::TYPE_STRING);
        } else {
            throw new Exception('Unable to bind unstringable object of type ' . $value::class);
        }

        return true;
    }
}
