<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/**
 * Chart axis, PhpSpreadsheet style (wave 5.4).
 *
 * Mapped onto excelize.ChartAxis, which covers label suppression, bounds,
 * major unit, log scale, reverse order, number format and label font colour.
 * Options excelize has no field for — tick-mark style, crossing point, axis
 * orientation, time units, display units — are accepted and ignored rather
 * than thrown, so a report that sets them still produces its chart
 * (COMPAT.md divergence 30).
 */
class Axis extends ChartColor
{
    public const AXIS_LABELS_LOW = 'low';

    public const AXIS_LABELS_HIGH = 'high';

    public const AXIS_LABELS_NEXT_TO = 'nextTo';

    public const AXIS_LABELS_NONE = 'none';

    public const TICK_MARK_NONE = 'none';

    public const TICK_MARK_INSIDE = 'in';

    public const TICK_MARK_OUTSIDE = 'out';

    public const TICK_MARK_CROSS = 'cross';

    public const HORIZONTAL_CROSSES_AUTOZERO = 'autoZero';

    public const HORIZONTAL_CROSSES_MAXIMUM = 'max';

    public const AXIS_ORIENTATION_MIN_MAX = 'minMax';

    public const AXIS_ORIENTATION_MAX_MIN = 'maxMin';

    private string $axisLabels = '';

    private ?float $minimum = null;

    private ?float $maximum = null;

    private ?float $majorUnit = null;

    private ?float $minorUnit = null;

    private ?float $logBase = null;

    private string $orientation = '';

    private string $numberFormat = '';

    private ?GridLines $majorGridlines = null;

    private ?GridLines $minorGridlines = null;

    /**
     * Positional signature mirrors PhpSpreadsheet exactly so existing calls —
     * commonly `setAxisOptionsProperties('none')` — bind correctly.
     */
    public function setAxisOptionsProperties(
        string $axisLabels,
        ?string $horizontalCrossesValue = null,
        ?string $horizontalCrosses = null,
        ?string $axisOrientation = null,
        ?string $majorTmt = null,
        ?string $minorTmt = null,
        null|float|int|string $minimum = null,
        null|float|int|string $maximum = null,
        null|float|int|string $majorUnit = null,
        null|float|int|string $minorUnit = null,
        null|float|int|string $textRotation = null,
        ?string $hidden = null,
        ?string $baseTimeUnit = null,
        ?string $majorTimeUnit = null,
        ?string $minorTimeUnit = null,
        null|float|int|string $logBase = null,
        ?string $dispUnitsBuiltIn = null,
    ): void {
        $this->axisLabels = $axisLabels;
        if ($axisOrientation !== null) {
            $this->orientation = $axisOrientation;
        }
        $num = static fn (null|float|int|string $v): ?float => ($v === null || $v === '') ? null : (float) $v;
        $this->minimum = $num($minimum) ?? $this->minimum;
        $this->maximum = $num($maximum) ?? $this->maximum;
        $this->majorUnit = $num($majorUnit) ?? $this->majorUnit;
        $this->minorUnit = $num($minorUnit) ?? $this->minorUnit;
        $this->logBase = $num($logBase) ?? $this->logBase;
    }

    public function setAxisNumberProperties(
        string $formatCode,
        ?bool $sourceLinked = null,
        int|string|null $axisNumber = null,
    ): void {
        $this->numberFormat = $formatCode;
    }

    public function getAxisNumberFormat(): string
    {
        return $this->numberFormat;
    }

    public function setFillParameters(?string $color, null|float|int|string $alpha = null, ?string $type = null): void
    {
        $this->setColorProperties($color, $alpha, $type);
    }

    public function getFillProperty(string $property = 'value'): string
    {
        return match ($property) {
            'value' => $this->getValue(),
            'type' => $this->getType(),
            default => '',
        };
    }

    public function setMajorGridlines(?GridLines $gridlines): static
    {
        $this->majorGridlines = $gridlines;

        return $this;
    }

    public function getMajorGridlines(): ?GridLines
    {
        return $this->majorGridlines;
    }

    public function setMinorGridlines(?GridLines $gridlines): static
    {
        $this->minorGridlines = $gridlines;

        return $this;
    }

    public function getMinorGridlines(): ?GridLines
    {
        return $this->minorGridlines;
    }

    public function getAxisOptionsProperty(string $property): ?string
    {
        $value = match ($property) {
            'axis_labels' => $this->axisLabels,
            'orientation' => $this->orientation,
            'minimum' => $this->minimum,
            'maximum' => $this->maximum,
            'major_unit' => $this->majorUnit,
            'minor_unit' => $this->minorUnit,
            'log_base' => $this->logBase,
            default => null,
        };

        return $value === null ? null : (string) $value;
    }

    /** @internal builds the wave-5.4 axis block of the native chart spec */
    public function buildSpec(): array
    {
        $spec = [];
        if ($this->axisLabels !== '') {
            $spec['labels'] = $this->axisLabels;
        }
        // Nulls are omitted rather than sent: the Go side distinguishes an
        // unset bound from a deliberate zero.
        foreach ([
            'minimum' => $this->minimum,
            'maximum' => $this->maximum,
            'majorUnit' => $this->majorUnit,
            'logBase' => $this->logBase,
        ] as $key => $value) {
            if ($value !== null) {
                $spec[$key] = $value;
            }
        }
        if ($this->orientation === self::AXIS_ORIENTATION_MAX_MIN) {
            $spec['reverseOrder'] = true;
        }
        if ($this->numberFormat !== '') {
            $spec['numFmt'] = $this->numberFormat;
        }
        if ($this->majorGridlines !== null) {
            $spec['majorGridlines'] = true;
        }
        if ($this->minorGridlines !== null) {
            $spec['minorGridlines'] = true;
        }

        return $spec;
    }
}
