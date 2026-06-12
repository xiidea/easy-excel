<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/**
 * Style facade for one range (Phase 2). Component setters send partial style
 * specs straight to the extension; the Go side keeps a per-sheet style log
 * and merges overlapping partial styles in application order, so chained
 * setters and repeated getStyle() calls layer exactly like PhpSpreadsheet's
 * supervisor model. Styles queued before their rows are written are inlined
 * into the stream — no flush here, and no degrade for the usual
 * "style header → bulk write" report.
 */
class Style
{
    private ?Font $font = null;
    private ?Fill $fill = null;
    private ?Borders $borders = null;
    private ?Alignment $alignment = null;
    private ?Protection $protection = null;

    /** @var array<string, mixed> collected spec when detached (conditional styles) */
    private array $collected = [];

    public function __construct(private ?Worksheet $worksheet, private string $range)
    {
    }

    /**
     * A detached Style collects its spec locally instead of styling cells;
     * used by Conditional::getStyle().
     */
    public static function detached(): self
    {
        return new self(null, '');
    }

    /** @internal @return array<string, mixed> */
    public function getCollectedSpec(): array
    {
        return $this->collected;
    }

    public function getFont(): Font
    {
        return $this->font ??= new Font($this);
    }

    public function getFill(): Fill
    {
        return $this->fill ??= new Fill($this);
    }

    public function getBorders(): Borders
    {
        return $this->borders ??= new Borders($this);
    }

    public function getAlignment(): Alignment
    {
        return $this->alignment ??= new Alignment($this);
    }

    public function getProtection(): Protection
    {
        return $this->protection ??= new Protection($this);
    }

    public function getNumberFormat(): NumberFormat
    {
        if ($this->worksheet === null) {
            return new NumberFormat(null, '', $this);
        }

        return new NumberFormat($this->worksheet, $this->range);
    }

    /**
     * Replaces the conditional-formatting rules for this style's range.
     *
     * @param list<Conditional> $conditionalStyles
     */
    public function setConditionalStyles(array $conditionalStyles): static
    {
        if ($this->worksheet === null) {
            throw new \EasyExcel\Compat\Exception('easy-excel: conditional styles need a worksheet range');
        }
        Native::setConditional(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->range,
            \array_map(static fn (Conditional $c): array => $c->toSpec(), $conditionalStyles),
        );

        return $this;
    }

    /** @param array<string, mixed> $styleArray PhpSpreadsheet's nested style array */
    public function applyFromArray(array $styleArray, bool $advancedBorders = true): static
    {
        $this->send($styleArray);

        return $this;
    }

    public function getQuotePrefix(): bool
    {
        return false;
    }

    /** @internal one partial spec component, e.g. ['font' => ['bold' => true]] */
    public function mergeComponent(string $component, array $values): void
    {
        $this->send([$component => $values]);
    }

    /** @param array<string, mixed> $spec */
    private function send(array $spec): void
    {
        if ($spec === []) {
            return;
        }
        if ($this->worksheet === null) {
            $this->collected = self::deepMerge($this->collected, $spec);

            return;
        }
        Native::applyStyle(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->range,
            $spec,
        );
    }

    private static function deepMerge(array $base, array $patch): array
    {
        foreach ($patch as $k => $v) {
            if (\is_array($v) && isset($base[$k]) && \is_array($base[$k])) {
                $base[$k] = self::deepMerge($base[$k], $v);
            } else {
                $base[$k] = $v;
            }
        }

        return $base;
    }
}
