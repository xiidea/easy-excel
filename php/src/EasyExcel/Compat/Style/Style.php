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

    public function __construct(private Worksheet $worksheet, private string $range)
    {
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
        return new NumberFormat($this->worksheet, $this->range);
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
        Native::applyStyle(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->range,
            $spec,
        );
    }
}
