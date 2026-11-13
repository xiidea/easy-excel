<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader\Xlsx;

use EasyExcel\Compat\Worksheet\Sparkline\Sparkline;
use EasyExcel\Compat\Worksheet\Sparkline\SparklineGroup;
use EasyExcel\Compat\Worksheet\Sparkline\SparklineType;
use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Reads a worksheet's `x14:sparklineGroups` extension into its sparkline
 * group collection, PhpSpreadsheet 5.10 style.
 *
 * Usable directly with any worksheet XML. Loaded Compat workbooks feed it
 * the sparkline <ext> the extension found in the source file (see
 * Worksheet::getSparklineGroupCollection()); groups read that way are
 * already in the file, so they are marked written and not re-sent on save.
 */
class Sparklines
{
    private const SPARKLINE_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
    private const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    private const X14 ='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    private const XM = 'http://schemas.microsoft.com/office/excel/2006/main';

    public function __construct(private readonly Worksheet $worksheet, private readonly \SimpleXMLElement $worksheetXml)
    {
    }

    public function load(): void
    {
        // by namespace, so a prefixed main namespace (<x:ext>) is found too;
        // the plain lookup covers XML without any namespace
        $extLst = $this->worksheetXml->children(self::MAIN)->extLst;
        if (!isset($extLst->ext)) {
            $extLst = $this->worksheetXml->extLst;
        }
        if (!isset($extLst->ext)) {
            return;
        }
        foreach ($extLst->ext as $ext) {
            if ((string) ($ext->attributes()['uri'] ?? '') !== self::SPARKLINE_URI) {
                continue;
            }
            $groups = $ext->children(self::X14)->sparklineGroups;
            if (!isset($groups)) {
                continue;
            }
            foreach ($groups->children(self::X14)->sparklineGroup ?? [] as $groupXml) {
                $group = self::readGroup($groupXml);
                if ($group !== null) {
                    $this->worksheet->addSparklineGroup($group);
                }
            }
        }
    }

    private static function readGroup(\SimpleXMLElement $groupXml): ?SparklineGroup
    {
        $group = new SparklineGroup();
        $attrs = $groupXml->attributes();

        $group->setType(SparklineType::tryFrom((string) ($attrs['type'] ?? 'line')) ?? SparklineType::Line);
        if (isset($attrs['lineWeight'])) {
            $group->setLineWeight((float) $attrs['lineWeight']);
        }
        if (isset($attrs['displayEmptyCellsAs'])) {
            $group->setDisplayEmptyCellsAs((string) $attrs['displayEmptyCellsAs']);
        }
        if (isset($attrs['minAxisType'])) {
            $group->setMinAxisType((string) $attrs['minAxisType']);
        }
        if (isset($attrs['maxAxisType'])) {
            $group->setMaxAxisType((string) $attrs['maxAxisType']);
        }
        if (isset($attrs['manualMin'])) {
            $group->setManualMin((float) $attrs['manualMin']);
        }
        if (isset($attrs['manualMax'])) {
            $group->setManualMax((float) $attrs['manualMax']);
        }
        $group->setDisplayMarkers(self::boolAttr($attrs, 'markers'))
            ->setDisplayHigh(self::boolAttr($attrs, 'high'))
            ->setDisplayLow(self::boolAttr($attrs, 'low'))
            ->setDisplayFirst(self::boolAttr($attrs, 'first'))
            ->setDisplayLast(self::boolAttr($attrs, 'last'))
            ->setDisplayNegative(self::boolAttr($attrs, 'negative'))
            ->setDisplayXAxis(self::boolAttr($attrs, 'displayXAxis'))
            ->setDisplayHidden(self::boolAttr($attrs, 'displayHidden'))
            ->setRightToLeft(self::boolAttr($attrs, 'rightToLeft'));

        // an absent colour element reads as null, like upstream
        $colors = $groupXml->children(self::X14);
        $group->setColorSeries(self::color($colors->colorSeries))
            ->setColorNegative(self::color($colors->colorNegative))
            ->setColorAxis(self::color($colors->colorAxis))
            ->setColorMarkers(self::color($colors->colorMarkers))
            ->setColorFirst(self::color($colors->colorFirst))
            ->setColorLast(self::color($colors->colorLast))
            ->setColorHigh(self::color($colors->colorHigh))
            ->setColorLow(self::color($colors->colorLow));

        $sparklines = $colors->sparklines;
        foreach (isset($sparklines) ? ($sparklines->children(self::X14)->sparkline ?? []) : [] as $sparklineXml) {
            $xm = $sparklineXml->children(self::XM);
            $location = isset($xm->sqref) ? (string) $xm->sqref : '';
            if ($location !== '') {
                $group->addSparkline(new Sparkline($location, isset($xm->f) ? (string) $xm->f : ''));
            }
        }

        // a group with no sparklines carries no information
        return $group->getSparklines() === [] ? null : $group;
    }

    /** OOXML booleans: "1"/"true" (any numeric non-zero), like upstream's BaseParserClass. */
    private static function boolAttr(?\SimpleXMLElement $attrs, string $name): bool
    {
        if ($attrs === null || !isset($attrs[$name])) {
            return false;
        }
        $value = (string) $attrs[$name];

        return \is_numeric($value) ? (bool) (float) $value : ($value === 'true' || $value === 'TRUE');
    }

    private static function color(?\SimpleXMLElement $colorXml): ?string
    {
        if ($colorXml === null || !isset($colorXml[0])) {
            return null;
        }
        $rgb = $colorXml->attributes()['rgb'] ?? null;

        return $rgb === null ? null : (string) $rgb;
    }
}
