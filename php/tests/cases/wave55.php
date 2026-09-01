<?php

declare(strict_types=1);

use EasyExcel\UnsupportedApiException;

use function EasyExcel\aliasAction;

/**
 * Wave 5.5 pins the *contract* around the by-design exclusions rather than
 * adding features: each must fail loudly, say why, and point at the documented
 * escape hatch. A later wave that quietly stubs one of these — or that breaks
 * the fallback deferral — should fail here.
 */
const W55_EXCLUDED = [
    'PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart',
    'PhpOffice\PhpSpreadsheet\Shared\XMLWriter',
    'PhpOffice\PhpSpreadsheet\Chart\Renderer\JpGraph',
    'PhpOffice\PhpSpreadsheet\Style\ConditionalFormatting\MergedCellStyle',
];

return [
    'wave55: every by-design exclusion throws in strict mode' => function (): void {
        foreach (W55_EXCLUDED as $class) {
            T::throws(
                UnsupportedApiException::class,
                static fn () => \class_exists($class),
                "$class must fail loudly, not resolve to something different",
            );
        }
    },

    'wave55: the failure message points at the escape hatch' => function (): void {
        foreach (W55_EXCLUDED as $class) {
            try {
                \class_exists($class);
                T::ok(false, "$class did not throw");
            } catch (UnsupportedApiException $e) {
                $msg = $e->getMessage();
                T::ok(\str_contains($msg, $class), 'names the class that failed');
                T::ok(\str_contains($msg, 'fallback'), 'names EASY_EXCEL_ALIAS=fallback');
                T::ok(\str_contains($msg, 'off'), 'names EASY_EXCEL_ALIAS=off');
            }
        }
    },

    'wave55: fallback mode defers every exclusion to the real package' => function (): void {
        foreach (W55_EXCLUDED as $class) {
            T::same(
                ['defer', null],
                aliasAction('fallback', $class),
                "$class must defer under fallback, not throw",
            );
        }
    },

    'wave55: strict mode still throws for the same classes' => function (): void {
        foreach (W55_EXCLUDED as $class) {
            [$action] = aliasAction('strict', $class);
            T::same('throw', $action, "$class throws under strict");
        }
    },

    'wave55: off mode leaves every PhpOffice name alone' => function (): void {
        foreach (W55_EXCLUDED as $class) {
            // 'off' registers no autoloader at all, so resolution is upstream's
            // job; aliasAction still reports the per-class decision.
            [$action] = aliasAction('off', $class);
            T::same('defer', $action, "$class defers under off");
        }
    },

    'wave55: the exclusions are genuinely absent from the Compat tree' => function (): void {
        // The alias surface is derived by scanning the directory, so a stray
        // file would silently un-exclude one of these.
        $base = \dirname(__DIR__, 2) . '/src/EasyExcel/Compat/';
        foreach (W55_EXCLUDED as $class) {
            $rel = \str_replace('\\', '/', \substr($class, \strlen('PhpOffice\PhpSpreadsheet\\')));
            T::ok(!\is_file($base . $rel . '.php'), "$rel.php must not exist");
        }
    },

    'wave55: the documented alternatives do exist' => function (): void {
        // Each exclusion is documented with a supported alternative; those
        // must actually be there or the guidance is wrong.
        T::ok(\method_exists('EasyExcel\Compat\Spreadsheet', 'copySheet'), 'copySheet for subclassing');
        T::ok(\method_exists('EasyExcel\Compat\Chart\Chart', 'render'), 'render() for the no-image branch');
        T::ok(\method_exists('EasyExcel\Compat\Settings', 'setChartRenderer'), 'setChartRenderer accepted');
        T::ok(\method_exists('EasyExcel\Compat\Style\Style', 'getConditionalStyles'), 'conditional rules readable');
    },

    'wave55: Chart::render returns false rather than throwing' => function (): void {
        // The documented contract: callers take their existing no-image path.
        $chart = new \EasyExcel\Compat\Chart\Chart('c');
        T::same(false, $chart->render(), 'no renderer, no exception');
    },

    'wave55: setChartRenderer is accepted, not thrown' => function (): void {
        \EasyExcel\Compat\Settings::reset();
        \EasyExcel\Compat\Settings::setChartRenderer('PhpOffice\PhpSpreadsheet\Chart\Renderer\JpGraph');
        T::same(
            'PhpOffice\PhpSpreadsheet\Chart\Renderer\JpGraph',
            \EasyExcel\Compat\Settings::getChartRenderer(),
            'accepted so chart-bearing reports still generate their xlsx',
        );
        \EasyExcel\Compat\Settings::reset();
    },
];
