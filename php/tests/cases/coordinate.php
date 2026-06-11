<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Exception;

return [
    'coordinate: columnIndexFromString' => function (): void {
        T::same(1, Coordinate::columnIndexFromString('A'));
        T::same(26, Coordinate::columnIndexFromString('Z'));
        T::same(27, Coordinate::columnIndexFromString('AA'));
        T::same(702, Coordinate::columnIndexFromString('ZZ'));
        T::same(703, Coordinate::columnIndexFromString('AAA'));
        T::same(2, Coordinate::columnIndexFromString('b'), 'lowercase accepted');
    },

    'coordinate: stringFromColumnIndex round-trip' => function (): void {
        foreach ([1, 26, 27, 52, 702, 703, 16384] as $i) {
            T::same($i, Coordinate::columnIndexFromString(Coordinate::stringFromColumnIndex($i)));
        }
        T::same('XFD', Coordinate::stringFromColumnIndex(16384));
    },

    'coordinate: coordinateFromString and indexes' => function (): void {
        T::same(['A', '1'], Coordinate::coordinateFromString('A1'));
        T::same(['$B', '$7'], Coordinate::coordinateFromString('$B$7'), 'absolute refs keep $');
        T::same([2, 7], Coordinate::indexesFromString('$B$7'));
        T::same([28, 105], Coordinate::indexesFromString('AB105'));
    },

    'coordinate: invalid input throws' => function (): void {
        T::throws(Exception::class, static fn () => Coordinate::coordinateFromString('A1:B2'), 'range rejected');
        T::throws(Exception::class, static fn () => Coordinate::coordinateFromString(''), 'empty rejected');
        T::throws(Exception::class, static fn () => Coordinate::coordinateFromString('123'), 'row-only rejected');
        T::throws(Exception::class, static fn () => Coordinate::stringFromColumnIndex(0), 'zero index rejected');
    },

    'coordinate: rangeBoundaries and dimension' => function (): void {
        T::same([[2, 2], [4, 4]], Coordinate::rangeBoundaries('B2:D4'));
        T::same([[3, 5], [3, 5]], Coordinate::rangeBoundaries('C5'), 'single cell expands');
        T::same([3, 3], Coordinate::rangeDimension('B2:D4'));
    },
];
