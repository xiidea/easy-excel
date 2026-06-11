<?php

declare(strict_types=1);

use EasyExcel\Compat\Shared\Date;

return [
    'date: unix epoch is serial 25569' => function (): void {
        T::same(25569.0, Date::formattedPHPToExcel(1970, 1, 1));
        T::same(25569.0, Date::PHPToExcel(new DateTime('1970-01-01 00:00:00', new DateTimeZone('UTC'))));
    },

    'date: known datetime serial' => function (): void {
        $dt = new DateTime('2023-06-15 12:00:00', new DateTimeZone('UTC'));
        T::same(45092.5, Date::dateTimeToExcel($dt));
    },

    'date: 1900 leap-year bug parity' => function (): void {
        // Excel: 1900-02-28 = 59, fake 1900-02-29 = 60, 1900-03-01 = 61
        T::same(59.0, Date::formattedPHPToExcel(1900, 2, 28));
        T::same(61.0, Date::formattedPHPToExcel(1900, 3, 1));
        T::same('1900-02-28', Date::excelToDateTimeObject(59)->format('Y-m-d'));
        T::same('1900-03-01', Date::excelToDateTimeObject(61)->format('Y-m-d'));
    },

    'date: serial -> DateTime round-trip' => function (): void {
        foreach ([25569.0, 45092.5, 44927.25, 2.0] as $serial) {
            $dt = Date::excelToDateTimeObject($serial);
            T::same($serial, Date::dateTimeToExcel($dt), "round-trip of $serial");
        }
    },

    'date: time-only serial' => function (): void {
        T::same('06:00:00', Date::excelToDateTimeObject(0.25)->format('H:i:s'));
    },

    'date: 1904 calendar' => function (): void {
        try {
            Date::setExcelCalendar(Date::CALENDAR_MAC_1904);
            T::same(0.0, Date::formattedPHPToExcel(1904, 1, 1));
            T::same('1904-01-02', Date::excelToDateTimeObject(1)->format('Y-m-d'));
        } finally {
            Date::setExcelCalendar(Date::CALENDAR_WINDOWS_1900);
        }
    },

    'date: string and timestamp inputs' => function (): void {
        T::same(45092.5, Date::PHPToExcel('2023-06-15 12:00:00'));
        T::same(25569.0, Date::PHPToExcel(0), 'unix timestamp 0');
        T::same(false, Date::PHPToExcel('not a date'));
    },
];
