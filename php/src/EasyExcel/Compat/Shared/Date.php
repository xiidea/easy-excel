<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Shared;

use EasyExcel\Compat\Exception;

/**
 * Port of PhpSpreadsheet's Shared\Date: Excel serial <-> PHP date
 * conversions including the 1900 leap-year bug and the 1904 Mac calendar.
 * Pure PHP — date math never crosses the extension boundary.
 */
class Date
{
    public const CALENDAR_WINDOWS_1900 = 1900;
    public const CALENDAR_MAC_1904 = 1904;

    protected static int $excelCalendar = self::CALENDAR_WINDOWS_1900;

    public static function setExcelCalendar(int $baseYear): bool
    {
        if ($baseYear === self::CALENDAR_WINDOWS_1900 || $baseYear === self::CALENDAR_MAC_1904) {
            self::$excelCalendar = $baseYear;

            return true;
        }

        return false;
    }

    public static function getExcelCalendar(): int
    {
        return self::$excelCalendar;
    }

    public static function excelToDateTimeObject(float|int $excelTimestamp, ?\DateTimeZone $timeZone = null): \DateTime
    {
        $tz = $timeZone ?? new \DateTimeZone('UTC');
        if (self::$excelCalendar === self::CALENDAR_WINDOWS_1900) {
            if ($excelTimestamp < 1) {
                // time-only value
                $baseDate = new \DateTime('1970-01-01', $tz);
            } else {
                // Excel treats 1900 as a leap year: serials below 60 sit
                // one day later than the real calendar
                $baseDate = ($excelTimestamp < 60)
                    ? new \DateTime('1899-12-31', $tz)
                    : new \DateTime('1899-12-30', $tz);
            }
        } else {
            $baseDate = new \DateTime('1904-01-01', $tz);
        }

        $days = \floor($excelTimestamp);
        $partDay = $excelTimestamp - $days;
        $hours = (int) \floor($partDay * 24);
        $partDay = $partDay * 24 - $hours;
        $minutes = (int) \floor($partDay * 60);
        $partDay = $partDay * 60 - $minutes;
        $seconds = (int) \round($partDay * 60);

        if ($days >= 0) {
            $days = '+' . $days;
        }
        $interval = $days . ' days';

        return $baseDate->modify($interval)->setTime($hours, $minutes, $seconds);
    }

    public static function excelToTimestamp(float|int $excelTimestamp, ?\DateTimeZone $timeZone = null): int
    {
        return (int) self::excelToDateTimeObject($excelTimestamp, $timeZone)->format('U');
    }

    public static function PHPToExcel(mixed $dateValue): float|bool
    {
        if ($dateValue instanceof \DateTimeInterface) {
            return self::dateTimeToExcel($dateValue);
        }
        if (\is_numeric($dateValue)) {
            return self::timestampToExcel((int) $dateValue);
        }
        if (\is_string($dateValue)) {
            return self::stringToExcel($dateValue);
        }

        return false;
    }

    public static function dateTimeToExcel(\DateTimeInterface $dateValue): float
    {
        return self::formattedPHPToExcel(
            (int) $dateValue->format('Y'),
            (int) $dateValue->format('m'),
            (int) $dateValue->format('d'),
            (int) $dateValue->format('H'),
            (int) $dateValue->format('i'),
            (int) $dateValue->format('s'),
        );
    }

    public static function timestampToExcel(int $unixTimestamp): float
    {
        return self::dateTimeToExcel(new \DateTime('@' . $unixTimestamp));
    }

    public static function stringToExcel(string $dateValue): float|bool
    {
        try {
            $dt = new \DateTime($dateValue, new \DateTimeZone('UTC'));
        } catch (\Exception) {
            return false;
        }

        return self::dateTimeToExcel($dt);
    }

    /** PhpSpreadsheet's Julian-day algorithm, ported verbatim. */
    public static function formattedPHPToExcel(int $year, int $month, int $day, int $hours = 0, int $minutes = 0, int $seconds = 0): float
    {
        if (self::$excelCalendar === self::CALENDAR_WINDOWS_1900) {
            // Excel believes 1900-02-29 existed
            $excel1900isLeapYear = ($year === 1900 && $month <= 2) ? 0 : 1;
            $myexcelBaseDate = 2415020;
        } else {
            $myexcelBaseDate = 2416481;
            $excel1900isLeapYear = 0;
        }

        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }
        $century = \intdiv($year, 100);
        $decade = $year % 100;

        $excelDate = \floor((146097 * $century) / 4)
            + \floor((1461 * $decade) / 4)
            + \floor((153 * $month + 2) / 5)
            + $day + 1721119 - $myexcelBaseDate + $excel1900isLeapYear;

        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

        return (float) $excelDate + $excelTime;
    }

    public static function isDateTime(mixed $value): bool
    {
        return $value instanceof \DateTimeInterface;
    }

    public static function convertIsoDate(mixed $value): float|int
    {
        if (!\is_string($value)) {
            throw new Exception('Non-string value supplied for Iso Date conversion');
        }
        $date = new \DateTime($value);

        return self::PHPToExcel($date) ?: 0;
    }
}
