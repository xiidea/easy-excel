<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/** Number-format component; format codes are PhpSpreadsheet's constants. */
class NumberFormat
{
    public const FORMAT_GENERAL = 'General';
    public const FORMAT_TEXT = '@';
    public const FORMAT_NUMBER = '0';
    public const FORMAT_NUMBER_0 = '0.0';
    public const FORMAT_NUMBER_00 = '0.00';
    public const FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    public const FORMAT_PERCENTAGE = '0%';
    public const FORMAT_PERCENTAGE_0 = '0.0%';
    public const FORMAT_PERCENTAGE_00 = '0.00%';
    public const FORMAT_DATE_YYYYMMDD = 'yyyy-mm-dd';
    public const FORMAT_DATE_DDMMYYYY = 'dd/mm/yyyy';
    public const FORMAT_DATE_DATETIME = 'd/m/yy h:mm';
    public const FORMAT_DATE_TIME1 = 'h:mm AM/PM';
    public const FORMAT_DATE_TIME4 = 'h:mm:ss';
    public const FORMAT_DATE_XLSX14 = 'mm-dd-yy';
    public const FORMAT_DATE_XLSX22 = 'm/d/yy h:mm';
    public const FORMAT_CURRENCY_USD = '$#,##0.00';
    public const FORMAT_CURRENCY_USD_INTEGER = '$#,##0';
    public const FORMAT_CURRENCY_EUR = '€#,##0.00';
    public const FORMAT_ACCOUNTING_USD = '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';

    private string $formatCode = self::FORMAT_GENERAL;

    public function __construct(
        private ?Worksheet $worksheet,
        private string $range,
        private ?Style $detachedParent = null,
    ) {
    }

    public function setFormatCode(string $formatCode): static
    {
        $this->formatCode = $formatCode;
        if ($this->worksheet === null) {
            // detached (conditional style): collect on the parent Style
            $this->detachedParent?->mergeComponent('numberFormat', ['formatCode' => $formatCode]);

            return $this;
        }
        // no flush: the Go style log is coordinate-based and keeps formats
        // queued before their rows stream-compatible
        Native::setNumberFormat(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->range,
            $formatCode,
        );

        return $this;
    }

    public function getFormatCode(): string
    {
        return $this->formatCode;
    }
}
