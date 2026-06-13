<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet\AutoFilter\Column;

use EasyExcel\Compat\Worksheet\AutoFilter\Column;

/**
 * One auto-filter column rule. Maps PhpSpreadsheet's rule operators onto the
 * excelize filter-expression syntax (x <op> value).
 */
class Rule
{
    public const AUTOFILTER_COLUMN_RULE_EQUAL = 'equal';
    public const AUTOFILTER_COLUMN_RULE_NOTEQUAL = 'notEqual';
    public const AUTOFILTER_COLUMN_RULE_GREATERTHAN = 'greaterThan';
    public const AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL = 'greaterThanOrEqual';
    public const AUTOFILTER_COLUMN_RULE_LESSTHAN = 'lessThan';
    public const AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL = 'lessThanOrEqual';

    private const OPERATORS = [
        self::AUTOFILTER_COLUMN_RULE_EQUAL => '==',
        self::AUTOFILTER_COLUMN_RULE_NOTEQUAL => '!=',
        self::AUTOFILTER_COLUMN_RULE_GREATERTHAN => '>',
        self::AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL => '>=',
        self::AUTOFILTER_COLUMN_RULE_LESSTHAN => '<',
        self::AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL => '<=',
    ];

    private string $operator = self::AUTOFILTER_COLUMN_RULE_EQUAL;
    private mixed $value = '';

    public function __construct(private Column $parent)
    {
    }

    public function setRule(string $operator, mixed $value): static
    {
        $this->operator = $operator;
        $this->value = $value;
        $this->parent->ruleChanged();

        return $this;
    }

    public function setValue(mixed $value): static
    {
        $this->value = $value;
        $this->parent->ruleChanged();

        return $this;
    }

    public function getValue(): mixed
    {
        return $this->value;
    }

    public function getOperator(): string
    {
        return $this->operator;
    }

    /** @internal the excelize expression fragment, e.g. "x > 2000" */
    public function toExpression(): string
    {
        $op = self::OPERATORS[$this->operator] ?? '==';

        return 'x ' . $op . ' ' . $this->value;
    }
}
