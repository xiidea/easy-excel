<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet\AutoFilter;

use EasyExcel\Compat\Worksheet\AutoFilter;
use EasyExcel\Compat\Worksheet\AutoFilter\Column\Rule;

/**
 * One auto-filter column. Rules join with AND/OR (excelize accepts a single
 * statement or two joined by 'and'/'or').
 */
class Column
{
    public const AUTOFILTER_FILTERTYPE_FILTER = 'filters';
    public const AUTOFILTER_FILTERTYPE_CUSTOMFILTER = 'customFilters';

    public const AUTOFILTER_COLUMN_JOIN_AND = 'and';
    public const AUTOFILTER_COLUMN_JOIN_OR = 'or';

    private string $filterType = self::AUTOFILTER_FILTERTYPE_FILTER;
    private string $join = self::AUTOFILTER_COLUMN_JOIN_OR;

    /** @var list<Rule> */
    private array $rules = [];

    public function __construct(private AutoFilter $parent, private string $columnIndex)
    {
    }

    public function getColumnIndex(): string
    {
        return $this->columnIndex;
    }

    public function setFilterType(string $filterType): static
    {
        $this->filterType = $filterType;

        return $this;
    }

    public function getFilterType(): string
    {
        return $this->filterType;
    }

    public function setJoin(string $join): static
    {
        $this->join = $join;

        return $this;
    }

    public function createRule(): Rule
    {
        $rule = new Rule($this);
        $this->rules[] = $rule;

        return $rule;
    }

    /** @return list<Rule> */
    public function getRules(): array
    {
        return $this->rules;
    }

    /** @internal called by a child Rule when its operator/value changes */
    public function ruleChanged(): void
    {
        $this->parent->applyColumns();
    }

    /** @internal the excelize expression for this column, or '' if no rules */
    public function toExpression(): string
    {
        if ($this->rules === []) {
            return '';
        }
        $fragments = \array_map(static fn (Rule $r): string => $r->toExpression(), \array_slice($this->rules, 0, 2));

        return \implode(' ' . $this->join . ' ', $fragments);
    }
}
