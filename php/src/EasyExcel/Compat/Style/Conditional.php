<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/**
 * One conditional-formatting rule, applied through
 * Style::setConditionalStyles(). getStyle() returns a detached Style that
 * collects the rule's formatting instead of styling cells.
 */
class Conditional
{
    public const CONDITION_NONE = 'none';
    public const CONDITION_CELLIS = 'cellIs';
    public const CONDITION_CONTAINSTEXT = 'containsText';
    public const CONDITION_EXPRESSION = 'expression';
    public const CONDITION_COLORSCALE = 'colorScale';
    public const CONDITION_DATABAR = 'dataBar';

    public const OPERATOR_NONE = '';
    public const OPERATOR_BEGINSWITH = 'beginsWith';
    public const OPERATOR_BETWEEN = 'between';
    public const OPERATOR_CONTAINSTEXT = 'containsText';
    public const OPERATOR_ENDSWITH = 'endsWith';
    public const OPERATOR_EQUAL = 'equal';
    public const OPERATOR_GREATERTHAN = 'greaterThan';
    public const OPERATOR_GREATERTHANOREQUAL = 'greaterThanOrEqual';
    public const OPERATOR_LESSTHAN = 'lessThan';
    public const OPERATOR_LESSTHANOREQUAL = 'lessThanOrEqual';
    public const OPERATOR_NOTBETWEEN = 'notBetween';
    public const OPERATOR_NOTCONTAINS = 'notContains';
    public const OPERATOR_NOTEQUAL = 'notEqual';

    private string $conditionType = self::CONDITION_NONE;
    private string $operatorType = self::OPERATOR_NONE;

    /** @var list<scalar> */
    private array $conditions = [];

    private ?Style $style = null;
    private bool $stopIfTrue = false;

    /** @var array{minColor?: string, midColor?: string, maxColor?: string} */
    private array $colorScale = [];

    private string $dataBarColor = '';

    public function setConditionType(string $type): static
    {
        $this->conditionType = $type;

        return $this;
    }

    public function getConditionType(): string
    {
        return $this->conditionType;
    }

    public function setOperatorType(string $type): static
    {
        $this->operatorType = $type;

        return $this;
    }

    public function getOperatorType(): string
    {
        return $this->operatorType;
    }

    /** @param scalar|list<scalar> $conditions */
    public function setConditions(mixed $conditions): static
    {
        $this->conditions = \is_array($conditions) ? \array_values($conditions) : [$conditions];

        return $this;
    }

    public function addCondition(mixed $condition): static
    {
        $this->conditions[] = $condition;

        return $this;
    }

    /** @return list<scalar> */
    public function getConditions(): array
    {
        return $this->conditions;
    }

    public function getStyle(): Style
    {
        return $this->style ??= Style::detached();
    }

    public function setStopIfTrue(bool $stop): static
    {
        $this->stopIfTrue = $stop;

        return $this;
    }

    /** easy-excel extra: 2/3-color scale (PhpSpreadsheet uses ConditionalColorScale objects) */
    public function setColorScale(string $minColor, string $maxColor, ?string $midColor = null): static
    {
        $this->conditionType = self::CONDITION_COLORSCALE;
        $this->colorScale = ['minColor' => $minColor, 'maxColor' => $maxColor];
        if ($midColor !== null) {
            $this->colorScale['midColor'] = $midColor;
        }

        return $this;
    }

    /** easy-excel extra: solid data bar */
    public function setDataBar(string $color): static
    {
        $this->conditionType = self::CONDITION_DATABAR;
        $this->dataBarColor = $color;

        return $this;
    }

    /** @internal the rule in the shape extension/compat/conditional.go expects */
    public function toSpec(): array
    {
        $spec = [
            'type' => $this->conditionType,
            'operator' => $this->operatorType,
            'conditions' => $this->conditions,
            'stopIfTrue' => $this->stopIfTrue,
        ];
        if ($this->style !== null && $this->style->getCollectedSpec() !== []) {
            $spec['style'] = $this->style->getCollectedSpec();
        }
        if ($this->colorScale !== []) {
            $spec['colorScale'] = $this->colorScale;
        }
        if ($this->dataBarColor !== '') {
            $spec['dataBar'] = ['color' => $this->dataBarColor];
        }

        return $spec;
    }
}
