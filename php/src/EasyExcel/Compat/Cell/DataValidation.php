<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/**
 * Data validation rule. Obtained bound via Cell::getDataValidation() (setters
 * push the full state) or built standalone and applied through
 * Worksheet::setDataValidation().
 */
class DataValidation
{
    public const TYPE_NONE = 'none';
    public const TYPE_CUSTOM = 'custom';
    public const TYPE_DATE = 'date';
    public const TYPE_DECIMAL = 'decimal';
    public const TYPE_LIST = 'list';
    public const TYPE_TEXTLENGTH = 'textLength';
    public const TYPE_TIME = 'time';
    public const TYPE_WHOLE = 'whole';

    public const STYLE_STOP = 'stop';
    public const STYLE_WARNING = 'warning';
    public const STYLE_INFORMATION = 'information';

    public const OPERATOR_BETWEEN = 'between';
    public const OPERATOR_EQUAL = 'equal';
    public const OPERATOR_GREATERTHAN = 'greaterThan';
    public const OPERATOR_GREATERTHANOREQUAL = 'greaterThanOrEqual';
    public const OPERATOR_LESSTHAN = 'lessThan';
    public const OPERATOR_LESSTHANOREQUAL = 'lessThanOrEqual';
    public const OPERATOR_NOTBETWEEN = 'notBetween';
    public const OPERATOR_NOTEQUAL = 'notEqual';

    private string $type = self::TYPE_NONE;
    private string $errorStyle = self::STYLE_STOP;
    private string $operator = self::OPERATOR_BETWEEN;
    private string $formula1 = '';
    private string $formula2 = '';
    private bool $allowBlank = false;
    private bool $showDropDown = false;
    private bool $showInputMessage = false;
    private bool $showErrorMessage = false;
    private string $errorTitle = '';
    private string $error = '';
    private string $promptTitle = '';
    private string $prompt = '';

    private ?Worksheet $worksheet = null;
    private string $range = '';

    /** @internal */
    public function bind(Worksheet $worksheet, string $range): static
    {
        $this->worksheet = $worksheet;
        $this->range = $range;

        return $this;
    }

    /** @internal full state in the shape extension/compat/validation.go expects */
    public function toSpec(): array
    {
        return [
            'type' => $this->type,
            'operator' => $this->operator,
            'formula1' => $this->formula1,
            'formula2' => $this->formula2,
            'allowBlank' => $this->allowBlank,
            'showDropDown' => $this->showDropDown,
            'showInputMessage' => $this->showInputMessage,
            'showErrorMessage' => $this->showErrorMessage,
            'errorStyle' => $this->errorStyle,
            'errorTitle' => $this->errorTitle,
            'error' => $this->error,
            'promptTitle' => $this->promptTitle,
            'prompt' => $this->prompt,
        ];
    }

    private function push(): void
    {
        if ($this->worksheet === null || $this->type === self::TYPE_NONE) {
            return;
        }
        Native::setValidation(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->range,
            $this->toSpec(),
        );
    }

    public function setType(string $type): static
    {
        $this->type = $type;
        $this->push();

        return $this;
    }

    public function getType(): string
    {
        return $this->type;
    }

    public function setErrorStyle(string $errorStyle): static
    {
        $this->errorStyle = $errorStyle;
        $this->push();

        return $this;
    }

    public function getErrorStyle(): string
    {
        return $this->errorStyle;
    }

    public function setOperator(string $operator): static
    {
        $this->operator = $operator;
        $this->push();

        return $this;
    }

    public function getOperator(): string
    {
        return $this->operator;
    }

    public function setFormula1(string $formula): static
    {
        $this->formula1 = $formula;
        $this->push();

        return $this;
    }

    public function getFormula1(): string
    {
        return $this->formula1;
    }

    public function setFormula2(string $formula): static
    {
        $this->formula2 = $formula;
        $this->push();

        return $this;
    }

    public function getFormula2(): string
    {
        return $this->formula2;
    }

    public function setAllowBlank(bool $allowBlank): static
    {
        $this->allowBlank = $allowBlank;
        $this->push();

        return $this;
    }

    public function getAllowBlank(): bool
    {
        return $this->allowBlank;
    }

    public function setShowDropDown(bool $showDropDown): static
    {
        $this->showDropDown = $showDropDown;
        $this->push();

        return $this;
    }

    public function getShowDropDown(): bool
    {
        return $this->showDropDown;
    }

    public function setShowInputMessage(bool $show): static
    {
        $this->showInputMessage = $show;
        $this->push();

        return $this;
    }

    public function getShowInputMessage(): bool
    {
        return $this->showInputMessage;
    }

    public function setShowErrorMessage(bool $show): static
    {
        $this->showErrorMessage = $show;
        $this->push();

        return $this;
    }

    public function getShowErrorMessage(): bool
    {
        return $this->showErrorMessage;
    }

    public function setErrorTitle(string $errorTitle): static
    {
        $this->errorTitle = $errorTitle;
        $this->push();

        return $this;
    }

    public function getErrorTitle(): string
    {
        return $this->errorTitle;
    }

    public function setError(string $error): static
    {
        $this->error = $error;
        $this->push();

        return $this;
    }

    public function getError(): string
    {
        return $this->error;
    }

    public function setPromptTitle(string $promptTitle): static
    {
        $this->promptTitle = $promptTitle;
        $this->push();

        return $this;
    }

    public function getPromptTitle(): string
    {
        return $this->promptTitle;
    }

    public function setPrompt(string $prompt): static
    {
        $this->prompt = $prompt;
        $this->push();

        return $this;
    }

    public function getPrompt(): string
    {
        return $this->prompt;
    }
}
