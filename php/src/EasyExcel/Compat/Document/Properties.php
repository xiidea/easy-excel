<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Document;

use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

/**
 * Document metadata (docProps/core.xml + company in app.xml + custom.xml).
 * Setters push the accumulated state; excelize exposes no manager field, so
 * setManager is accepted but only kept PHP-side (COMPAT.md).
 */
class Properties
{
    public const PROPERTY_TYPE_BOOLEAN = 'b';
    public const PROPERTY_TYPE_INTEGER = 'i';
    public const PROPERTY_TYPE_FLOAT = 'f';
    public const PROPERTY_TYPE_DATE = 'd';
    public const PROPERTY_TYPE_STRING = 's';
    public const PROPERTY_TYPE_UNKNOWN = 'u';

    private const VALID_PROPERTY_TYPE_LIST = [
        self::PROPERTY_TYPE_BOOLEAN,
        self::PROPERTY_TYPE_INTEGER,
        self::PROPERTY_TYPE_FLOAT,
        self::PROPERTY_TYPE_DATE,
        self::PROPERTY_TYPE_STRING,
    ];

    private array $props = [];

    private string $manager = '';

    private int $created;

    private int $modified;

    /** @var array<string, array{value: mixed, type: string}> */
    private array $customProperties = [];

    public function __construct(private Spreadsheet $spreadsheet)
    {
        $this->created = \time();
        $this->modified = $this->created;
    }

    public function setTitle(string $title): static
    {
        return $this->set('title', $title);
    }

    public function getTitle(): string
    {
        return $this->props['title'] ?? '';
    }

    public function setSubject(string $subject): static
    {
        return $this->set('subject', $subject);
    }

    public function getSubject(): string
    {
        return $this->props['subject'] ?? '';
    }

    public function setCreator(string $creator): static
    {
        return $this->set('creator', $creator);
    }

    public function getCreator(): string
    {
        return $this->props['creator'] ?? '';
    }

    public function setLastModifiedBy(string $by): static
    {
        return $this->set('lastModifiedBy', $by);
    }

    public function setDescription(string $description): static
    {
        return $this->set('description', $description);
    }

    public function getDescription(): string
    {
        return $this->props['description'] ?? '';
    }

    public function setKeywords(string $keywords): static
    {
        return $this->set('keywords', $keywords);
    }

    public function setCategory(string $category): static
    {
        return $this->set('category', $category);
    }

    public function setCompany(string $company): static
    {
        return $this->set('company', $company);
    }

    public function setManager(string $manager): static
    {
        $this->manager = $manager; // PHP-side only, see class doc

        return $this;
    }

    public function getManager(): string
    {
        return $this->manager;
    }

    public function setCreated(null|float|int|string $timestamp): static
    {
        $this->created = $this->coerceTimestamp($timestamp);

        return $this->set('created', \gmdate('Y-m-d\TH:i:s\Z', $this->created));
    }

    public function getCreated(): int
    {
        return $this->created;
    }

    public function setModified(null|float|int|string $timestamp): static
    {
        $this->modified = $this->coerceTimestamp($timestamp);

        return $this->set('modified', \gmdate('Y-m-d\TH:i:s\Z', $this->modified));
    }

    public function getModified(): int
    {
        return $this->modified;
    }

    // --- custom properties ---------------------------------------------------

    public function setCustomProperty(string $name, mixed $value = '', ?string $type = null): static
    {
        if ($type === null || !\in_array($type, self::VALID_PROPERTY_TYPE_LIST, true)) {
            $type = self::identifyPropertyType($value);
        }
        $this->customProperties[$name] = ['value' => $value, 'type' => $type];

        $payload = ['name' => $name, 'type' => $type];
        if ($type === self::PROPERTY_TYPE_DATE) {
            $payload['value'] = \gmdate('Y-m-d\TH:i:s\Z', $this->coerceTimestamp($value));
        } else {
            $payload['value'] = $value;
        }
        Native::customProp($this->spreadsheet->getHandle(), $payload);

        return $this;
    }

    /** @return list<string> */
    public function getCustomProperties(): array
    {
        return \array_keys($this->customProperties);
    }

    public function isCustomPropertySet(string $name): bool
    {
        return isset($this->customProperties[$name]);
    }

    public function getCustomPropertyValue(string $name): mixed
    {
        return $this->customProperties[$name]['value'] ?? null;
    }

    public function getCustomPropertyType(string $name): ?string
    {
        return $this->customProperties[$name]['type'] ?? null;
    }

    public function removeCustomProperty(string $name): static
    {
        if (isset($this->customProperties[$name])) {
            unset($this->customProperties[$name]);
            Native::customProp($this->spreadsheet->getHandle(), ['name' => $name, 'remove' => true]);
        }

        return $this;
    }

    public static function getTypeForValue(mixed $value): string
    {
        return self::identifyPropertyType($value);
    }

    private static function identifyPropertyType(mixed $value): string
    {
        return match (true) {
            \is_bool($value) => self::PROPERTY_TYPE_BOOLEAN,
            \is_int($value) => self::PROPERTY_TYPE_INTEGER,
            \is_float($value) => self::PROPERTY_TYPE_FLOAT,
            $value === null => self::PROPERTY_TYPE_STRING,
            default => self::PROPERTY_TYPE_STRING,
        };
    }

    private function coerceTimestamp(null|float|int|string $timestamp): int
    {
        if ($timestamp === null) {
            return \time();
        }
        if (\is_string($timestamp)) {
            return \is_numeric($timestamp) ? (int) $timestamp : (int) \strtotime($timestamp);
        }

        return (int) $timestamp;
    }

    private function set(string $key, string $value): static
    {
        $this->props[$key] = $value;
        Native::docProps($this->spreadsheet->getHandle(), $this->props);

        return $this;
    }
}
