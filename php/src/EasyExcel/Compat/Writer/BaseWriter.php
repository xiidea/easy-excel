<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;


/**
 * Shared base for writers, PhpSpreadsheet-compatible. Holds the include-charts,
 * pre-calculate and disk-caching state plus the file-handle plumbing so custom
 * writers (`class MyWriter extends BaseWriter`) only have to implement save().
 *
 * Subclasses must declare the IWriter constructor (taking a Spreadsheet); this
 * class deliberately leaves construction to them so each writer can capture the
 * workbook in whatever shape it needs.
 */
abstract class BaseWriter implements IWriter
{
    protected bool $includeCharts = false;

    protected bool $preCalculateFormulas = true;

    /**
     * Whether setPreCalculateFormulas() was called explicitly.
     *
     * The accessor keeps PhpSpreadsheet's `true` default for API parity, but
     * easy-excel must not *act* on it unhinted: pre-calculation reads every
     * formula back, which forces a streamed workbook into the full in-memory
     * model. Trading away constant-memory streaming — the reason this engine
     * exists — is a decision the caller has to make, not one inherited from an
     * upstream default (review blocker 3).
     */
    protected bool $preCalculateFormulasSet = false;

    private bool $useDiskCaching = false;

    private string $diskCachingDirectory = './';

    /** @var resource */
    protected $fileHandle;

    private bool $shouldCloseFile = false;

    public function getIncludeCharts(): bool
    {
        return $this->includeCharts;
    }

    public function setIncludeCharts(bool $includeCharts): self
    {
        $this->includeCharts = $includeCharts;

        return $this;
    }

    public function getPreCalculateFormulas(): bool
    {
        return $this->preCalculateFormulas;
    }

    public function setPreCalculateFormulas(bool $precalculateFormulas): self
    {
        $this->preCalculateFormulas = $precalculateFormulas;
        $this->preCalculateFormulasSet = true;

        return $this;
    }

    /**
     * @internal the effective setting: only an explicit opt-in enables the
     *           save-time evaluate-and-cache pass (see $preCalculateFormulasSet)
     */
    public function shouldPreCalculateFormulas(): bool
    {
        return $this->preCalculateFormulasSet && $this->preCalculateFormulas;
    }

    public function getUseDiskCaching(): bool
    {
        return $this->useDiskCaching;
    }

    public function setUseDiskCaching(bool $useDiskCache, ?string $cacheDirectory = null): self
    {
        $this->useDiskCaching = $useDiskCache;

        if ($cacheDirectory !== null) {
            if (\is_dir($cacheDirectory)) {
                $this->diskCachingDirectory = $cacheDirectory;
            } else {
                throw new Exception("Directory does not exist: $cacheDirectory");
            }
        }

        return $this;
    }

    public function getDiskCachingDirectory(): string
    {
        return $this->diskCachingDirectory;
    }

    /** Apply the save()-time $flags onto this writer's state. */
    protected function processFlags(int $flags): void
    {
        if (($flags & self::SAVE_WITH_CHARTS) !== 0) {
            $this->setIncludeCharts(true);
        }
        if (($flags & self::DISABLE_PRECALCULATE_FORMULAE) !== 0) {
            $this->setPreCalculateFormulas(false);
        }
    }

    /**
     * Open the destination as a writable handle. A resource is used as-is and
     * left open for the caller; a string is fopen'd and closed by
     * {@see maybeCloseFileHandle()}.
     *
     * @param resource|string $filename
     */
    public function openFileHandle($filename): void
    {
        if (!\is_string($filename)) {
            $this->fileHandle = $filename;
            $this->shouldCloseFile = false;

            return;
        }

        $fileHandle = $filename !== '' ? \fopen($filename, 'wb') : false;
        if ($fileHandle === false) {
            throw new Exception('Could not open file "' . $filename . '" for writing.');
        }

        $this->fileHandle = $fileHandle;
        $this->shouldCloseFile = true;
    }

    protected function tryClose(): bool
    {
        return \fclose($this->fileHandle);
    }

    /** Close the handle only if we opened it ourselves. */
    protected function maybeCloseFileHandle(): void
    {
        if ($this->shouldCloseFile && !$this->tryClose()) {
            throw new Exception('Could not close file after writing.');
        }
    }
}
