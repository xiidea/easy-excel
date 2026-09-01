<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Shared;

/**
 * Filesystem helpers, PhpSpreadsheet style (wave 5.2). Pure PHP; the upload
 * temp-dir override exists because PhpSpreadsheet callers set it before
 * generating chart/image temp files.
 */
final class File
{
    private static bool $useUploadTempDirectory = false;

    /**
     * Route temp files through PHP's upload_tmp_dir instead of the system dir.
     *
     * Stores the flag, not the resolved path: upload_tmp_dir is commonly empty,
     * and caching its value here made the setter silently a no-op on those
     * configurations. The directory is resolved at use time, as upstream does.
     */
    public static function setUseUploadTempDirectory(bool $useUploadTempDir): void
    {
        self::$useUploadTempDirectory = $useUploadTempDir;
    }

    public static function getUseUploadTempDirectory(): bool
    {
        return self::$useUploadTempDirectory;
    }

    /**
     * The directory temp files should be written to.
     *
     * Canonicalised with realpath(): on macOS sys_get_temp_dir() reports
     * /var/... while the file the OS actually creates lives under
     * /private/var/..., so callers comparing the two would not match.
     */
    public static function sysGetTempDir(): string
    {
        $upload = self::$useUploadTempDirectory ? (\ini_get('upload_tmp_dir') ?: '') : '';
        $dir = ($upload !== '' && \is_dir($upload)) ? $upload : \sys_get_temp_dir();

        return \realpath($dir) ?: $dir;
    }

    /** Create a uniquely named temp file in sysGetTempDir(). */
    public static function temporaryFilename(): string
    {
        $filename = \tempnam(self::sysGetTempDir(), 'easyexcel');
        if ($filename === false) {
            throw new \RuntimeException('easy-excel: could not create a temporary file');
        }

        return $filename;
    }

    /** True when the path exists as a regular file. */
    public static function fileExists(string $filename): bool
    {
        return \is_file($filename);
    }

    /** Resolve to an absolute path, or '' when it does not exist. */
    public static function realpath(string $filename): string
    {
        return \realpath($filename) ?: '';
    }

    /** @internal test seam */
    public static function reset(): void
    {
        self::$useUploadTempDirectory = false;
    }
}
