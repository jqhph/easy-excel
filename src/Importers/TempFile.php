<?php

namespace Dcat\EasyExcel\Importers;

use Illuminate\Support\Str;
use League\Flysystem\FileNotFoundException;
use League\Flysystem\FilesystemInterface;

trait TempFile
{
    protected $tempFolder;

    protected $tempFile;

    /**
     * @param FilesystemInterface $filesystem
     * @param string $filePath
     * @return string
     * @throws \League\Flysystem\FileNotFoundException
     */
    public function moveFileToTemp(FilesystemInterface $filesystem, string $filePath)
    {
        $this->tempFile = $newPath = $this->generateTempPath($filePath);

        file_put_contents($newPath, $filesystem->read($filePath));

        return $newPath;
    }

    protected function removeTempFile()
    {
        if ($this->tempFile && is_file($this->tempFile)) {
            @unlink($this->tempFile);
        }
    }

    /**
     * @param string $filePath
     * @return string
     */
    private function generateTempPath(string $filePath)
    {
        $extension = pathinfo($filePath)['extension'] ?? null;

        return $this->getTempFolder()
            .'/'
            .uniqid(microtime(true).static::generateRandom())
            .($extension ? ".{$extension}" : '');
    }

    /**
     * @return string
     */
    private function getTempFolder()
    {
        return sys_get_temp_dir();
    }
}
