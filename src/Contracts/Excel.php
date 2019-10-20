<?php

namespace Dcat\EasyExcel\Contracts;

use League\Flysystem\FilesystemInterface;
use Illuminate\Contracts\Filesystem\Filesystem as LaravelFilesystem;

/**
 * @method $this xlsx()
 * @method $this csv()
 * @method $this ods()
 */
interface Excel
{
    /**
     * @param \Closure $callback
     * @return $this
     */
    public function option(\Closure $callback);

    /**
     * @param array|\Closure $headers
     * @return $this
     */
    public function headers($headers);

    /**
     * @return array|false
     */
    public function getHeaders();

    /**
     * @return $this
     */
    public function withoutHeaders();

    /**
     * @param string $type
     * @return $this
     */
    public function type(string $type);

    /**
     * @return string|null
     */
    public function getType();

    /**
     * @param FilesystemInterface|LaravelFilesystem|string $filesystem
     * @return $this
     */
    public function disk($filesystem);

}
