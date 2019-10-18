<?php

namespace Dcat\EasyExcel\Contracts;

use League\Flysystem\Filesystem;
use Illuminate\Contracts\Filesystem\Filesystem as LaravelFilesystem;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

interface Exporter
{
    /**
     * 设置导出数据
     *
     * @param array|\Closure|\Generator $data
     * @return $this
     */
    public function data($data);

    /**
     * @param callable $callback
     * @return $this
     */
    public function row(callable $callback);

    /**
     * 分块导入
     *
     * @param callable|callable[] $callbacks
     * @return $this
     */
    public function chunk($callbacks);

    /**
     * 下载导出文件
     *
     * @param string|null $fileName
     *
     * @return BinaryFileResponse
     */
    public function download(string $fileName);

    /**
     * 存储导出文件
     *
     * @param string $filePath
     * @param array $diskConfig
     * @return bool
     */
    public function store(string $filePath, array $diskConfig = []);

    /**
     * @param Filesystem|LaravelFilesystem|string $filesystem
     * @return $this
     */
    public function disk($filesystem);

    /**
     * @param string|null $type
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function raw(string $type = null);

}
