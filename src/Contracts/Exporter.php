<?php

namespace Dcat\EasyExcel\Contracts;

use Dcat\EasyExcel\Exporters\ChunkingQuery;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

interface Exporter extends Excel
{
    /**
     * 设置导出数据
     *
     * @param array|\Closure|\Generator|ChunkingQuery $data
     * @return $this
     */
    public function data($data);

    /**
     * @param callable $callback
     * @return $this
     */
    public function row(callable $callback);

    /**
     * 分批次导入数据
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
     * @return mixed
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
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function raw();

}
