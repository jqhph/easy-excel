<?php

namespace Dcat\EasyExcel\Contracts;

use Dcat\EasyExcel\Support\SheetCollection;
use Symfony\Component\HttpFoundation\File\UploadedFile;

interface Importer
{
    /**
     * @param string|UploadedFile $filePath
     * @return Sheets
     */
    public function import($filePath);

    /**
     * @param callable $callback
     * @return $this
     */
    public function each(callable $callback);

    /**
     * 获取第一个sheet
     *
     * @return Sheet
     */
    public function first();

    /**
     * 获取当前打开的sheet
     *
     * @return Sheet
     */
    public function working();

    /**
     * 根据名称或序号获取sheet
     *
     * @param int|string $indexOrName
     * @return Sheet
     */
    public function index($indexOrName);

    /**
     * @return array
     */
    public function toArray(): array;

    /**
     * @return SheetCollection
     */
    public function toCollection(): SheetCollection;

}
