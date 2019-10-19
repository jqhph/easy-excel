<?php

namespace Dcat\EasyExcel;

use Box\Spout\Common\Type;
use Dcat\EasyExcel\Exporters\Exporter;
use Dcat\EasyExcel\Importers\Importer;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Dcat\EasyExcel\Contracts;

class Excel
{
    const XLSX = Type::XLSX;
    const CSV  = Type::CSV;
    const ODS  = Type::ODS;

    /**
     * 导入
     *
     * @param string|UploadedFile $filePath
     * @return Contracts\Importer
     */
    public static function import($filePath): Contracts\Importer
    {
        return new Importer($filePath);
    }

    /**
     * 导出
     *
     * @param array|\Closure|\Generator $data
     * @return Contracts\Exporter
     */
    public static function export($data = null): Contracts\Exporter
    {
        return new Exporter($data);
    }

    /**
     * @param mixed ...$params
     * @return Factory
     */
    public static function xlsx(...$params)
    {
        return new Factory(static::XLSX, $params);
    }

    /**
     * @param mixed ...$params
     * @return Factory
     */
    public static function csv(...$params)
    {
        return new Factory(static::CSV, $params);
    }

    /**
     * @param mixed ...$params
     * @return Factory
     */
    public static function ods(...$params)
    {
        return new Factory(static::ODS, $params);
    }
}
