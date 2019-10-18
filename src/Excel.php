<?php

namespace Dcat\EasyExcel;

use Dcat\EasyExcel\Exporters\Exporter;
use Dcat\EasyExcel\Importers\Importer;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Dcat\EasyExcel\Contracts;

class Excel
{
    const XLSX = 'xlsx';
    const CSV  = 'csv';
    const ODS  = 'ods';

    /**
     * 导入
     *
     * @param string|UploadedFile $filePath
     * @return Contracts\Importer|Contracts\Excel
     */
    public static function import($filePath)
    {
        return new Importer($filePath);
    }

    /**
     * 导出
     *
     * @param array|\Closure|\Generator $data
     * @return Contracts\Exporter|Contracts\Excel
     */
    public static function export($data = null)
    {
        return new Exporter($data);
    }
}
