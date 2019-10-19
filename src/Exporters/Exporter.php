<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Writer\Common\Creator\WriterFactory;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\Traits\Macroable;
use Dcat\EasyExcel\Traits\Excel;

/**
 * @method $this xlsx()
 * @method $this csv()
 * @method $this ods()
 */
class Exporter implements Contracts\Exporter
{
    use Macroable,
        Excel,
        WriteSheet;

    /**
     * @var array|\Closure|\Generator
     */
    protected $data;

    /**
     * @var callable
     */
    protected $rowCallback;

    public function __construct($data = null)
    {
        if ($data) {
            $this->data($data);
        }
    }

    /**
     * 设置导出数据
     *
     * e.g:
     *
     * $this->data([
     *     ['id' => 1, 'name' => 'marry'], ...
     * ]);
     *
     * $this->data([
     *     'sheet-name' => [
     *         ['id' => 1, 'name' => 'marry'], ...
     *     ],
     *     ...
     * ]);
     *
     * $this->data(function () {
     *    return [
     *        'sheet-name' => [
     *            ['id' => 1, 'name' => 'marry'], ...
     *        ],
     *        ...
     *    ];
     * });
     *
     * @param array|\Closure|\Generator|GeneratorFactory $data
     * @return $this
     */
    public function data($data)
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param callable $callback
     * @return $this
     */
    public function row(callable $callback)
    {
        $this->rowCallback = $callback;

        return $this;
    }

    /**
     * 分批次生成数据
     *
     * e.g:
     *
     * $this->generate(function (int $times) {
     *     $size = 100;
     *
     *     return User::query()->forPage($times, $size)->toArray();
     * });
     *
     * $this->generate([
     *     "sheet-name" => function (int $times) {
     *         $size = 100;
     *
     *         return User::query()->forPage($times, $size)->toArray();
     *     }
     * ]);
     *
     * @param callable|callable[] $callbacks
     * @return $this
     */
    public function generate($callbacks)
    {
        return $this->data(new GeneratorFactory($callbacks));
    }

    /**
     * 下载导出文件
     *
     * @param string|null $fileName
     * @return void
     */
    public function download(string $fileName)
    {
        try {
            /* @var \Box\Spout\Writer\WriterInterface $writer */
            $writer = $this->makeWriter($fileName);

            $writer->openToBrowser($this->prepareFileName($fileName));

            $this->writeSheets($writer)->close();
        } catch (\Throwable $e) {
            $this->removeHttpHeaders();

            throw $e;
        }
    }

    /**
     * 存储导出文件
     *
     * @param string $filePath
     * @param array $diskConfig
     * @return bool
     *
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function store(string $filePath, array $diskConfig = [])
    {
        $filePath = $this->prepareFileName($filePath);

        if (! ($filesystem = $this->filesystem())) {
            return $this->storeInLocal($filePath);
        }

        if (empty($this->type)) {
            $this->type(pathinfo($filePath)['extension'] ?? null);
        }

        return $filesystem->put($filePath, $this->raw(), $diskConfig);

    }

    /**
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function raw()
    {
        try {
            ob_start();

            /* @var \Box\Spout\Writer\WriterInterface $writer */
            $writer = $this->makeWriter();

            $writer->openToBrowser('excel');

            $this->writeSheets($writer)->close();

            $this->removeHttpHeaders();

            return ob_get_clean();
        } catch (\Throwable $e) {
            $this->removeHttpHeaders();

            throw $e;
        }
    }

    /**
     * @param string $filePath
     * @return bool
     * @throws \Box\Spout\Common\Exception\IOException
     */
    protected function storeInLocal(string $filePath)
    {
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer = $this->makeWriter($filePath);

        $writer->openToFile($filePath);

        $this->writeSheets($writer)->close();

        return true;
    }

    /**
     * @param string $path
     * @return WriterInterface
     */
    protected function makeWriter(?string $path = null)
    {
        /* @var WriterInterface $writer */
        if ($this->type) {
            $writer = WriterFactory::createFromType($this->type);
        } else {
            $writer = WriterFactory::createFromFile($path);
        }

        $this->configure($writer);

        return $writer;
    }

    protected function removeHttpHeaders()
    {
        if (! headers_sent()) {
            header_remove();
        }
    }

}
