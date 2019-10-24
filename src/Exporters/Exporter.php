<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\Traits\Macroable;
use Dcat\EasyExcel\Traits\Excel;
use Dcat\EasyExcel\Spout\WriterFactory;
use Box\Spout\Writer\Common\Creator\WriterFactory as SpoutWriterFactory;

/**
 * @method $this xlsx()
 * @method $this csv()
 * @method $this ods()
 */
class Exporter implements Contracts\Exporter
{
    use Macroable, Excel, WriteSheet;

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
     * @param array|\Closure|\Generator|ChunkingQuery $data
     * @return $this
     */
    public function data($data)
    {
        if (is_scalar($data)) {
            throw new \InvalidArgumentException('Not support scalar variable.');
        }

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
     * 分批次导出数据
     *
     * e.g:
     *
     * $this->chunk(function (int $times) {
     *     $size = 100;
     *
     *     return User::query()->forPage($times, $size)->toArray();
     * });
     *
     * $this->chunk([
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
    public function chunk($callbacks)
    {
        return $this->data(new ChunkingQuery($callbacks));
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
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer = $this->makeWriter(null, WriterFactory::class);

        ob_start();

        $writer->openToOutput();

        $this->writeSheets($writer)->close();

        return ob_get_clean();
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
     * @param string $factory
     * @return WriterInterface
     */
    protected function makeWriter(?string $path = null, string $factory = null)
    {
        $factory = $factory ?: SpoutWriterFactory::class;

        /* @var WriterInterface $writer */
        if ($this->type) {
            $writer = $factory::createFromType($this->type);
        } else {
            $writer = $factory::createFromFile($path);
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
