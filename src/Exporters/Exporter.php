<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Writer\Common\Creator\WriterFactory;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\AbstractExcel;
use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\Traits\Macroable;
use Illuminate\Support\Facades\Storage;
use League\Flysystem\Filesystem;
use Illuminate\Contracts\Filesystem\Filesystem as LaravelFilesystem;

class Exporter extends AbstractExcel implements Contracts\Exporter
{
    use Macroable, WriteSheet;

    /**
     * @var Filesystem
     */
    protected $filesystem;

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
     * @param array|\Closure|\Generator|ChunkReading $data
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
     * 分块导入
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
        return $this->data(new ChunkReading($callbacks));
    }

    /**
     * 下载导出文件
     *
     * @param string|null $fileName
     * @return void
     */
    public function download(string $fileName)
    {
        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer = $this->makeWriter($fileName);

        $writer->openToBrowser($fileName);

        $this->writeSheets($writer);

        $writer->close();
    }

    /**
     * 存储导出文件
     *
     * @param string $filePath
     * @param array $diskConfig
     * @return bool
     */
    public function store(string $filePath, array $diskConfig = [])
    {
        if (! $this->filesystem) {
            /* @var \Box\Spout\Writer\WriterInterface $writer */
            $writer = $this->makeWriter($filePath);

            $writer->openToFile($filePath);

            $this->writeSheets($writer);

            $writer->close();

            return true;
        }

        $extension = pathinfo($filePath)['extension'] ?? null;

        if (empty($this->type)) {
            $this->type($extension);
        }

        return $this->filesystem->put($filePath, $this->raw(), $diskConfig);

    }

    /**
     * @param Filesystem|LaravelFilesystem|string $filesystem
     * @return $this
     */
    public function disk($filesystem)
    {
        if (is_string($filesystem)) {
            $filesystem = Storage::disk($filesystem);
        }

        if ($filesystem instanceof LaravelFilesystem) {
            $filesystem = $filesystem->getDriver();
        }

        $this->filesystem = $filesystem;

        return $this;
    }


    /**
     * @param string|null $type
     * @return string
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function raw(string $type = null)
    {
        $type && $this->type($type);

        ob_start();

        /* @var \Box\Spout\Writer\WriterInterface $writer */
        $writer = $this->makeWriter();

        $writer->openToBrowser('excel');

        $this->writeSheets($writer);

        $writer->close();

        $content = ob_get_clean();

        header_remove();

        return $content;
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

}
