<?php

namespace Dcat\EasyExcel\Importers;

use Box\Spout\Reader\Common\Creator\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\SheetCollection;
use Dcat\EasyExcel\Support\Traits\Macroable;
use Dcat\EasyExcel\Traits\Excel;
use Symfony\Component\HttpFoundation\File\UploadedFile;

/**
 * @method $this xlsx()
 * @method $this csv()
 * @method $this ods()
 */
class Importer implements Contracts\Importer
{
    use Macroable, Excel;

    /**
     * @var string|UploadedFile
     */
    protected $filePath;

    public function __construct($filePath)
    {
        $this->file($filePath);
    }

    /**
     * @param string|UploadedFile $filePath
     * @return $this
     */
    public function file($filePath)
    {
        $this->filePath = $filePath;

        return $this;
    }

    /**
     * @return Contracts\Sheets
     */
    public function sheets()
    {
        $filePath = $this->prepareFileName($this->filePath);

        $reader = $this->makeReader($filePath);

        return new LazySheets($this->readSheets($reader));
    }

    /**
     * 根据名称或序号获取sheet
     *
     * @param int|string $indexOrName
     * @return Contracts\Sheet
     */
    public function sheet($indexOrName): Contracts\Sheet
    {
        return $this->sheets()->index($indexOrName) ?: $this->makeNullSheet();
    }


    /**
     * @return array
     */
    public function toArray(): array
    {
        return $this->sheets()->toArray();
    }

    /**
     * @return SheetCollection
     */
    public function collect(): SheetCollection
    {
        return $this->sheets()->collect();
    }

    /**
     * @param callable $callback
     * @return $this
     */
    public function each(callable $callback)
    {
        $this->sheets()->each($callback);

        return $this;
    }

    /**
     * 获取第一个sheet
     *
     * @return Contracts\Sheet
     */
    public function first(): Contracts\Sheet
    {
        $sheet = null;

        $this->sheets()->each(function (Sheet $value) use (&$sheet) {
            $sheet = $value;

            return false;
        });

        return $sheet ?: $this->makeNullSheet();
    }

    /**
     * 获取当前打开的sheet
     *
     * @return Contracts\Sheet
     */
    public function working(): Contracts\Sheet
    {
        $sheet = null;

        $this->sheets()->each(function (Sheet $value) use (&$sheet) {
            if ($value->isWorking()) {
                $sheet = $value;

                return false;
            }

        });

        return $sheet ?: $this->makeNullSheet();
    }

    /**
     * @param \Box\Spout\Reader\ReaderInterface $reader
     * @return \Generator
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    protected function readSheets(ReaderInterface $reader)
    {
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            yield new Sheet($this, $sheet);
        }

        $reader->close();
    }

    /**
     * @param string|UploadedFile $path
     *
     * @return \Box\Spout\Reader\ReaderInterface
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Common\Exception\IOException
     */
    protected function makeReader($path)
    {
        $extension = null;
        if ($path instanceof UploadedFile) {
            $extension = $path->guessClientExtension();
            $path      = $path->getRealPath();
        }

        /* @var \Box\Spout\Reader\ReaderInterface $reader */
        if ($this->type || $extension) {
            $reader = ReaderFactory::createFromType($this->type ?: $extension);
        } else {
            $reader = ReaderFactory::createFromFile($path);
        }

        $reader->open($path);

        $this->configure($reader);

        return $reader;
    }

    /**
     * @return NullSheet
     */
    protected function makeNullSheet()
    {
        return new NullSheet();
    }

}
