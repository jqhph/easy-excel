<?php

namespace Dcat\EasyExcel\Importers;

use Box\Spout\Reader\Common\Creator\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Dcat\EasyExcel\AbstractExcel;
use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\SheetCollection;
use Dcat\EasyExcel\Support\Traits\Macroable;
use Symfony\Component\HttpFoundation\File\UploadedFile;

class Importer extends AbstractExcel implements Contracts\Importer
{
    use Macroable;

    /**
     * @var string|UploadedFile
     */
    protected $filePath;

    public function __construct($filePath)
    {
        $this->filePath = $filePath;
    }

    /**
     * @param string|UploadedFile $filePath
     * @return Contracts\Sheets
     */
    public function import($filePath)
    {
        $reader = $this->makeReader($filePath);

        return new LazySheets($this->readSheets($reader));
    }

    /**
     * @return array
     */
    public function toArray(): array
    {
        return $this->import($this->filePath)->toArray();
    }

    /**
     * @return SheetCollection
     */
    public function toCollection(): SheetCollection
    {
        return $this->import($this->filePath)->toCollection();
    }

    /**
     * @param callable $callback
     * @return $this
     */
    public function each(callable $callback)
    {
        $this->import($this->filePath)->each($callback);

        return $this;
    }

    /**
     * 获取第一个sheet
     *
     * @return Contracts\Sheet
     */
    public function first()
    {
        $sheet = null;

        $this->import($this->filePath)->each(function (Sheet $value) use (&$sheet) {
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
    public function working()
    {
        $sheet = null;

        $this->import($this->filePath)->each(function (Sheet $value) use (&$sheet) {
            if ($value->isWorking()) {
                $sheet = $value;

                return false;
            }

        });

        return $sheet ?: $this->makeNullSheet();
    }

    /**
     * 根据名称或序号获取sheet
     *
     * @param int|string $indexOrName
     * @return Contracts\Sheet
     */
    public function sheet($indexOrName)
    {
        return $this->import($this->filePath)->get($indexOrName) ?: $this->makeNullSheet();
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
