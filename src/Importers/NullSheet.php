<?php

namespace Dcat\EasyExcel\Importers;

use Box\Spout\Reader\SheetInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Dcat\EasyExcel\Contracts;

class NullSheet implements Contracts\Sheet
{
    /**
     * @return bool
     */
    public function valid(): bool
    {
        return false;
    }

    /**
     * @return int
     */
    public function getIndex()
    {
        return null;
    }

    /**
     * @return string
     */
    public function getName()
    {
        return null;
    }

    /**
     * @return bool
     */
    public function isWorking()
    {
        return false;
    }

    /**
     * @return SheetInterface
     */
    public function getSheet()
    {
    }

    /**
     * 逐行读取
     *
     * @param callable|null $callback
     * @return $this
     */
    public function each(callable $callback)
    {
        return $this;
    }

    public function chunk(int $size, callable $callback)
    {
        return $this;
    }

    /**
     * @return array
     */
    public function toArray(): array
    {
        return [];
    }

    /**
     * @return SheetCollection
     */
    public function collect(): SheetCollection
    {
        return new SheetCollection($this->toArray());
    }

}
