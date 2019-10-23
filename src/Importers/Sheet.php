<?php

namespace Dcat\EasyExcel\Importers;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\SheetInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Illuminate\Support\Arr;
use Dcat\EasyExcel\Contracts;

class Sheet implements Contracts\Sheet
{
    /**
     * @var Importer
     */
    protected $importer;

    /**
     * @var SheetInterface
     */
    protected $sheet;

    public function __construct(Importer $importer, SheetInterface $sheet)
    {
        $this->importer = $importer;
        $this->sheet    = $sheet;
    }

    /**
     * @return bool
     */
    public function valid(): bool
    {
        return $this->sheet ? true : false;
    }

    /**
     * @return int
     */
    public function getIndex()
    {
        return $this->sheet->getIndex();
    }

    /**
     * @return string
     */
    public function getName()
    {
        return $this->sheet->getName();
    }

    /**
     * @return bool
     */
    public function isActive()
    {
        return $this->sheet->isActive();
    }

    /**
     * @return bool
     */
    public function isVisible()
    {
        return $this->sheet->isVisible();
    }

    /**
     * @return SheetInterface
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * 逐行读取
     *
     * @param callable|null $callback
     * @return $this
     */
    public function each(callable $callback)
    {
        $headers         = [];
        $originalHeaders = [];

        foreach ($this->sheet->getRowIterator() as $k => $row) {
            $row = $row instanceof Row ? $row->toArray() : [];

            if ($this->importer->getHeaders() !== false && $k == 1) {
                $originalHeaders = $row;
                $headers         = $this->formatHeaders($row);

                continue;
            }

            $row = $this->formatRow($row, $headers);

            call_user_func($callback, $row, $k, $originalHeaders);
        }
    }

    /**
     * 分块读取
     *
     * @param int $size
     * @param callable $callback
     * @return $this
     */
    public function chunk(int $size, callable $callback)
    {
        $chunkData = [];

        $this->each(function (array $row) use ($size, &$chunkData, &$callback) {
            $chunkData[] = $row;

            if (count($chunkData) >= $size) {
                call_user_func($callback, new SheetCollection($chunkData));

                $chunkData = [];
            }
        });

        if ($chunkData) {
            call_user_func($callback, new SheetCollection($chunkData));
        }

        unset($chunkData);

        return $this;
    }

    /**
     * @return array
     */
    public function toArray(): array
    {
        $array = [];

        $this->each(function (array $row, $k) use (&$array) {
            $array[$k] = $row;
        });

        return $array;
    }

    /**
     * @return SheetCollection
     */
    public function collect(): SheetCollection
    {
        return new SheetCollection($this->toArray());
    }

    /**
     * @param array $row
     * @param array $headers
     * @return array
     */
    protected function formatRow(array &$row, array $headers)
    {
        if ($this->importer->getHeaders() === false) {
            return $row;
        }

        $countHeaders = count($headers);

        $countRow = count($row);

        if ($countHeaders > $countRow) {
            $row = array_merge($row, array_fill(0, $countHeaders - $countRow, null));

        } elseif ($countHeaders < $countRow) {
            $row = array_slice($row, 0, $countHeaders);

        }

        return array_combine($headers, $row);
    }

    /**
     * @param array $row
     * @return array|false|mixed
     */
    protected function formatHeaders(&$row)
    {
        if ($headers = $this->importer->getHeaders()) {
            return Arr::isAssoc($headers) ? array_keys($headers) : $headers;
        }

        return $this->toStrings($row);
    }

    /**
     * @param array $values
     * @return array
     */
    protected function toStrings(array $values)
    {
        foreach ($values as &$value) {
            if ($value instanceof \Datetime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string)$value;
            }
        }

        return $values;
    }

}
