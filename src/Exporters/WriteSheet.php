<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Generator;
use Illuminate\Support\Arr;

trait WriteSheet
{
    protected $writedHeaders = [];

    protected function writeSheets(WriterInterface $writer)
    {
        $data = $this->prepareData();

        $keys    = array_keys($data);
        $lastKey = end($keys);

        $hasSheets = ($writer instanceof \Box\Spout\Writer\XLSX\Writer || $writer instanceof \Box\Spout\Writer\ODS\Writer);

        foreach ($data as $index => $collection) {
            if (is_array($collection)) {
                $collection = new SheetCollection($collection);
            }

            if ($collection instanceof SheetCollection) {
                $this->writeRowsFromCollection($writer, $index, $collection);
            } elseif ($collection instanceof \Generator) {
                $this->writeRowsFromGenerator($writer, $index, $collection);
            }

            if (is_string($index)) {
                $writer->getCurrentSheet()->setName($index);
            }

            if ($hasSheets && $lastKey !== $index) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
    }

    protected function writeRowsFromCollection(WriterInterface $writer, $index, SheetCollection $collection)
    {
        // Add header row.
        if (empty($this->writedHeaders[$index]) && $this->headers !== false) {
            $this->writeHeaders($writer, $collection->first());

            $this->writedHeaders[$index] = true;
        }

        foreach ($collection->toArray() as $item) {
            $this->writeRow($writer, $item);
        }
    }

    protected function writeRowsFromGenerator(WriterInterface $writer, $index, Generator $generator)
    {
        foreach ($generator as $key => $items) {
            $items = $this->convertToArray($items);

            if (! is_array(current($items))) {
                $items = [$items];
            }

            foreach ($items as $item) {
                // Add header row.
                if (empty($this->writedHeaders[$index]) && $this->headers !== false) {
                    $this->writeHeaders($writer, $item);

                    $this->writedHeaders[$index] = true;
                }

                $this->writeRow($writer, $item);
            }
        }
    }

    /**
     * @param $writer
     * @param array $item
     * @param callable|null $callback
     */
    protected function writeRow(WriterInterface $writer, array $item)
    {
        $item = $this->filterRow($item);

        // Prepare row (i.e remove non-string)
        $item = $this->transformRow($item);

        $item = $this->convertToArray($item);

        if ($this->rowCallback) {
            call_user_func($this->rowCallback, $item);
        } else {
            $item = $this->makeDefaultRow($item);
        }

        // Write rows (one by one).
        $writer->addRow($item);
    }

    protected function makeDefaultRow(array $item)
    {
        return WriterEntityFactory::createRowFromArray($item);
    }

    protected function writeHeaders(WriterInterface $writer, $firstRow)
    {
        if ($this->headers) {
            $keys = $this->headers;
        } else {
            $keys = array_keys($this->convertToArray($firstRow));
        }

        if ($callback = $this->headerCallback) {
            $keys = $callback($keys);
        } else {
            $keys = $this->makeDefaultRow($keys);
        }

        $writer->addRow($keys);
    }

    /**
     * @param array|SheetCollection $row
     * @return array
     */
    public function filterRow($row)
    {
        if (! $this->headers) {
            return $row;
        }

        $row = $this->convertToArray($row);

        $newRow = [];
        foreach ($this->headers as $key => $label) {
            $newRow[$key] = $row[$key] ?? null;
        }

        return $newRow;
    }

    /**
     * @param mixed $value
     * @return array
     */
    protected function convertToArray($value)
    {
        if (is_array($value)) {
            return $value;
        }

        if ($value instanceof SheetCollection || method_exists($value, 'toArray')) {
            return $value->toArray();
        }

        return (array) $value;
    }

    /**
     * @return \Generator[]|array
     */
    protected function prepareData()
    {
        $data = $this->data;

        if ($data instanceof ChunkReading) {
            return $data->makeGenerators();
        }

        if ($this->data instanceof \Closure) {
            $data = $data($this);
        }

        if ($data instanceof SheetCollection) {
            $data = $data->toArray();
        }

        if (is_array($data) && ! Arr::isAssoc($data)) {
            return [new SheetCollection($data)];
        }

        return [$data];
    }

    /**
     * @param SheetCollection $collection
     */
    protected function transform(SheetCollection $collection)
    {
        $collection->transform(function ($data) {
            return $this->transformRow($data);
        });
    }

    /**
     * @param $data
     * @return array
     */
    protected function transformRow($data)
    {
        return (new SheetCollection($data))->map(function ($value) {
            return is_int($value) || is_float($value) || is_null($value) ? (string)$value : $value;
        })->filter(function ($value) {
            return is_string($value);
        })->toArray();
    }

}
