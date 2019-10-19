<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Generator;
use Illuminate\Support\Arr;
use Box\Spout\Writer\XLSX\Writer as XLSXWriter;
use Box\Spout\Writer\ODS\Writer as ODSWriter;

trait WriteSheet
{
    protected $writedHeaders = [];

    protected function writeSheets(WriterInterface $writer)
    {
        $data = $this->makeSheetsArray();

        $keys    = array_keys($data);
        $lastKey = end($keys);

        $hasSheets = ($writer instanceof XLSXWriter || $writer instanceof ODSWriter);

        foreach ($data as $index => $collection) {
            if ($collection instanceof \Generator) {
                $this->writeRowsFromGenerator($writer, $index, $collection);
            } else {
                $collection = $this->convertToArray($collection);

                $this->writeRowsFromArray($writer, $index, $collection);
            }

            if (is_string($index)) {
                $writer->getCurrentSheet()->setName($index);
            }

            if ($hasSheets && $lastKey !== $index) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }

        return $writer;
    }

    protected function writeRowsFromArray(WriterInterface $writer, $index, array &$collection)
    {
        // Add header row.
        if (empty($this->writedHeaders[$index]) && $this->headers !== false) {
            $this->writeHeaders($writer, current($collection));

            $this->writedHeaders[$index] = true;
        }

        foreach ($collection as $item) {
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

            foreach ($items as &$item) {
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
        $item = $this->filterAndSort($item);

        $item = $this->transformRow($item);

        if ($this->rowCallback) {
            $item = call_user_func($this->rowCallback, $item);
        }

        if (is_array($item)) {
            $item = $this->makeDefaultRow($item);
        }

        // Write rows (one by one).
        $writer->addRow($item);
    }

    protected function makeDefaultRow(array $item)
    {
        return WriterEntityFactory::createRowFromArray($item);
    }

    protected function writeHeaders(WriterInterface $writer, array $firstRow)
    {
        if ($this->headers) {
            $keys = $this->headers;
        } else {
            $keys = array_keys($firstRow);
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
    public function filterAndSort($row)
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
    protected function makeSheetsArray(): array
    {
        $data = $this->data;

        if ($data instanceof GeneratorFactory) {
            return $data->make();
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
     * @param $data
     * @return array
     */
    protected function transformRow(array $data)
    {
        $strings = [];

        foreach ($data as &$value) {
            $value = is_int($value) || is_float($value) || is_null($value) ? (string)$value : $value;

            if (is_string($value)) {
                $strings[] = $value;
            }
        }

        return $strings;
    }

}
