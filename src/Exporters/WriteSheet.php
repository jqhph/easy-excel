<?php

namespace Dcat\EasyExcel\Exporters;

use Box\Spout\Common\Entity\Row;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\CSV\Writer as CsvWriter;
use Box\Spout\Writer\WriterInterface;
use Dcat\EasyExcel\Support\SheetCollection;
use Generator;
use Dcat\EasyExcel\Support\Arr;

/**
 * @mixin \Dcat\EasyExcel\Contracts\Exporter
 */
trait WriteSheet
{
    protected $writedHeaders = [];

    protected function writeSheets(WriterInterface $writer)
    {
        $data    = $this->makeSheetsArray();
        $keys    = array_keys($data);
        $lastKey = end($keys);

        foreach ($data as $index => $collection) {
            if ($collection instanceof \Generator) {
                $this->writeRowsFromGenerator($writer, $index, $collection);
            } else {
                $collection = $this->convertToArray($collection);

                $this->writeRowsFromArray($writer, $index, $collection);
            }

            if (is_string($index) && method_exists($writer, 'getCurrentSheet')) {
                $writer->getCurrentSheet()->setName($index);
            }

            if ($lastKey !== $index && method_exists($writer, 'addNewSheetAndMakeItCurrent')) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }

        return $writer;
    }

    protected function writeRowsFromArray(WriterInterface $writer, $index, array &$collection)
    {
        // Add header row.
        if ($this->canWriteHeaders($writer, $index)) {
            $this->writeHeaders($writer, current($collection));
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
                if ($this->canWriteHeaders($writer, $index)) {
                    $this->writeHeaders($writer, $item);
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

    protected function writeHeaders(WriterInterface $writer, array $firstRow)
    {
        $writer->addRow(
            $this->makeDefaultRow(
                $this->headers ?: array_keys($firstRow),
                $this->headerStyle
            )
        );
    }

    protected function makeDefaultRow(array $item, ?Style $style = null)
    {
        if ($style) {
            return WriterEntityFactory::createRowFromArray($item, $style);
        }

        return WriterEntityFactory::createRowFromArray($item);
    }

    /**
     * @param WriterInterface $writer
     * @param $index
     * @return bool
     */
    protected function canWriteHeaders(WriterInterface $writer, $index): bool
    {
        if (
            $this->headers === false
            || ! empty($this->writedHeaders[$index])
            || ($this->writedHeaders && $writer instanceof CsvWriter)
        ) {
            return false;
        }

        $this->writedHeaders[$index] = true;

        return true;
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
     * @return Generator[]|array
     */
    protected function makeSheetsArray(): array
    {
        $data = $this->data;

        if ($data instanceof ChunkingQuery) {
            return $data->makeGenerators();
        }

        if ($this->data instanceof \Closure) {
            $data = $data($this);
        }

        if ($data instanceof SheetCollection) {
            $data = $data->toArray();
        }

        if (
            (is_array($data) && ! Arr::isAssoc($data))
            || $data instanceof Generator
        ) {
            return [&$data];
        }

        return (array) $data;
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
